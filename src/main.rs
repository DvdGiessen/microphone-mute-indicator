// Simple microphone mute indicator for Windows
// Author: Daniël van de Giessen <daniel@dvdgiessen.nl>

#![windows_subsystem = "windows"]

use argh::FromArgs;
use std::{cell::RefCell, ffi::c_void, path::PathBuf};
use strum_macros::EnumString;
use windows::{
    core::*,
    Win32::{
        Devices::FunctionDiscovery::PKEY_Device_FriendlyName,
        Foundation::*,
        Graphics::{Gdi, Gdi::*},
        Media::Audio::{Endpoints::*, *},
        System::{
            Com::*, Console::*, LibraryLoader::GetModuleHandleW, Registry::*,
            SystemInformation::GetSystemDirectoryW, Threading::*,
        },
        UI::{
            Shell::{PropertiesSystem::PROPERTYKEY, *},
            WindowsAndMessaging::*,
        },
    },
};

const WM_APP_NOTIFYICON: u32 = WM_APP + 1;
const WM_APP_CALLBACK_ENDPOINT: u32 = WM_APP + 2;
const WM_APP_CALLBACK_VOLUME: u32 = WM_APP + 3;

const IDM_EXIT: u16 = 0u16.wrapping_sub(1);
const IDM_OPEN_SOUNDCONTROLPANEL: u16 = 0u16.wrapping_sub(2);
const IDM_SET_MAX_VOLUME: u16 = 0u16.wrapping_sub(3);
const IDM_SEPARATOR: u16 = 0u16.wrapping_sub(4);
const IDM_NO_ENDPOINTS: u16 = 0u16.wrapping_sub(5);

const LABEL_EXIT: PCWSTR = w!("E&xit\0");
const LABEL_OPEN_SOUNDCONTROLPANEL: PCWSTR = w!("Open Sound Control Panel\0");
const LABEL_SET_MAX_VOLUME: PCWSTR = w!("Set recording volume to 100%\0");
const LABEL_FORCE_MAX_VOLUME: &str = "Keeping recording volume at 100%";
const LABEL_NO_CAPTURE_DEVICES: PCWSTR = w!("No audio capture devices found\0");
const LABEL_NO_DEFAULT_DEVICE: &str = "No default communications audio capture device found!";
const LABEL_MUTED: &str = "muted";
const LABEL_VOLUME_UNKNOWN: &str = "volume unknown";

#[derive(Clone, Copy, EnumString, PartialEq)]
#[strum(serialize_all = "snake_case")]
enum IconTheme {
    Auto,
    Normal,
    Inverted,
}

// Message received when the taskbar is (re)created
thread_local!(static WM_TASKBAR_CREATED: RefCell<Option<u32>> = const { RefCell::new(None) });

// COM objects for interacting with Windows Audio
thread_local!(static AUDIO_POLICY_CONFIG: RefCell<Option<IPolicyConfig>> = const { RefCell::new(None) });
thread_local!(static AUDIO_ENDPOINT_ENUMERATOR: RefCell<Option<IMMDeviceEnumerator>> = const { RefCell::new(None) });
thread_local!(static AUDIO_DEFAULT_ENDPOINT: RefCell<Option<IMMDevice>> = const { RefCell::new(None) });
thread_local!(static AUDIO_DEFAULT_ENDPOINT_VOLUME: RefCell<Option<IAudioEndpointVolume>> = const { RefCell::new(None) });

// Volume configuration
thread_local!(static CONFIG_FORCE_MAX_VOLUME: RefCell<bool> = const { RefCell::new(false) });

// Icon configuration
thread_local!(static CONFIG_ICON_THEME: RefCell<IconTheme> = const { RefCell::new(IconTheme::Auto) });
thread_local!(static CONFIG_ICON_ACTIVE: RefCell<(Option<Vec<u16>>, PCWSTR, u32, Option<IconTheme>)> = const { RefCell::new((None, w!("%SystemRoot%\\System32\\SndVolSSO.dll"), 141u32.wrapping_neg(), None)) });
thread_local!(static CONFIG_ICON_MUTED: RefCell<(Option<Vec<u16>>, PCWSTR, u32, Option<IconTheme>)> = const { RefCell::new((None, w!("%SystemRoot%\\System32\\SndVolSSO.dll"), 140u32.wrapping_neg(), None)) });

// Icons for active and muted states
thread_local!(static ICON_ACTIVE: RefCell<Option<HICON>> = const { RefCell::new(None) });
thread_local!(static ICON_ACTIVE_INVERTED: RefCell<Option<HICON>> = const { RefCell::new(None) });
thread_local!(static ICON_MUTED: RefCell<Option<HICON>> = const { RefCell::new(None) });
thread_local!(static ICON_MUTED_INVERTED: RefCell<Option<HICON>> = const { RefCell::new(None) });

// Notify icon data registered to show in the notification tray
thread_local!(static NOTIFY_ICON_DATA: RefCell<Option<NOTIFYICONDATAW>> = const { RefCell::new(None) });

// Context menu shown when right-clicking the notify icon
thread_local!(static MENU: RefCell<Option<HMENU>> = const { RefCell::new(None) });
thread_local!(static MENU_AUDIO_ENDPOINTS: RefCell<Vec<PWSTR>> = const { RefCell::new(Vec::new()) });

// Callbacks for receiving notifications about changes
thread_local!(static AUDIO_ENDPOINT_CALLBACK: RefCell<Option<IMMNotificationClient>> = const { RefCell::new(None) });
thread_local!(static AUDIO_ENDPOINT_VOLUME_CALLBACK: RefCell<Option<IAudioEndpointVolumeCallback>> = const { RefCell::new(None) });

#[implement(IMMNotificationClient)]
struct AudioEndpointCallback {
    window: HWND,
}

#[allow(non_snake_case)]
impl IMMNotificationClient_Impl for AudioEndpointCallback_Impl {
    fn OnDeviceStateChanged(
        &self,
        _pwstrdeviceid: &PCWSTR,
        _dwnewstate: DEVICE_STATE,
    ) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)) }
    }
    fn OnDeviceAdded(&self, _pwstrdeviceid: &PCWSTR) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)) }
    }
    fn OnDeviceRemoved(&self, _pwstrdeviceid: &PCWSTR) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)) }
    }
    fn OnDefaultDeviceChanged(
        &self,
        _flow: EDataFlow,
        _role: ERole,
        _pwstrdefaultdeviceid: &PCWSTR,
    ) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)) }
    }
    fn OnPropertyValueChanged(&self, _pwstrdeviceid: &PCWSTR, _key: &PROPERTYKEY) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)) }
    }
}

#[implement(IAudioEndpointVolumeCallback)]
struct AudioEndpointVolumeCallback {
    window: HWND,
}

#[allow(non_snake_case)]
impl IAudioEndpointVolumeCallback_Impl for AudioEndpointVolumeCallback_Impl {
    fn OnNotify(&self, _pnotify: *mut AUDIO_VOLUME_NOTIFICATION_DATA) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_VOLUME, WPARAM(0), LPARAM(0)) }
    }
}

// Implementation of reversed engineered COM object for changing default audio endpoint
#[allow(non_upper_case_globals)]
pub const PolicyConfig: GUID = GUID::from_u128(0x870af99c_171d_4f9e_af0d_e63df40c2bc9);

define_interface!(
    IPolicyConfig,
    IPolicyConfig_Vtbl,
    0xf8679f50_850a_41cf_9c72_430f290290c8
);
impl std::ops::Deref for IPolicyConfig {
    type Target = IUnknown;
    fn deref(&self) -> &Self::Target {
        unsafe { std::mem::transmute(self) }
    }
}
interface_hierarchy!(IPolicyConfig, IUnknown);
impl IPolicyConfig {
    #[allow(non_snake_case, clippy::missing_safety_doc)]
    pub unsafe fn SetDefaultEndpoint<P0>(&self, wszDeviceId: P0, role: ERole) -> Result<()>
    where
        P0: Param<PWSTR>,
    {
        (Interface::vtable(self).SetDefaultEndpoint)(
            Interface::as_raw(self),
            wszDeviceId.param().abi(),
            role,
        )
        .ok()
    }
}

#[allow(non_snake_case)]
#[repr(C)]
pub struct IPolicyConfig_Vtbl {
    pub base__: IUnknown_Vtbl,
    pub GetMixFormat: unsafe extern "system" fn(
        this: *mut c_void,
        pwstrid: PWSTR,
        waveformatex: *mut c_void,
    ) -> HRESULT,
    pub GetDeviceFormat: unsafe extern "system" fn(
        this: *mut c_void,
        pwstrid: PWSTR,
        param0: i32,
        waveformatex: *mut c_void,
    ) -> HRESULT,
    pub ResetDeviceFormat: unsafe extern "system" fn(this: c_void, pwstrid: PWSTR) -> HRESULT,
    pub SetDeviceFormat: unsafe extern "system" fn(
        this: *mut c_void,
        pwstrid: PWSTR,
        waveformatex0: c_void,
        waveformatex1: *mut c_void,
    ) -> HRESULT,
    pub GetProcessingPeriod: unsafe extern "system" fn(
        this: *mut c_void,
        pwstrid: PWSTR,
        param0: i32,
        param1: c_void,
        param1: *mut c_void,
    ) -> HRESULT,
    pub SetProcessingPeriod:
        unsafe extern "system" fn(this: c_void, pwstrid: PWSTR, param0: c_void) -> HRESULT,
    pub GetShareMode: unsafe extern "system" fn(
        this: *mut c_void,
        pwstrid: PWSTR,
        devicesharemode: *mut c_void,
    ) -> HRESULT,
    pub SetShareMode: unsafe extern "system" fn(
        this: *mut c_void,
        pwstrid: PWSTR,
        devicesharemode: *mut c_void,
    ) -> HRESULT,
    pub GetPropertyValue: unsafe extern "system" fn(
        this: *mut c_void,
        pwstrid: PWSTR,
        key: c_void,
        propvariant: *mut c_void,
    ) -> HRESULT,
    pub SetPropertyValue: unsafe extern "system" fn(
        this: *mut c_void,
        pwstrid: PWSTR,
        key: c_void,
        propvariant: *mut c_void,
    ) -> HRESULT,
    pub SetDefaultEndpoint:
        unsafe extern "system" fn(this: *mut c_void, pwstrid: PWSTR, role: ERole) -> HRESULT,
    pub SetEndpointVisibility:
        unsafe extern "system" fn(this: *mut c_void, pwstrid: PWSTR, param0: i32) -> HRESULT,
}

// Audio initialization functions
fn init_audio_endpoint() -> Result<()> {
    deinit_audio_endpoint();
    AUDIO_DEFAULT_ENDPOINT.with(|global_audio_default_endpoint| {
        global_audio_default_endpoint.replace(AUDIO_ENDPOINT_ENUMERATOR.with(
            |global_audio_endpoint_enumerator| {
                match &*global_audio_endpoint_enumerator.borrow() {
                    Some(audio_endpoint_enumerator) => unsafe {
                        audio_endpoint_enumerator.GetDefaultAudioEndpoint(eCapture, eCommunications)
                    }
                    .map_or_else(
                        |error| {
                            if error.code() == ERROR_NOT_FOUND.to_hresult() {
                                Ok(None)
                            } else {
                                Err(error)
                            }
                        },
                        |audio_endpoint| Ok(Some(audio_endpoint)),
                    ),
                    _ => Ok(None),
                }
            },
        )?);
        Ok(())
    })
}

fn init_audio_endpoint_volume() -> Result<()> {
    deinit_audio_endpoint_volume();
    AUDIO_DEFAULT_ENDPOINT_VOLUME.with(|global_audio_endpoint_volume| {
        global_audio_endpoint_volume.replace(AUDIO_DEFAULT_ENDPOINT.with(
            |global_audio_endpoint| {
                match &*global_audio_endpoint.borrow() {
                    Some(audio_endpoint) => {
                        unsafe { audio_endpoint.Activate::<IAudioEndpointVolume>(CLSCTX_ALL, None) }
                            .map_or_else(
                                |error| {
                                    if error.code() == E_NOINTERFACE {
                                        Ok(None)
                                    } else {
                                        Err(error)
                                    }
                                },
                                |audio_endpoint_volume| {
                                    AUDIO_ENDPOINT_VOLUME_CALLBACK
                                        .with(|audio_endpoint_volume_callback| unsafe {
                                            audio_endpoint_volume.RegisterControlChangeNotify(
                                                (*audio_endpoint_volume_callback.borrow()).as_ref(),
                                            )
                                        })
                                        .map(|_| Some(audio_endpoint_volume))
                                },
                            )
                    }
                    _ => Ok(None),
                }
            },
        )?);
        Ok(())
    })
}

// Audio deinitialization functions
fn deinit_audio_endpoint_volume() {
    AUDIO_DEFAULT_ENDPOINT_VOLUME.with(|global| {
        if let Some(audio_endpoint_volume) = global.replace(None) {
            AUDIO_ENDPOINT_VOLUME_CALLBACK
                .with(|audio_endpoint_volume_callback| unsafe {
                    audio_endpoint_volume.UnregisterControlChangeNotify(
                        (*audio_endpoint_volume_callback.borrow()).as_ref(),
                    )
                })
                .ok();
        }
    });
}

fn deinit_audio_endpoint() {
    deinit_audio_endpoint_volume();
    AUDIO_DEFAULT_ENDPOINT.with(|global| {
        global.replace(None);
    });
}

// Function for toggling mute, used when clicking the icon
fn toggle_mute() -> Result<()> {
    AUDIO_DEFAULT_ENDPOINT_VOLUME.with(|global| match &*global.borrow() {
        Some(audio_endpoint_volume) => unsafe {
            audio_endpoint_volume.SetMute(
                !audio_endpoint_volume.GetMute()?.as_bool(),
                std::ptr::null(),
            )
        },
        _ => Ok(()),
    })
}

// Functions for setting the muted state
fn set_muted(muted: bool) -> Result<()> {
    AUDIO_DEFAULT_ENDPOINT_VOLUME.with(|global| match &*global.borrow() {
        Some(audio_endpoint_volume) => unsafe {
            audio_endpoint_volume.SetMute(muted, std::ptr::null())
        },
        _ => Ok(()),
    })
}

// Function for opening the Sound Control Panel
fn open_sound_control_panel_recording_tab() -> Result<()> {
    let system_directory = {
        let mut buffer = [0u16; MAX_PATH as usize];
        let length = unsafe { GetSystemDirectoryW(Some(&mut buffer)) };
        if length == 0 || length > buffer.len() as u32 {
            Err(Error::from_win32())
        } else {
            Ok(buffer
                .into_iter()
                .take(length as usize)
                .chain(std::iter::once(0))
                .collect::<Vec<u16>>())
        }
    }?;
    let command = system_directory
        .clone()
        .into_iter()
        .take_while(|c| *c != 0)
        .chain("\\control.exe".encode_utf16())
        .chain(std::iter::once(0))
        .collect::<Vec<u16>>();
    let mut arguments = command
        .clone()
        .into_iter()
        .take_while(|c| *c != 0)
        .chain(" mmsys.cpl,,1".to_string().encode_utf16())
        .chain(std::iter::once(0))
        .collect::<Vec<u16>>();
    unsafe {
        CreateProcessW(
            PCWSTR(command.as_ptr()),
            PWSTR(arguments.as_mut_ptr()),
            None,
            None,
            FALSE,
            DETACHED_PROCESS,
            None,
            PCWSTR(system_directory.as_ptr()),
            &STARTUPINFOW {
                ..Default::default()
            },
            &mut PROCESS_INFORMATION {
                ..Default::default()
            },
        )
    }
}

// Function for setting the volume to the maximum value
fn set_volume_to_max() -> Result<()> {
    AUDIO_DEFAULT_ENDPOINT_VOLUME.with(|global| match &*global.borrow() {
        Some(audio_endpoint_volume) => unsafe {
            if audio_endpoint_volume
                .GetMasterVolumeLevelScalar()
                .is_ok_and(|v| v < 1.0)
            {
                audio_endpoint_volume.SetMasterVolumeLevelScalar(1.0, std::ptr::null())
            } else {
                Ok(())
            }
        },
        _ => Ok(()),
    })
}

// Function for setting the default audio device
fn set_default_audio_capture_device(device_id: PWSTR) -> Result<()> {
    AUDIO_POLICY_CONFIG.with(|global_audio_policy_config| {
        match &*global_audio_policy_config.borrow() {
            Some(audio_policy_config) => {
                for role in [eConsole, eMultimedia, eCommunications] {
                    unsafe { audio_policy_config.SetDefaultEndpoint(device_id, role) }?;
                }
                Ok(())
            }
            .map_or_else(
                |error: Error| {
                    if error.code() == NTE_NOT_FOUND {
                        Ok(())
                    } else {
                        Err(error)
                    }
                },
                |_| Ok(()),
            ),
            _ => Ok(()),
        }
    })
}

// Create a copy of the icon with its colors inverted
fn invert_icon(icon: HICON) -> Result<HICON> {
    unsafe {
        // Get info about the icon
        let mut icon_info: ICONINFO = Default::default();
        GetIconInfo(icon, &mut icon_info)?;

        // Check we have a color icon we can invert
        if icon_info.hbmColor.is_invalid() {
            return Ok(icon);
        }

        // Retrieve the icon bitmap from the handle
        let mut icon_bitmap: BITMAP = Default::default();
        assert!(
            Gdi::GetObjectW(
                icon_info.hbmColor,
                std::mem::size_of::<BITMAP>() as i32,
                Some(&mut icon_bitmap as *mut _ as *mut _)
            ) as usize
                == std::mem::size_of::<BITMAP>(),
            "Failed to read icon bitmap"
        );

        // Create device context for accessing the icon
        let icon_dc: HDC = CreateCompatibleDC(None);
        let icon_dc_prevobj = Gdi::SelectObject(icon_dc, icon_info.hbmColor);

        // Create new 32-bit RGBA bitmap to contain inverted icon
        let mut inverted_info = BITMAPV5HEADER {
            bV5Size: std::mem::size_of::<BITMAPV5HEADER>() as u32,
            bV5Width: icon_bitmap.bmWidth,
            bV5Height: icon_bitmap.bmHeight,
            bV5Planes: 1,
            bV5BitCount: 32,
            bV5Compression: BI_BITFIELDS,
            bV5RedMask: 0x00ff0000,
            bV5GreenMask: 0x0000ff00,
            bV5BlueMask: 0x000000ff,
            bV5AlphaMask: 0xff000000,
            ..Default::default()
        };

        // Pointer to the RGBA pixels of the inverted bitmap
        let mut inverted_pixels: *mut u32 = std::ptr::null_mut();

        // Create inverted bitmap using a new device context
        let inverted_dc: HDC = CreateCompatibleDC(None);
        let inverted_bitmap = CreateDIBSection(
            inverted_dc,
            &mut inverted_info as *mut _ as *const BITMAPINFO,
            DIB_RGB_COLORS,
            &mut inverted_pixels as *mut _ as *mut *mut c_void,
            None,
            0,
        )?;
        let inverted_dc_prevobj = Gdi::SelectObject(inverted_dc, inverted_bitmap);

        // Create a copy of the icon by blitting from the DC with the original to the new one
        BitBlt(
            inverted_dc,
            0,
            0,
            icon_bitmap.bmWidth,
            icon_bitmap.bmHeight,
            icon_dc,
            0,
            0,
            SRCCOPY,
        )?;

        // Remove DC for original icon since we don't need it after blitting is done
        Gdi::SelectObject(icon_dc, icon_dc_prevobj);
        DeleteDC(icon_dc).ok()?;

        // Invert the pixels which we can now access via the pointer
        for i in 0..(icon_bitmap.bmWidth * icon_bitmap.bmHeight) as isize {
            *(inverted_pixels.offset(i)) ^= 0x00ffffff;
        }

        // Create a new icon with our modified color data
        icon_info.hbmColor = inverted_bitmap;
        let inverted = CreateIconIndirect(&icon_info)?;

        // Remove DC for inverted icon since we have now created it
        Gdi::SelectObject(inverted_dc, inverted_dc_prevobj);
        DeleteDC(inverted_dc).ok()?;

        Ok(inverted)
    }
}

// Load an icon and invert it if required
fn load_icon(
    instance: HINSTANCE,
    path: PCWSTR,
    index: u32,
    theme: Option<IconTheme>,
) -> Result<(HICON, HICON)> {
    let icon = unsafe { ExtractIconW(instance, path, index) };
    assert!(!icon.is_invalid(), "Icon is not valid.");
    let icon_inverted = invert_icon(icon)?;
    assert!(!icon_inverted.is_invalid(), "Inverted icon is not valid.");
    Ok(
        match theme.unwrap_or(CONFIG_ICON_THEME.with(|global| *global.borrow())) {
            IconTheme::Auto => (icon, icon_inverted),
            IconTheme::Normal => (icon, icon),
            IconTheme::Inverted => (icon_inverted, icon_inverted),
        },
    )
}

// Load the icons to use from the icon file
fn load_icons(instance: HINSTANCE) -> Result<()> {
    // Load the icons
    let (icon_active, icon_active_inverted) = CONFIG_ICON_ACTIVE.with(|global| {
        let (_, path, index, theme) = &*global.borrow();
        load_icon(instance, *path, *index, *theme)
    })?;
    let (icon_muted, icon_muted_inverted) = CONFIG_ICON_MUTED.with(|global| {
        let (_, path, index, theme) = &*global.borrow();
        load_icon(instance, *path, *index, *theme)
    })?;

    // Replace currently loaded icons
    ICON_ACTIVE.with(|global| match global.replace(Some(icon_active)) {
        Some(old_icon) => unsafe { DestroyIcon(old_icon) },
        _ => Ok(()),
    })?;
    ICON_ACTIVE_INVERTED.with(|global| match global.replace(Some(icon_active_inverted)) {
        Some(old_icon) => unsafe { DestroyIcon(old_icon) },
        _ => Ok(()),
    })?;
    ICON_MUTED.with(|global| match global.replace(Some(icon_muted)) {
        Some(old_icon) => unsafe { DestroyIcon(old_icon) },
        _ => Ok(()),
    })?;
    ICON_MUTED_INVERTED.with(|global| match global.replace(Some(icon_muted_inverted)) {
        Some(old_icon) => unsafe { DestroyIcon(old_icon) },
        _ => Ok(()),
    })?;

    Ok(())
}

// Function determining whether we should use the inverted icon by checking if Windows is in dark or light mode
fn should_use_inverted_icon() -> bool {
    let mut buffer = [0u8; 4];
    let mut size: u32 = 4;
    if unsafe {
        RegGetValueW(
            HKEY_CURRENT_USER,
            w!("Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize"),
            w!("AppsUseLightTheme"),
            RRF_RT_REG_DWORD,
            None,
            Some(buffer.as_mut_ptr() as *mut c_void),
            Some(&mut size as *mut u32),
        )
        .ok()
        .is_err()
    } {
        return false;
    }
    assert!(size == 4, "Invalid size for DWORD.");
    i32::from_le_bytes(buffer) != 0
}

// Retrieves the microphone status and updates the icon and tooltip
fn update_icon_data() -> Result<()> {
    let (icon_active, icon_muted) = if should_use_inverted_icon() {
        (&ICON_ACTIVE_INVERTED, &ICON_MUTED_INVERTED)
    } else {
        (&ICON_ACTIVE, &ICON_MUTED)
    };
    let (icon, text) = AUDIO_DEFAULT_ENDPOINT.with(|global_audio_endpoint| {
        match &*global_audio_endpoint.borrow() {
            Some(audio_endpoint) => {
                let device_name = unsafe {
                    audio_endpoint
                        .OpenPropertyStore(STGM_READ)?
                        .GetValue(&PKEY_Device_FriendlyName)?
                        .to_string()
                };
                let device_name_prefix = device_name.encode_utf16().chain(": ".encode_utf16());
                AUDIO_DEFAULT_ENDPOINT_VOLUME.with(|global| match &*global.borrow() {
                    Some(audio_endpoint_volume) => {
                        let volume = if unsafe { audio_endpoint_volume.GetMute() }?.into() {
                            None
                        } else {
                            Some(format!(
                                "{:.0}%",
                                100f32
                                    * unsafe {
                                        audio_endpoint_volume.GetMasterVolumeLevelScalar()
                                    }?
                            ))
                        };

                        Ok::<_, HRESULT>((
                            if volume.is_some() {
                                icon_active
                            } else {
                                icon_muted
                            },
                            device_name_prefix
                                .chain(
                                    match volume {
                                        Some(volume_text) => volume_text,
                                        _ => LABEL_MUTED.to_owned(),
                                    }
                                    .encode_utf16(),
                                )
                                .collect::<Vec<_>>(),
                        ))
                    }
                    _ => Ok((
                        icon_muted,
                        device_name_prefix
                            .chain(LABEL_VOLUME_UNKNOWN.encode_utf16())
                            .collect(),
                    )),
                })
            }
            _ => Ok((icon_muted, LABEL_NO_DEFAULT_DEVICE.encode_utf16().collect())),
        }
    })?;
    NOTIFY_ICON_DATA.with(|global_notify_icon_data| {
        global_notify_icon_data.replace_with(|previous_notify_icon_data| {
            let mut notify_icon_data = previous_notify_icon_data.unwrap_or_default();
            let max_text_len = notify_icon_data.szTip.len() - 1;
            icon.with(|icon_ref| match *icon_ref.borrow() {
                Some(icon_id) if !icon_id.is_invalid() => {
                    notify_icon_data.hIcon = icon_id;
                    notify_icon_data.uFlags |= NIF_ICON;
                }
                _ => {
                    notify_icon_data.hIcon = Default::default();
                    notify_icon_data.uFlags &= !NIF_ICON;
                }
            });
            notify_icon_data
                .szTip
                .iter_mut()
                .zip(
                    text.clone()
                        .into_iter()
                        .take(max_text_len)
                        .chain(std::iter::repeat(0)),
                )
                .for_each(|(ptr, chr)| *ptr = chr);
            if notify_icon_data.szTip[0] == 0 {
                notify_icon_data.uFlags &= !(NIF_TIP | NIF_SHOWTIP);
            } else {
                notify_icon_data.uFlags |= NIF_TIP | NIF_SHOWTIP;
            }
            Some(notify_icon_data)
        })
    });

    Ok(())
}

// Add the notify icon for when it does not already exists
fn add_notify_icon() -> Result<()> {
    update_icon_data()?;
    NOTIFY_ICON_DATA.with(|global| {
        if let Some(notify_icon_data) = global.borrow().as_ref() {
            unsafe {
                Shell_NotifyIconW(NIM_ADD, notify_icon_data).ok()?;
                Shell_NotifyIconW(NIM_SETVERSION, notify_icon_data).ok()?;
            }
            Ok(())
        } else {
            Ok(())
        }
    })
}

// Update the existing notify icon
fn update_notify_icon() -> Result<()> {
    update_icon_data()?;
    NOTIFY_ICON_DATA.with(|global| {
        if let Some(notify_icon_data) = global.borrow().as_ref() {
            unsafe { Shell_NotifyIconW(NIM_MODIFY, notify_icon_data).ok() }
        } else {
            Ok(())
        }
    })
}

fn pwstr_eq(a: PWSTR, b: PWSTR) -> bool {
    let mut offset = 0;
    loop {
        let (chr_a, chr_b) = unsafe { (*a.0.add(offset), *b.0.add(offset)) };
        if chr_a != chr_b {
            return false;
        }
        if chr_a == 0 || chr_b == 0 {
            return true;
        }
        offset += 1;
    }
}

// Update the menu
fn update_menu() -> Result<()> {
    MENU.with(|global_menu| {
        AUDIO_ENDPOINT_ENUMERATOR.with(|global_audio_endpoint_enumerator| {
            match &*global_audio_endpoint_enumerator.borrow() {
                Some(audio_endpoint_enumerator) => {
                    let devices = unsafe {
                        audio_endpoint_enumerator.EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE)
                    }?;
                    let devices_count = unsafe { devices.GetCount() }? as usize;
                    if global_menu.borrow().is_none() {
                        global_menu.replace(Some({
                            let menu = unsafe { CreatePopupMenu() }?;
                            unsafe {
                                AppendMenuW(
                                    menu,
                                    MF_DISABLED | MF_GRAYED,
                                    IDM_NO_ENDPOINTS as usize,
                                    LABEL_NO_CAPTURE_DEVICES,
                                )?;
                                AppendMenuW(
                                    menu,
                                    MF_SEPARATOR,
                                    IDM_SEPARATOR as usize,
                                    PCWSTR::null(),
                                )?;
                                CONFIG_FORCE_MAX_VOLUME.with(|force_max_volume| {
                                    if *force_max_volume.borrow() {
                                        let mut label_buffer = LABEL_FORCE_MAX_VOLUME
                                            .to_string()
                                            .encode_utf16()
                                            .chain(std::iter::once(0))
                                            .collect::<Vec<u16>>();
                                        let label = PWSTR(label_buffer.as_mut_ptr());
                                        InsertMenuItemW(
                                            menu,
                                            IDM_SET_MAX_VOLUME as u32,
                                            false,
                                            &MENUITEMINFOW {
                                                cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                                                fMask: MIIM_FTYPE
                                                    | MIIM_ID
                                                    | MIIM_STATE
                                                    | MIIM_STRING,
                                                fType: MFT_STRING,
                                                fState: MFS_CHECKED | MFS_DISABLED,
                                                wID: IDM_SET_MAX_VOLUME as u32,
                                                dwTypeData: label,
                                                ..Default::default()
                                            },
                                        )
                                    } else {
                                        AppendMenuW(
                                            menu,
                                            MF_ENABLED | MF_STRING,
                                            IDM_SET_MAX_VOLUME as usize,
                                            LABEL_SET_MAX_VOLUME,
                                        )
                                    }
                                })?;
                                AppendMenuW(
                                    menu,
                                    MF_ENABLED | MF_STRING,
                                    IDM_OPEN_SOUNDCONTROLPANEL as usize,
                                    LABEL_OPEN_SOUNDCONTROLPANEL,
                                )?;
                                AppendMenuW(
                                    menu,
                                    MF_ENABLED | MF_STRING,
                                    IDM_EXIT as usize,
                                    LABEL_EXIT,
                                )?;
                            }
                            menu
                        }));
                    }

                    let menuref = &*global_menu.borrow();
                    let menu = menuref.as_ref();
                    let default_endpoint_id =
                        AUDIO_DEFAULT_ENDPOINT.with(|global_audio_default_endpoint| {
                            global_audio_default_endpoint.borrow().as_ref().and_then(
                                |audio_default_endpoint| {
                                    unsafe { audio_default_endpoint.GetId() }.ok()
                                },
                            )
                        });
                    MENU_AUDIO_ENDPOINTS.with(|global_menu_audio_endpoints| -> Result<()> {
                        let mut menu_audio_endpoints = global_menu_audio_endpoints.borrow_mut();
                        if menu_audio_endpoints.len() == 0 && devices_count > 0 {
                            unsafe {
                                RemoveMenu(menu, IDM_NO_ENDPOINTS as u32, MF_BYCOMMAND)?;
                            }
                        }
                        for i in 0..devices_count {
                            let device = unsafe { devices.Item(i as u32)? };
                            let device_id = unsafe { device.GetId() }?;
                            let mut device_name_buffer = unsafe {
                                device
                                    .OpenPropertyStore(STGM_READ)?
                                    .GetValue(&PKEY_Device_FriendlyName)?
                            }
                            .to_string()
                            .encode_utf16()
                            .chain(std::iter::once(0))
                            .collect::<Vec<u16>>();
                            let device_name = PWSTR(device_name_buffer.as_mut_ptr());
                            let device_is_default = match default_endpoint_id {
                                Some(id) => pwstr_eq(device_id, id),
                                _ => false,
                            };
                            let mut found = false;
                            for j in i..menu_audio_endpoints.len() {
                                if pwstr_eq(device_id, menu_audio_endpoints[i]) {
                                    found = true;
                                    for _ in 0..(j - i) {
                                        unsafe {
                                            CoTaskMemFree(Some(
                                                menu_audio_endpoints.remove(i).0 as *const c_void,
                                            ));
                                            RemoveMenu(menu, i as u32, MF_BYPOSITION)?;
                                        }
                                    }
                                    unsafe {
                                        SetMenuItemInfoW(
                                            menu,
                                            i as u32,
                                            true,
                                            &MENUITEMINFOW {
                                                cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                                                fMask: MIIM_ID | MIIM_STATE | MIIM_STRING,
                                                fState: if device_is_default {
                                                    MFS_CHECKED
                                                } else {
                                                    MFS_UNCHECKED
                                                },
                                                wID: i as u32,
                                                dwTypeData: device_name,
                                                ..Default::default()
                                            },
                                        )?;
                                    }
                                    break;
                                }
                            }
                            if !found {
                                // Doesn't exist, insert
                                menu_audio_endpoints.insert(i, device_id);
                                unsafe {
                                    InsertMenuItemW(
                                        menu,
                                        i as u32,
                                        true,
                                        &MENUITEMINFOW {
                                            cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                                            fMask: MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_STRING,
                                            fType: MFT_STRING,
                                            fState: if device_is_default {
                                                MFS_CHECKED
                                            } else {
                                                MFS_UNCHECKED
                                            },
                                            wID: i as u32,
                                            dwTypeData: device_name,
                                            ..Default::default()
                                        },
                                    )?;
                                }
                            }
                        }
                        let add_no_devices_label =
                            devices_count == 0 && menu_audio_endpoints.len() > 0;
                        while menu_audio_endpoints.len() > devices_count {
                            unsafe {
                                CoTaskMemFree(Some(
                                    menu_audio_endpoints.pop().unwrap().0 as *const c_void,
                                ));
                                RemoveMenu(menu, devices_count as u32, MF_BYPOSITION)?
                            };
                        }
                        if add_no_devices_label {
                            unsafe {
                                InsertMenuItemW(
                                    menu,
                                    0,
                                    true,
                                    &MENUITEMINFOW {
                                        cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                                        fMask: MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_STRING,
                                        fType: MFT_STRING,
                                        fState: MFS_DISABLED | MFS_GRAYED,
                                        wID: IDM_NO_ENDPOINTS as u32,
                                        dwTypeData: PWSTR(
                                            LABEL_NO_CAPTURE_DEVICES
                                                .to_string()?
                                                .encode_utf16()
                                                .chain(std::iter::once(0))
                                                .collect::<Vec<u16>>()
                                                .as_mut_ptr(),
                                        ),
                                        ..Default::default()
                                    },
                                )?;
                            }
                        }
                        Ok(())
                    })?;

                    if let Some(id) = default_endpoint_id {
                        unsafe {
                            CoTaskMemFree(Some(id.0 as *const c_void));
                        }
                    };
                    Ok(())
                }
                _ => Ok(()),
            }
        })
    })
}

// Show the menu at the given coordinates
fn show_menu(window: HWND, x: i32, y: i32) -> Result<()> {
    update_menu()?;
    MENU.with(|menu| {
        unsafe {
            // Set our window as foreground so the menu disappears when focus is lost
            SetForegroundWindow(window).ok().ok();
            TrackPopupMenuEx(
                (*menu.borrow()).as_ref(),
                (TPM_RIGHTBUTTON
                    | if GetSystemMetrics(SM_MENUDROPALIGNMENT) != 0 {
                        TPM_RIGHTALIGN
                    } else {
                        TPM_LEFTALIGN
                    })
                .0,
                x,
                y,
                window,
                None,
            )
            .ok()
        }
    })
}

// Callback for message pump
extern "system" fn window_callback(
    window: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match message {
        WM_CREATE => {
            // Listen for taskbar created messages since we should re-add our icon
            let wm_taskbar_created = unsafe { RegisterWindowMessageW(w!("TaskbarCreated")) };
            assert!(wm_taskbar_created != 0);
            WM_TASKBAR_CREATED.with(|global| {
                global.replace(Some(wm_taskbar_created));
            });
            LRESULT(0)
        }
        WM_DPICHANGED => {
            let instance: HINSTANCE = unsafe { GetModuleHandleW(None).unwrap().into() };
            assert!(!instance.is_invalid());
            load_icons(instance).unwrap();
            update_notify_icon().ok();
            LRESULT(0)
        }
        WM_SETTINGCHANGE => {
            update_notify_icon().ok();
            LRESULT(0)
        }
        WM_APP_NOTIFYICON => match lparam.0 as u32 & 0xffff {
            NIN_SELECT => {
                toggle_mute().ok();
                LRESULT(0)
            }
            WM_CONTEXTMENU => {
                show_menu(
                    window,
                    (wparam.0 as u32 & 0xffff) as i16 as i32,
                    ((wparam.0 as u32 >> 16) & 0xffff) as i16 as i32,
                )
                .unwrap();
                LRESULT(0)
            }
            _ => LRESULT(0),
        },
        WM_APP_CALLBACK_ENDPOINT => {
            // Audio endpoint has changed
            init_audio_endpoint()
                .and_then(|_| init_audio_endpoint_volume())
                .ok();
            update_notify_icon().ok();
            update_menu()
                .and_then(|()| unsafe { DrawMenuBar(window) })
                .ok();
            LRESULT(0)
        }
        WM_APP_CALLBACK_VOLUME => {
            // Audio endpoint volume/mute has changed
            update_notify_icon().ok();
            CONFIG_FORCE_MAX_VOLUME
                .with(|global| {
                    if *global.borrow() {
                        set_volume_to_max()
                    } else {
                        Ok(())
                    }
                })
                .ok();
            LRESULT(0)
        }
        WM_COMMAND => {
            match (wparam.0 as u32 & 0xffff) as u16 {
                IDM_EXIT => unsafe {
                    DestroyWindow(window).ok();
                },
                IDM_OPEN_SOUNDCONTROLPANEL => {
                    open_sound_control_panel_recording_tab().ok();
                }
                IDM_SET_MAX_VOLUME => {
                    set_volume_to_max().ok();
                }
                i => {
                    let i = i as usize;
                    MENU_AUDIO_ENDPOINTS.with(|global_menu_audio_endpoints| {
                        let menu_audio_endpoints = global_menu_audio_endpoints.borrow();
                        if i < menu_audio_endpoints.len() {
                            set_default_audio_capture_device(menu_audio_endpoints[i]).ok();
                        }
                    });
                }
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            unsafe {
                DestroyWindow(window).ok();
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            unsafe {
                PostQuitMessage(0);
            }
            LRESULT(0)
        }
        _ => WM_TASKBAR_CREATED.with(|wm_taskbar_created| {
            if wm_taskbar_created.borrow().map_or(false, |m| m == message) {
                // We also get TaskbarCreated messages when the DPI is changed,
                // but in that case the icon still exists, so if adding it fails
                // we try updating it instead
                add_notify_icon().or_else(|_| update_notify_icon()).unwrap();
                LRESULT(0)
            } else {
                unsafe { DefWindowProcW(window, message, wparam, lparam) }
            }
        }),
    }
}

#[derive(FromArgs)]
/// Show the microphone mute status in the systray.
struct CliArgs {
    /// config: force keep volume at 100%
    #[argh(switch)]
    config_force_keep_volume_at_max: bool,

    /// config: icon theme selection
    #[argh(option)]
    config_icon_theme: Option<IconTheme>,

    /// config: custom icon when not muted
    #[argh(option)]
    config_icon_active: Option<PathBuf>,

    /// config: custom icon when muted
    #[argh(option)]
    config_icon_muted: Option<PathBuf>,

    /// action: mute microphone
    #[argh(switch)]
    action_mute: bool,

    /// action: unmute microphone
    #[argh(switch)]
    action_unmute: bool,

    /// action: toggle microphone mute
    #[argh(switch)]
    action_toggle_mute: bool,

    /// action: set recording volume to 100%
    #[argh(switch)]
    action_set_volume_to_max: bool,

    /// action: exit immediately
    #[argh(switch)]
    action_exit: bool,
}

use std::os::windows::ffi::OsStrExt;

fn main() -> Result<()> {
    let instance: HINSTANCE = unsafe { GetModuleHandleW(None)?.into() };
    assert!(!instance.is_invalid());

    // Attach to parent console so we can output help messages etc
    unsafe { AttachConsole(ATTACH_PARENT_PROCESS) }.ok();

    // Parse CLI arguments
    let args: CliArgs = argh::from_env();

    // Set configuration options
    CONFIG_FORCE_MAX_VOLUME.with(|global| global.replace(args.config_force_keep_volume_at_max));
    if let Some(icon_theme) = args.config_icon_theme {
        CONFIG_ICON_THEME.with(|global| global.replace(icon_theme));
    }
    let custom_icon_theme = args.config_icon_theme.or(Some(IconTheme::Normal));
    if let Some(icon_active) = args.config_icon_active {
        let path_buffer = icon_active
            .as_os_str()
            .encode_wide()
            .chain(std::iter::once(0))
            .collect::<Vec<_>>();
        let path = PCWSTR(path_buffer.as_ptr());
        CONFIG_ICON_ACTIVE
            .with(|global| global.replace((Some(path_buffer), path, 0, custom_icon_theme)));
    }
    if let Some(icon_muted) = args.config_icon_muted {
        let path_buffer = icon_muted
            .as_os_str()
            .encode_wide()
            .chain(std::iter::once(0))
            .collect::<Vec<_>>();
        let path = PCWSTR(path_buffer.as_ptr());
        CONFIG_ICON_MUTED
            .with(|global| global.replace((Some(path_buffer), path, 0, custom_icon_theme)));
    }

    // Main window class definition
    let window_class_name_buffer = "MicrophoneMuteIndicator\0"
        .encode_utf16()
        .collect::<Vec<_>>();
    let window_class_name = PCWSTR(window_class_name_buffer.as_ptr());
    let window_class = WNDCLASSW {
        hInstance: instance,
        lpszClassName: window_class_name,
        lpfnWndProc: Some(window_callback),
        ..Default::default()
    };

    // Register the main window class
    let window_class_atom = unsafe { RegisterClassW(&window_class) };
    assert!(window_class_atom != 0);

    // Create a window using the class we just defined
    let window = unsafe {
        CreateWindowExW(
            WS_EX_NOACTIVATE,
            window_class_name,
            window_class_name,
            WS_DISABLED,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            None,
            None,
            instance,
            None,
        )
    }?;
    assert!(!window.is_invalid());

    // Initialize COM runtime
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED).ok()?;
    }

    // Initiate callbacks
    AUDIO_ENDPOINT_CALLBACK
        .with(|global| global.replace(Some((AudioEndpointCallback { window }).into())));
    AUDIO_ENDPOINT_VOLUME_CALLBACK
        .with(|global| global.replace(Some((AudioEndpointVolumeCallback { window }).into())));

    // Set up audio endpoint configuration
    let audio_policy_config: IPolicyConfig =
        unsafe { CoCreateInstance(&PolicyConfig, None, CLSCTX_ALL) }?;
    AUDIO_POLICY_CONFIG.with(|global| {
        global.replace(Some(audio_policy_config));
    });

    // Set up audio endpoint enumeration
    let audio_endpoint_enumerator: IMMDeviceEnumerator =
        unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL) }?;
    AUDIO_ENDPOINT_CALLBACK.with(|audio_endpoint_callback| unsafe {
        audio_endpoint_enumerator
            .RegisterEndpointNotificationCallback((*audio_endpoint_callback.borrow()).as_ref())
    })?;
    AUDIO_ENDPOINT_ENUMERATOR.with(|global| {
        global.replace(Some(audio_endpoint_enumerator));
    });

    // Get the audio endpoint and volume objects
    init_audio_endpoint()?;
    init_audio_endpoint_volume()?;

    // Execute actions
    if args.action_mute {
        set_muted(true)?;
    }
    if args.action_unmute {
        set_muted(false)?;
    }
    if args.action_toggle_mute {
        toggle_mute()?;
    }
    if args.action_set_volume_to_max {
        set_volume_to_max()?;
    }

    // Only add icon if we're not exiting immediately
    let mut exit_result = Ok(());
    if !args.action_exit {
        // Add the notify icon
        let notify_icon_data = NOTIFYICONDATAW {
            cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
            hWnd: window,
            uID: 0,
            uFlags: NIF_MESSAGE,
            uCallbackMessage: WM_APP_NOTIFYICON,
            Anonymous: NOTIFYICONDATAW_0 {
                uVersion: NOTIFYICON_VERSION_4,
            },
            ..Default::default()
        };
        NOTIFY_ICON_DATA.with(|global| {
            global.replace(Some(notify_icon_data));
        });

        load_icons(instance)?;
        add_notify_icon()?;
    } else {
        unsafe { DestroyWindow(window) }?;
    }

    // Message pump
    let mut message = MSG::default();
    while unsafe { GetMessageW(&mut message, HWND(std::ptr::null_mut()), 0, 0) }.into() {
        unsafe {
            DispatchMessageW(&message);
        }
    }

    if message.wParam.0 != 0 {
        exit_result = Err(Error::from_win32())
    }

    // Remove the notification icon
    NOTIFY_ICON_DATA.with(|global| match global.replace(None) {
        Some(notify_icon_data) => unsafe { Shell_NotifyIconW(NIM_DELETE, &notify_icon_data).ok() },
        _ => Ok(()),
    })?;

    // Release the COM objects
    deinit_audio_endpoint();
    AUDIO_POLICY_CONFIG.with(|global| {
        global.replace(None);
    });
    AUDIO_ENDPOINT_ENUMERATOR.with(|global| {
        if let Some(audio_endpoint_enumerator) = global.replace(None) {
            AUDIO_ENDPOINT_CALLBACK.with(|audio_endpoint_callback| unsafe {
                audio_endpoint_enumerator
                    .UnregisterEndpointNotificationCallback(
                        (*audio_endpoint_callback.borrow()).as_ref(),
                    )
                    .unwrap()
            });
        }
    });

    // Deinitialize COM runtime
    unsafe {
        CoUninitialize();
    }

    // Deallocate icons
    ICON_ACTIVE.with(|global| match global.replace(None) {
        Some(old_icon) => unsafe { DestroyIcon(old_icon) },
        _ => Ok(()),
    })?;
    ICON_MUTED.with(|global| match global.replace(None) {
        Some(old_icon) => unsafe { DestroyIcon(old_icon) },
        _ => Ok(()),
    })?;

    // Destroy the menu
    MENU.with(|global| {
        if let Some(menu) = global.replace(None) {
            unsafe {
                DestroyMenu(menu).unwrap();
            }
        }
    });

    // Unregister the window class
    unsafe {
        UnregisterClassW(window_class_name, instance)?;
    }

    exit_result
}
