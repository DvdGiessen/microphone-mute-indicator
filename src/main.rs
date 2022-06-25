// Simple microphone mute indicator for Windows
// Author: Daniël van de Giessen <daniel@dvdgiessen.nl>

#![windows_subsystem = "windows"]

use core::fmt;
use std::mem;
use std::{cell::RefCell, ffi::c_void, thread::LocalKey};

use windows::{
    core::*,
    Win32::Foundation::*,
    Win32::Media::Audio::Endpoints::*,
    Win32::Media::Audio::*,
    Win32::System::Com::*,
    Win32::UI::WindowsAndMessaging::*,
    Win32::{
        Devices::FunctionDiscovery::PKEY_Device_FriendlyName,
        System::Com::StructuredStorage::STGM_READ, UI::Shell::*,
    },
    Win32::{System::LibraryLoader::GetModuleHandleW, UI::Shell::PropertiesSystem::PROPERTYKEY},
};

const WM_APP_NOTIFYICON: u32 = WM_APP + 1;
const WM_APP_CALLBACK_ENDPOINT: u32 = WM_APP + 2;
const WM_APP_CALLBACK_VOLUME: u32 = WM_APP + 3;

const IDM_EXIT: u16 = 0u16.wrapping_sub(1);
const IDM_SEPARATOR: u16 = 0u16.wrapping_sub(2);
const IDM_NO_ENDPOINTS: u16 = 0u16.wrapping_sub(3);

const LABEL_EXIT: &str = "E&xit\0";
const LABEL_NO_CAPTURE_DEVICES: &str = "No audio capture devices found\0";
const LABEL_NO_DEFAULT_DEVICE: &str = "No default communications audio capture device found!";
const LABEL_MUTED: &str = "muted";
const LABEL_VOLUME_UNKNOWN: &str = "volume unknown";

// Message received when the taskbar is (re)created
thread_local!(static WM_TASKBAR_CREATED: RefCell<Option<u32>> = RefCell::new(None));

// COM objects for interacting with Windows Audio
thread_local!(static AUDIO_POLICY_CONFIG: RefCell<Option<IPolicyConfig>> = RefCell::new(None));
thread_local!(static AUDIO_ENDPOINT_ENUMERATOR: RefCell<Option<IMMDeviceEnumerator>> = RefCell::new(None));
thread_local!(static AUDIO_DEFAULT_ENDPOINT: RefCell<Option<IMMDevice>> = RefCell::new(None));
thread_local!(static AUDIO_DEFAULT_ENDPOINT_VOLUME: RefCell<Option<IAudioEndpointVolume>> = RefCell::new(None));

// Icons for active and muted states
thread_local!(static ICON_ACTIVE: RefCell<Option<HICON>> = RefCell::new(None));
thread_local!(static ICON_MUTED: RefCell<Option<HICON>> = RefCell::new(None));

// Notify icon data registered to show in the notification tray
thread_local!(static NOTIFY_ICON_DATA: RefCell<NOTIFYICONDATAW> = RefCell::new(Default::default()));

// Context menu shown when right-clicking the notify icon
thread_local!(static MENU: RefCell<Option<HMENU>> = RefCell::new(None));
thread_local!(static MENU_AUDIO_ENDPOINTS: RefCell<Vec<PWSTR>> = RefCell::new(Vec::new()));

// Callbacks for receiving notifications about changes
thread_local!(static AUDIO_ENDPOINT_CALLBACK: RefCell<Option<IMMNotificationClient>> = RefCell::new(None));
thread_local!(static AUDIO_ENDPOINT_VOLUME_CALLBACK: RefCell<Option<IAudioEndpointVolumeCallback>> = RefCell::new(None));

#[windows::core::implement(IMMNotificationClient)]
struct AudioEndpointCallback {
    window: HWND,
}

#[allow(non_snake_case)]
impl IMMNotificationClient_Impl for AudioEndpointCallback {
    fn OnDeviceStateChanged(&self, _pwstrdeviceid: &PCWSTR, _dwnewstate: u32) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)).ok() }
    }
    fn OnDeviceAdded(&self, _pwstrdeviceid: &PCWSTR) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)).ok() }
    }
    fn OnDeviceRemoved(&self, _pwstrdeviceid: &PCWSTR) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)).ok() }
    }
    fn OnDefaultDeviceChanged(
        &self,
        _flow: EDataFlow,
        _role: ERole,
        _pwstrdefaultdeviceid: &PCWSTR,
    ) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)).ok() }
    }
    fn OnPropertyValueChanged(&self, _pwstrdeviceid: &PCWSTR, _key: &PROPERTYKEY) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_ENDPOINT, WPARAM(0), LPARAM(0)).ok() }
    }
}

#[windows::core::implement(IAudioEndpointVolumeCallback)]
struct AudioEndpointVolumeCallback {
    window: HWND,
}

#[allow(non_snake_case)]
impl IAudioEndpointVolumeCallback_Impl for AudioEndpointVolumeCallback {
    fn OnNotify(&self, _pnotify: *mut AUDIO_VOLUME_NOTIFICATION_DATA) -> Result<()> {
        unsafe { PostMessageW(self.window, WM_APP_CALLBACK_VOLUME, WPARAM(0), LPARAM(0)).ok() }
    }
}

// Implementation of reversed engineered COM object for changing default audio endpoint
#[allow(non_upper_case_globals)]
pub const PolicyConfig: GUID = GUID::from_u128(0x870af99c_171d_4f9e_af0d_e63df40c2bc9);

#[repr(transparent)]
pub struct IPolicyConfig(pub IUnknown);
impl IPolicyConfig {
    #[allow(non_snake_case, clippy::missing_safety_doc)]
    pub unsafe fn SetDefaultEndpoint<'a, Param0: IntoParam<'a, PWSTR>>(
        &self,
        wszDeviceId: Param0,
        role: ERole,
    ) -> Result<()> {
        (Interface::vtable(self).SetDefaultEndpoint)(
            Interface::as_raw(self),
            wszDeviceId.into_param().abi(),
            mem::transmute(role),
        )
        .ok()
    }
}
impl From<IPolicyConfig> for IUnknown {
    fn from(value: IPolicyConfig) -> Self {
        unsafe { mem::transmute(value) }
    }
}
impl From<&IPolicyConfig> for IUnknown {
    fn from(value: &IPolicyConfig) -> Self {
        From::from(Clone::clone(value))
    }
}
impl<'a> IntoParam<'a, IUnknown> for IPolicyConfig {
    fn into_param(self) -> Param<'a, IUnknown> {
        Param::Owned(unsafe { mem::transmute(self) })
    }
}
impl<'a> IntoParam<'a, IUnknown> for &'a IPolicyConfig {
    fn into_param(self) -> Param<'a, IUnknown> {
        Param::Borrowed(unsafe { mem::transmute(self) })
    }
}
impl Clone for IPolicyConfig {
    fn clone(&self) -> Self {
        Self(self.0.clone())
    }
}
impl PartialEq for IPolicyConfig {
    fn eq(&self, other: &Self) -> bool {
        self.0 == other.0
    }
}
impl Eq for IPolicyConfig {}
impl fmt::Debug for IPolicyConfig {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_tuple("IPolicyConfig").field(&self.0).finish()
    }
}
unsafe impl Interface for IPolicyConfig {
    type Vtable = IPolicyConfig_Vtbl;
    const IID: GUID = GUID::from_u128(0xf8679f50_850a_41cf_9c72_430f290290c8);
}

#[allow(non_snake_case)]
#[repr(C)]
#[doc(hidden)]
pub struct IPolicyConfig_Vtbl {
    pub base__: IUnknownVtbl,
    pub GetMixFormat: unsafe extern "system" fn(
        this: RawPtr,
        pwstrid: PWSTR,
        waveformatex: *mut RawPtr,
    ) -> HRESULT,
    pub GetDeviceFormat: unsafe extern "system" fn(
        this: RawPtr,
        pwstrid: PWSTR,
        param0: i32,
        waveformatex: *mut RawPtr,
    ) -> HRESULT,
    pub ResetDeviceFormat: unsafe extern "system" fn(this: RawPtr, pwstrid: PWSTR) -> HRESULT,
    pub SetDeviceFormat: unsafe extern "system" fn(
        this: RawPtr,
        pwstrid: PWSTR,
        waveformatex0: RawPtr,
        waveformatex1: *mut RawPtr,
    ) -> HRESULT,
    pub GetProcessingPeriod: unsafe extern "system" fn(
        this: RawPtr,
        pwstrid: PWSTR,
        param0: i32,
        param1: RawPtr,
        param1: *mut RawPtr,
    ) -> HRESULT,
    pub SetProcessingPeriod:
        unsafe extern "system" fn(this: RawPtr, pwstrid: PWSTR, param0: RawPtr) -> HRESULT,
    pub GetShareMode: unsafe extern "system" fn(
        this: RawPtr,
        pwstrid: PWSTR,
        devicesharemode: *mut RawPtr,
    ) -> HRESULT,
    pub SetShareMode: unsafe extern "system" fn(
        this: RawPtr,
        pwstrid: PWSTR,
        devicesharemode: *mut RawPtr,
    ) -> HRESULT,
    pub GetPropertyValue: unsafe extern "system" fn(
        this: RawPtr,
        pwstrid: PWSTR,
        key: RawPtr,
        propvariant: *mut RawPtr,
    ) -> HRESULT,
    pub SetPropertyValue: unsafe extern "system" fn(
        this: RawPtr,
        pwstrid: PWSTR,
        key: RawPtr,
        propvariant: *mut RawPtr,
    ) -> HRESULT,
    pub SetDefaultEndpoint:
        unsafe extern "system" fn(this: RawPtr, pwstrid: PWSTR, role: ERole) -> HRESULT,
    pub SetEndpointVisibility:
        unsafe extern "system" fn(this: RawPtr, pwstrid: PWSTR, param0: i32) -> HRESULT,
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
                            if error.code() == NTE_NOT_FOUND {
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
            |global_audio_endpoint| match &*global_audio_endpoint.borrow() {
                Some(audio_endpoint) => {
                    let mut audio_endpoint_volume_pointer: Option<IAudioEndpointVolume> = None;
                    unsafe {
                        audio_endpoint.Activate(
                            &IAudioEndpointVolume::IID,
                            CLSCTX_ALL,
                            std::ptr::null(),
                            &mut audio_endpoint_volume_pointer as *mut _ as *mut _,
                        )
                    }
                    .map_or_else(
                        |error| {
                            if error.code() == E_NOINTERFACE {
                                Ok(None)
                            } else {
                                Err(error)
                            }
                        },
                        |_| match audio_endpoint_volume_pointer {
                            Some(audio_endpoint_volume) => AUDIO_ENDPOINT_VOLUME_CALLBACK
                                .with(|audio_endpoint_volume_callback| unsafe {
                                    audio_endpoint_volume.RegisterControlChangeNotify(
                                        &*audio_endpoint_volume_callback.borrow(),
                                    )
                                })
                                .map(|_| Some(audio_endpoint_volume)),
                            _ => Ok(None),
                        },
                    )
                }
                _ => Ok(None),
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
                    audio_endpoint_volume
                        .UnregisterControlChangeNotify(&*audio_endpoint_volume_callback.borrow())
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

// Function for setting the default audio device
fn set_default_audio_capture_device(device_id: PWSTR) -> Result<()> {
    AUDIO_POLICY_CONFIG.with(|global_audio_policy_config| {
        match &*global_audio_policy_config.borrow() {
            Some(audio_policy_config) => {
                for role in [eConsole, eMultimedia, eCommunications] {
                    unsafe {
                        audio_policy_config.SetDefaultEndpoint(device_id, role)?;
                    }
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

// Load the icons to use from the icon file
fn load_icons(instance: HINSTANCE) -> Result<()> {
    // Load the icons
    let icon_active = unsafe {
        ExtractIconW(
            instance,
            "%SystemRoot%\\System32\\SndVolSSO.dll",
            141u32.wrapping_neg(),
        )
    };
    let icon_muted = unsafe {
        ExtractIconW(
            instance,
            "%SystemRoot%\\System32\\SndVolSSO.dll",
            140u32.wrapping_neg(),
        )
    };

    assert!(!icon_active.is_invalid(), "Active icon is not valid.");
    assert!(!icon_muted.is_invalid(), "Muted icon is not valid.");

    // Replace currently loaded icons
    ICON_ACTIVE.with(|global| match global.replace(Some(icon_active)) {
        Some(old_icon) => unsafe { DestroyIcon(old_icon).ok() },
        _ => Ok(()),
    })?;
    ICON_MUTED.with(|global| match global.replace(Some(icon_muted)) {
        Some(old_icon) => unsafe { DestroyIcon(old_icon).ok() },
        _ => Ok(()),
    })?;
    Ok(())
}

// Retrieves the microphone status and updates the icon and tooltip
fn update_icon_data() -> Result<()> {
    NOTIFY_ICON_DATA.with(|global_notify_icon_data| {
        let notify_icon_data = &mut *global_notify_icon_data.borrow_mut();
        let max_text_len = notify_icon_data.szTip.len() - 1;
        let mut set_icon_data =
            |icon: &'static LocalKey<RefCell<Option<HICON>>>,
             text: &mut dyn Iterator<Item = u16>| {
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
                    .zip(text.take(max_text_len).chain(std::iter::repeat(0)))
                    .for_each(|(ptr, chr)| *ptr = chr);
                if notify_icon_data.szTip[0] == 0 {
                    notify_icon_data.uFlags &= !(NIF_TIP | NIF_SHOWTIP);
                } else {
                    notify_icon_data.uFlags |= NIF_TIP | NIF_SHOWTIP;
                }
            };
        AUDIO_DEFAULT_ENDPOINT.with(|global_audio_endpoint| {
            match &*global_audio_endpoint.borrow() {
                Some(audio_endpoint) => {
                    let mut device_name_buffer = unsafe {
                        audio_endpoint
                            .OpenPropertyStore(STGM_READ)?
                            .GetValue(&PKEY_Device_FriendlyName)?
                            .Anonymous
                            .Anonymous
                            .Anonymous
                            .pwszVal
                            .0
                    };
                    let device_name_prefix = std::iter::from_fn(move || {
                        let chr;
                        unsafe {
                            chr = if *device_name_buffer != 0 {
                                Some(*device_name_buffer)
                            } else {
                                None
                            };
                            device_name_buffer = device_name_buffer.add(1);
                        }
                        chr
                    })
                    .chain(": ".encode_utf16());
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

                            set_icon_data(
                                if volume.is_some() {
                                    &ICON_ACTIVE
                                } else {
                                    &ICON_MUTED
                                },
                                &mut device_name_prefix.chain(
                                    match volume {
                                        Some(volume_text) => volume_text,
                                        _ => LABEL_MUTED.to_owned(),
                                    }
                                    .encode_utf16(),
                                ),
                            );
                            Ok(())
                        }
                        _ => {
                            set_icon_data(
                                &ICON_MUTED,
                                &mut device_name_prefix.chain(LABEL_VOLUME_UNKNOWN.encode_utf16()),
                            );
                            Ok(())
                        }
                    })
                }
                _ => {
                    set_icon_data(&ICON_MUTED, &mut LABEL_NO_DEFAULT_DEVICE.encode_utf16());
                    Ok(())
                }
            }
        })
    })
}

// Add the notify icon for when it does not already exists
fn add_notify_icon() -> Result<()> {
    update_icon_data()?;
    NOTIFY_ICON_DATA.with(|global| {
        let notify_icon_data = &*global.borrow();
        unsafe {
            Shell_NotifyIconW(NIM_ADD, notify_icon_data).ok()?;
            Shell_NotifyIconW(NIM_SETVERSION, notify_icon_data).ok()?;
        }
        Ok(())
    })
}

// Update the existing notify icon
fn update_notify_icon() -> Result<()> {
    update_icon_data()?;
    NOTIFY_ICON_DATA.with(|global| {
        let notify_icon_data = &*global.borrow();
        unsafe {
            Shell_NotifyIconW(NIM_MODIFY, notify_icon_data).ok()?;
        }
        Ok(())
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
                                )
                                .ok()?;
                            }
                            unsafe {
                                AppendMenuW(
                                    menu,
                                    MF_SEPARATOR,
                                    IDM_SEPARATOR as usize,
                                    PCWSTR::default(),
                                )
                                .ok()?;
                            }
                            unsafe {
                                AppendMenuW(
                                    menu,
                                    MF_ENABLED | MF_STRING,
                                    IDM_EXIT as usize,
                                    LABEL_EXIT,
                                )
                                .ok()?;
                            }
                            menu
                        }));
                    }

                    let menu = &*global_menu.borrow();
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
                            unsafe { RemoveMenu(menu, IDM_NO_ENDPOINTS as u32, MF_BYCOMMAND) }
                                .ok()?;
                        }
                        for i in 0..devices_count {
                            let device = unsafe { devices.Item(i as u32) }?;
                            let device_id = unsafe { device.GetId() }?;
                            let device_name_buffer = unsafe {
                                device
                                    .OpenPropertyStore(STGM_READ)?
                                    .GetValue(&PKEY_Device_FriendlyName)?
                                    .Anonymous
                                    .Anonymous
                                    .Anonymous
                                    .pwszVal
                                    .0
                            };
                            let device_name = PWSTR(device_name_buffer);
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
                                            CoTaskMemFree(
                                                menu_audio_endpoints.remove(i).0 as *const c_void,
                                            );
                                            RemoveMenu(menu, i as u32, MF_BYPOSITION).ok()?;
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
                                        )
                                        .ok()?;
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
                                    )
                                    .ok()?;
                                }
                            }
                        }
                        let add_no_devices_label =
                            devices_count == 0 && menu_audio_endpoints.len() > 0;
                        while menu_audio_endpoints.len() > devices_count {
                            unsafe {
                                CoTaskMemFree(
                                    menu_audio_endpoints.pop().unwrap().0 as *const c_void,
                                );
                                RemoveMenu(menu, devices_count as u32, MF_BYPOSITION)
                            }
                            .ok()?;
                        }
                        if add_no_devices_label {
                            let mut no_devices_label_buffer =
                                LABEL_NO_CAPTURE_DEVICES.encode_utf16().collect::<Vec<_>>();
                            let no_devices_label = PWSTR(no_devices_label_buffer.as_mut_ptr());
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
                                        dwTypeData: no_devices_label,
                                        ..Default::default()
                                    },
                                )
                                .ok()?;
                            }
                        }
                        Ok(())
                    })?;

                    if let Some(id) = default_endpoint_id {
                        unsafe { CoTaskMemFree(id.0 as *const c_void) }
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
            SetForegroundWindow(window);
            TrackPopupMenuEx(
                *menu.borrow(),
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
                std::ptr::null(),
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
            let wm_taskbar_created = unsafe { RegisterWindowMessageW("TaskbarCreated") };
            assert!(wm_taskbar_created != 0);
            WM_TASKBAR_CREATED.with(|global| {
                global.replace(Some(wm_taskbar_created));
            });
            LRESULT(0)
        }
        WM_DPICHANGED => {
            let instance = unsafe { GetModuleHandleW(None) }.unwrap();
            assert!(instance.0 != 0);
            load_icons(instance).unwrap();
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
                    wparam.0 as i32 & 0xffff,
                    (wparam.0 as i32 >> 16) & 0xffff,
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
                .and_then(|()| unsafe { DrawMenuBar(window).ok() })
                .ok();
            LRESULT(0)
        }
        WM_APP_CALLBACK_VOLUME => {
            // Audio endpoint volume/mute has changed
            update_notify_icon().ok();
            LRESULT(0)
        }
        WM_COMMAND => {
            match (wparam.0 as u32 & 0xffff) as u16 {
                IDM_EXIT => unsafe {
                    DestroyWindow(window);
                },
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
                DestroyWindow(window);
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

fn main() -> Result<()> {
    let instance = unsafe { GetModuleHandleW(None) }?;
    assert!(instance.0 != 0);

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
            std::ptr::null(),
        )
    };
    assert!(window.0 != 0);

    // Initialize COM runtime
    unsafe {
        CoInitializeEx(std::ptr::null(), COINIT_MULTITHREADED)?;
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
            .RegisterEndpointNotificationCallback(&*audio_endpoint_callback.borrow())
    })?;
    AUDIO_ENDPOINT_ENUMERATOR.with(|global| {
        global.replace(Some(audio_endpoint_enumerator));
    });

    // Get the audio endpoint and volume objects
    init_audio_endpoint()?;
    init_audio_endpoint_volume()?;

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
        global.replace(notify_icon_data);
    });

    load_icons(instance)?;
    add_notify_icon()?;

    // Message pump
    let mut message = MSG::default();
    while unsafe { GetMessageW(&mut message, HWND(0), 0, 0) }.into() {
        unsafe {
            DispatchMessageW(&message);
        }
    }

    // Remove the notification icon
    unsafe {
        Shell_NotifyIconW(NIM_DELETE, &notify_icon_data).ok()?;
    }

    // Release the COM objects
    deinit_audio_endpoint();
    AUDIO_POLICY_CONFIG.with(|global| {
        global.replace(None);
    });
    AUDIO_ENDPOINT_ENUMERATOR.with(|global| {
        if let Some(audio_endpoint_enumerator) = global.replace(None) {
            AUDIO_ENDPOINT_CALLBACK.with(|audio_endpoint_callback| unsafe {
                audio_endpoint_enumerator
                    .UnregisterEndpointNotificationCallback(&*audio_endpoint_callback.borrow())
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
        Some(old_icon) => unsafe { DestroyIcon(old_icon).ok() },
        _ => Ok(()),
    })?;
    ICON_MUTED.with(|global| match global.replace(None) {
        Some(old_icon) => unsafe { DestroyIcon(old_icon).ok() },
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
        UnregisterClassW(window_class_name, instance).ok()?;
    }

    if message.wParam.0 == 0 {
        Ok(())
    } else {
        Err(Error::from_win32())
    }
}
