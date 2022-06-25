// Simple microphone mute indicator for Windows
// Author: Daniël van de Giessen <daniel@dvdgiessen.nl>

#![windows_subsystem = "windows"]

use std::{cell::RefCell, thread::LocalKey};

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

const IDM_EXIT: usize = 100;

// Message received when the taskbar is (re)created
thread_local!(static WM_TASKBAR_CREATED: RefCell<Option<u32>> = RefCell::new(None));

// COM objects for audio endpoints
thread_local!(static AUDIO_ENDPOINT_ENUMERATOR: RefCell<Option<IMMDeviceEnumerator>> = RefCell::new(None));
thread_local!(static AUDIO_ENDPOINT: RefCell<Option<IMMDevice>> = RefCell::new(None));
thread_local!(static AUDIO_ENDPOINT_VOLUME: RefCell<Option<IAudioEndpointVolume>> = RefCell::new(None));

// Icons for active and muted states
thread_local!(static ICON_ACTIVE: RefCell<Option<HICON>> = RefCell::new(None));
thread_local!(static ICON_MUTED: RefCell<Option<HICON>> = RefCell::new(None));

// Notify icon data registered to show in the notification tray
thread_local!(static NOTIFY_ICON_DATA: RefCell<NOTIFYICONDATAW> = RefCell::new(Default::default()));

// Context menu shown when right-clicking the notify icon
thread_local!(static MENU: RefCell<Option<HMENU>> = RefCell::new(None));

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

// Audio initialization functions
fn init_audio_endpoint() -> Result<()> {
    deinit_audio_endpoint();
    AUDIO_ENDPOINT.with(|global| {
        global.replace(AUDIO_ENDPOINT_ENUMERATOR.with(|global| {
            match &*global.borrow() {
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
        })?);
        Ok(())
    })
}

fn init_audio_endpoint_volume() -> Result<()> {
    deinit_audio_endpoint_volume();
    AUDIO_ENDPOINT_VOLUME.with(|global_audio_endpoint_volume| {
        global_audio_endpoint_volume.replace(AUDIO_ENDPOINT.with(|global_audio_endpoint| {
            match &*global_audio_endpoint.borrow() {
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
            }
        })?);
        Ok(())
    })
}

// Audio deinitialization functions
fn deinit_audio_endpoint_volume() {
    AUDIO_ENDPOINT_VOLUME.with(|global| match global.replace(None) {
        Some(audio_endpoint_volume) => {
            AUDIO_ENDPOINT_VOLUME_CALLBACK
                .with(|audio_endpoint_volume_callback| unsafe {
                    audio_endpoint_volume
                        .UnregisterControlChangeNotify(&*audio_endpoint_volume_callback.borrow())
                })
                .ok();
        }
        _ => (),
    });
}

fn deinit_audio_endpoint() {
    deinit_audio_endpoint_volume();
    AUDIO_ENDPOINT.with(|global| {
        drop(global.replace(None));
    });
}

// Function for toggling mute, used when clicking the icon
fn toggle_mute() -> Result<()> {
    AUDIO_ENDPOINT_VOLUME.with(|global| match &*global.borrow() {
        Some(audio_endpoint_volume) => unsafe {
            audio_endpoint_volume.SetMute(
                !audio_endpoint_volume.GetMute()?.as_bool(),
                std::ptr::null(),
            )
        },
        _ => Ok(()),
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
        AUDIO_ENDPOINT.with(
            |global_audio_endpoint| match &*global_audio_endpoint.borrow() {
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
                    AUDIO_ENDPOINT_VOLUME.with(|global| match &*global.borrow() {
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
                                        _ => "muted".to_owned(),
                                    }
                                    .encode_utf16(),
                                ),
                            );
                            Ok(())
                        }
                        _ => {
                            set_icon_data(
                                &ICON_MUTED,
                                &mut device_name_prefix.chain("volume unknown!".encode_utf16()),
                            );
                            Ok(())
                        }
                    })
                }
                _ => {
                    set_icon_data(
                        &ICON_MUTED,
                        &mut "No default communications audio input found!".encode_utf16(),
                    );
                    Ok(())
                }
            },
        )
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

// Update the menu
fn update_menu() -> Result<()> {
    MENU.with(|global| {
        let menu = unsafe { CreatePopupMenu() }?;
        unsafe {
            AppendMenuW(menu, MF_ENABLED | MF_STRING, IDM_EXIT, "E&xit\0").ok()?;
        }
        match global.replace(Some(menu)) {
            Some(old_menu) => unsafe { DestroyMenu(old_menu).ok() },
            _ => Ok(()),
        }
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
            LRESULT(0)
        }
        WM_APP_CALLBACK_VOLUME => {
            // Audio endpoint volume/mute has changed
            update_notify_icon().ok();
            LRESULT(0)
        }
        WM_COMMAND => {
            match wparam.0 & 0xffff {
                IDM_EXIT => unsafe {
                    DestroyWindow(window);
                },
                _ => (),
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
    AUDIO_ENDPOINT_ENUMERATOR.with(|global| match global.replace(None) {
        Some(audio_endpoint_enumerator) => {
            AUDIO_ENDPOINT_CALLBACK.with(|audio_endpoint_callback| unsafe {
                audio_endpoint_enumerator
                    .UnregisterEndpointNotificationCallback(&*audio_endpoint_callback.borrow())
                    .unwrap()
            });
        }
        _ => (),
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
    MENU.with(|global| match global.replace(None) {
        Some(menu) => unsafe {
            DestroyMenu(menu).unwrap();
        },
        _ => (),
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
