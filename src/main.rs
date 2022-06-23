// Simple microphone mute indicator for Windows
// Author: Daniël van de Giessen <daniel@dvdgiessen.nl>

#![windows_subsystem = "windows"]

use std::cell::RefCell;

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
thread_local!(static ICON_ACTIVE: RefCell<HICON> = RefCell::new(Default::default()));
thread_local!(static ICON_MUTED: RefCell<HICON> = RefCell::new(Default::default()));

// Notify icon data registered to show in the notification tray
thread_local!(static NOTIFY_ICON_DATA: RefCell<NOTIFYICONDATAW> = RefCell::new(Default::default()));

// Context menu shown when right-clicking the notify icon
thread_local!(static MENU: RefCell<HMENU> = RefCell::new(Default::default()));

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
    let audio_endpoint: IMMDevice =
        AUDIO_ENDPOINT_ENUMERATOR.with(|global| match &*global.borrow() {
            Some(audio_endpoint_enumerator) => unsafe {
                audio_endpoint_enumerator.GetDefaultAudioEndpoint(eCapture, eCommunications)
            },
            _ => panic!(),
        })?;
    AUDIO_ENDPOINT.with(|global| {
        global.replace(Some(audio_endpoint));
    });
    Ok(())
}

fn init_audio_endpoint_volume() -> Result<()> {
    deinit_audio_endpoint_volume();
    let audio_endpoint_volume: IAudioEndpointVolume =
        AUDIO_ENDPOINT.with(|global| match &*global.borrow() {
            Some(audio_endpoint) => {
                let mut audio_endpoint_volume_pointer = None;
                unsafe {
                    audio_endpoint.Activate(
                        &IAudioEndpointVolume::IID,
                        CLSCTX_ALL,
                        std::ptr::null(),
                        &mut audio_endpoint_volume_pointer as *mut _ as *mut _,
                    )?;
                }
                audio_endpoint_volume_pointer.ok_or(Error::OK)
            }
            _ => panic!(),
        })?;
    AUDIO_ENDPOINT_VOLUME_CALLBACK.with(|audio_endpoint_volume_callback| unsafe {
        audio_endpoint_volume.RegisterControlChangeNotify(&*audio_endpoint_volume_callback.borrow())
    })?;
    AUDIO_ENDPOINT_VOLUME.with(|global| {
        global.replace(Some(audio_endpoint_volume));
    });
    Ok(())
}

// Audio deinitialization functions
fn deinit_audio_endpoint_volume() {
    AUDIO_ENDPOINT_VOLUME.with(|global| {
        match &*global.borrow() {
            Some(audio_endpoint_volume) => {
                AUDIO_ENDPOINT_VOLUME_CALLBACK.with(|audio_endpoint_volume_callback| unsafe {
                    audio_endpoint_volume
                        .UnregisterControlChangeNotify(&*audio_endpoint_volume_callback.borrow())
                        .unwrap()
                });
            }
            _ => (),
        }
        drop(global.replace(None));
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
        _ => panic!(),
    })
}

// Retrieves the microphone status and updates the icon and tooltip
fn update_icon_data() -> Result<()> {
    AUDIO_ENDPOINT.with(
        |global_audio_endpoint| match &*global_audio_endpoint.borrow() {
            Some(audio_endpoint) => AUDIO_ENDPOINT_VOLUME.with(|global| match &*global.borrow() {
                Some(audio_endpoint_volume) => NOTIFY_ICON_DATA.with(|global_notify_icon_data| {
                    let notify_icon_data = &mut *global_notify_icon_data.borrow_mut();
                    let tooltip_text: Vec<u16>;
                    let mut device_name = unsafe {
                        audio_endpoint
                            .OpenPropertyStore(STGM_READ)?
                            .GetValue(&PKEY_Device_FriendlyName)?
                            .Anonymous
                            .Anonymous
                            .Anonymous
                            .pwszVal
                            .0
                    };
                    if unsafe { audio_endpoint_volume.GetMute() }?.into() {
                        ICON_MUTED.with(|icon_muted| notify_icon_data.hIcon = *icon_muted.borrow());
                        tooltip_text = ": muted".encode_utf16().collect();
                    } else {
                        ICON_ACTIVE
                            .with(|icon_active| notify_icon_data.hIcon = *icon_active.borrow());
                        tooltip_text = format!(
                            ": {:.0}%",
                            100f32 * unsafe { audio_endpoint_volume.GetMasterVolumeLevelScalar() }?
                        )
                        .encode_utf16()
                        .collect();
                    }
                    notify_icon_data
                        .szTip
                        .iter_mut()
                        .zip(
                            std::iter::from_fn(move || {
                                let chr;
                                unsafe {
                                    chr = if *device_name != 0 {
                                        Some(*device_name)
                                    } else {
                                        None
                                    };
                                    device_name = device_name.add(1);
                                }
                                chr
                            })
                            .chain(tooltip_text)
                            .chain(std::iter::repeat(0)),
                        )
                        .for_each(|(ptr, chr)| *ptr = chr);

                    Ok(())
                }),
                _ => panic!(),
            }),
            _ => panic!(),
        },
    )
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
        WM_APP_NOTIFYICON => {
            match lparam.0 as u32 & 0xffff {
                NIN_SELECT => {
                    toggle_mute().unwrap();
                    update_notify_icon().unwrap();
                }
                WM_CONTEXTMENU => {
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
                                wparam.0 as i32 & 0xffff,
                                (wparam.0 as i32 >> 16) & 0xffff,
                                window,
                                std::ptr::null(),
                            );
                        }
                    });
                }
                _ => (),
            }
            LRESULT(0)
        }
        WM_APP_CALLBACK_ENDPOINT => {
            // Audio endpoint has changed
            init_audio_endpoint().unwrap();
            init_audio_endpoint_volume().unwrap();
            update_notify_icon().unwrap();
            LRESULT(0)
        }
        WM_APP_CALLBACK_VOLUME => {
            // Audio endpoint volume/mute has changed
            update_notify_icon().unwrap();
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
                add_notify_icon().unwrap();
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

    // Create a menu
    let menu = unsafe { CreatePopupMenu() }?;
    unsafe {
        AppendMenuW(menu, MF_ENABLED | MF_STRING, IDM_EXIT, "E&xit\0").ok()?;
    }
    MENU.with(|global| {
        global.replace(menu);
    });

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

    ICON_ACTIVE.with(|global| {
        global.replace(icon_active);
    });
    ICON_MUTED.with(|global| {
        global.replace(icon_muted);
    });

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
        uFlags: NIF_ICON | NIF_MESSAGE | NIF_TIP | NIF_SHOWTIP,
        uCallbackMessage: WM_APP_NOTIFYICON,
        Anonymous: NOTIFYICONDATAW_0 {
            uVersion: NOTIFYICON_VERSION_4,
        },
        ..Default::default()
    };
    NOTIFY_ICON_DATA.with(|global| {
        global.replace(notify_icon_data);
    });

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
    AUDIO_ENDPOINT_ENUMERATOR.with(|global| {
        match &*global.borrow() {
            Some(audio_endpoint_enumerator) => {
                AUDIO_ENDPOINT_CALLBACK.with(|audio_endpoint_callback| unsafe {
                    audio_endpoint_enumerator
                        .UnregisterEndpointNotificationCallback(&*audio_endpoint_callback.borrow())
                        .unwrap()
                });
            }
            _ => (),
        }
        drop(global.replace(None));
    });

    // Deinitialize COM runtime
    unsafe {
        CoUninitialize();
    }

    // Deallocate icons
    unsafe {
        DestroyIcon(icon_active).ok()?;
        DestroyIcon(icon_muted).ok()?;
    }

    // Destroy the menu
    unsafe {
        DestroyMenu(menu);
    }

    // Unregister the window class
    unsafe {
        UnregisterClassW("MicrophoneMuteIndicator", instance).ok()?;
    }

    if message.wParam.0 == 0 {
        Ok(())
    } else {
        Err(Error::from_win32())
    }
}
