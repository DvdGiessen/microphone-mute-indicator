Simple systray microphone mute indicator for Windows.

Build options to create a nice small executable:
```
cargo +nightly build -Z build-std=std,panic_abort -Z build-std-features=panic_immediate_abort --target x86_64-pc-windows-msvc --release
```

