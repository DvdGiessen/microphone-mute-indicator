# Microphone Mute Indicator

A tiny program for showing the microphone mute status in the Windows taskbar
notification area.

## Installation

Download the latest executable from the
[releases page](https://github.com/DvdGiessen/microphone-mute-indicator/releases),
and run it.

If you want the program to run automatically at startup, you can put it in your
Startup directory (`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`).

## Usage

The program shows the status of the current default communication audio capture
device.

- The notification icon indicates whether the device is currently muted.
- Hover the notification icon to display the name and volume of the device.
- Left-click the notification icon to mute and unmute the device.
- Right-click the notification icon to select a different capture device, or to
  exit the program.

## Building

The build command to create a nice small executable:

```sh
cargo +nightly build -Z build-std=std,panic_abort -Z build-std-features=panic_immediate_abort --target x86_64-pc-windows-msvc --release
```

## Issues

If you have any issues with Microphone Mute Indicator, first check the issue
tracker to see whether it was already reported by someone else. If not, go ahead
and create a new issue. Try to include as much information (version of the
program, version of Windows, steps to reproduce) as possible.

## License

Microphone Mute Indicator is freely distributable under the terms of the MIT
license.
