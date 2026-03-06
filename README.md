# FreeSpeech

![FreeSpeech Logo](assets/icon-512.png)

FreeSpeech is a standalone Windows text-to-speech app focused on one job: quickly reading selected text out loud with high-quality Microsoft neural voices (via `edge-tts`).

It provides a floating command bar UI, tray integration, browser right-click support for Chrome, and practical controls for voice tuning, logging, and startup behavior.

## Why FreeSpeech

Most simple TTS tools are either expensive, low-quality, or inconvenient in daily use.  
FreeSpeech is designed to stay lightweight and fast while still giving enough control for real workflows (classes, research, long reading sessions, and accessibility support).

## Highlights (From In-App About)

I made this application because I needed an easy way to read text from my online classes out loud. Every option I found either cost money, did not work properly, or used robotic voices.

This sucks and was completely unacceptable.

FreeSpeech is 100% free and open source for further development, thanks to several external projects:

- Edge-TTS  
Used for all text-to-speech synthesis with Microsoft's neural voices.  
https://github.com/rany2/edge-tts
- CustomTkinter  
Used for building the FreeSpeech desktop user interface.  
https://github.com/TomSchimansky/CustomTkinter
- Silent_Chrome  
Used for silent Chrome extension installation support.  
https://github.com/asaurusrex/Silent_Chrome?ref=blog.sunggwanchoi.com

## How It Works

1. FreeSpeech captures text from one of three sources:
   - highlighted text (`Read Selection`)
   - clipboard (`Speak Clipboard`)
   - browser extension requests (Chrome context menu)
2. Text is queued in the app and synthesized through `edge-tts`.
3. Audio is played immediately and can be interrupted with `Stop`.
4. If enabled, generated audio is also saved as `.mp3` files to your selected folder.

## Main UI

The primary floating window includes:

- `Read Selection`
- `Test Voice`
- `Speak Clipboard`
- `Browser Right-Click Support`
- Voice controls:
  - language-region filter
  - voice dropdown
  - rate, pitch, volume
- top-right toggles:
  - `Always on Top`
  - `Dark Mode`
- `Show Advanced Settings`
- `About`

## Advanced Settings

The Advanced Settings dialog includes:

- `Capture Settings`
  - capture delay (ms)
  - max chars per read
- `Scaling`
  - UI scaling from compact to large presets
- `Generated Speech Output`
  - switch to save generated speech as MP3
  - folder picker
- `Start on Boot`
  - Windows startup registration toggle
- utility actions
  - save settings
  - open config folder
  - open HTML error log

## Browser Right-Click Support (Chrome)

FreeSpeech includes a Chrome extension flow from the app:

- `Install Chrome Right-Click Support` for automatic install attempt.
- `Manual Install` with step-by-step instructions and quick-open extension folder.

The extension adds a context-menu action for selected text and sends that text to FreeSpeech through a local app bridge.  
If FreeSpeech is not running, the extension warns the user to launch the app.

## Tray Behavior

- Closing the main window minimizes FreeSpeech to tray.
- Tray notification confirms: `Minimized to taskbar`.
- Left-click tray icon toggles show/hide.
- Right-click tray menu includes:
  - `Show`
  - `Stop Speech` (visible only while speaking)
  - `Read Selection`
  - `Test Voice`
  - `Speak Clipboard`
  - `Exit`

## File Input Support

FreeSpeech can read full documents from:

- opening `.txt` or `.docx` with FreeSpeech
- dragging and dropping supported files into the app window

## Theming and Branding

- UI framework: `CustomTkinter`
- active theme file: `themes/red.json`
- app icon assets:
  - main: `assets/icon-512.png` / `assets/icon-512.ico`
  - maskable: `assets/icon-512-maskable.png`
  - tray/small icon: `assets/favicon.ico`

## Versioning

Version is intentionally manual (no automatic stamping):

- edit `freespeech/version.py`
- update `APP_VERSION = "2.0"` as needed

The About dialog reads from that constant and displays it at the bottom-right.

## Installation (Development)

```powershell
cd "<project-folder>"
python -m pip install -r requirements.txt
python -m freespeech
```

## Build Standalone EXE

```powershell
cd "<project-folder>"
.\build_standalone.ps1
```

Build output:

- `dist/FreeSpeech.exe`

## Project Structure

- `freespeech/main.py` - main app UI, tray behavior, dialogs, browser bridge
- `freespeech/speech_service.py` - speech queue and playback flow
- `freespeech/backends/` - speech backend adapter(s), voice handling
- `freespeech/config.py` - persisted app settings
- `freespeech/version.py` - manual version constant
- `themes/red.json` - CustomTkinter color/theme definition
- `tools/silent_chrome_windows.py` - Chrome silent-install helper integration

## Dependencies

Runtime Python packages:

- `edge-tts`
- `customtkinter`
- `pystray`
- `Pillow`
- `pyperclip`
- `pynput`

## Troubleshooting

- Speech does not start:
  - verify internet access (required by Edge TTS service)
  - verify text is actually captured (try `Speak Clipboard`)
- Build fails with `Access is denied` for `dist/FreeSpeech.exe`:
  - close any running FreeSpeech instance
  - rebuild
- Chrome context menu action does nothing:
  - confirm FreeSpeech is running
  - reinstall extension from Browser Right-Click Support dialog
  - use Manual Install path if automatic install is blocked by policy

## Credits

- Edge-TTS: https://github.com/rany2/edge-tts
- CustomTkinter: https://github.com/TomSchimansky/CustomTkinter
- Silent_Chrome: https://github.com/asaurusrex/Silent_Chrome?ref=blog.sunggwanchoi.com
