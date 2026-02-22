# Push-to-Talk Dictation (Groq Whisper)

Hold **Scroll Lock** to record voice, release to transcribe and paste text into the active window.

## Quick Start

1. Copy `start-dictation.vbs.example` to `start-dictation.vbs`
2. Edit `start-dictation.vbs` - set your Groq API key and path to `dictation.exe`
3. Run `start-dictation.vbs` (launches hidden, no terminal window)

## Autostart

Copy your `start-dictation.vbs` to `shell:startup` folder. It sets `GROQ_API_KEY` in the process scope only and launches `dictation.exe` with no visible window on every logon.

## How It Works

1. **Hold Scroll Lock** - high beep confirms recording started
2. **Speak** in Polish or English (auto-detected by Whisper)
3. **Release Scroll Lock** - low beep confirms recording stopped
4. **Text appears** in the active window after 1-3 seconds

## Architecture

```
Keyboard Hook (WH_KEYBOARD_LL) - event-driven, 0 CPU
    |
    v  (sets flag, returns instantly)
Forms.Timer (30ms poll on UI thread)
    |
    v  (start/stop waveIn recording)
ThreadPool -> Groq Whisper API (HTTP POST, multipart/form-data)
    |
    v  (STA thread)
Clipboard.SetText + SendInput(Ctrl+V) -> active window
    |
    v
Restore previous clipboard
```

### Why this architecture:

- **Hook callback must return fast** (<1ms) - Windows removes hooks that block >300ms
- **Hook only sets a boolean flag** - timer picks it up on the next tick
- **waveIn API** with single 60s buffer, no callback - simple and crash-free
- **Transcription on ThreadPool** - doesn't block the message pump
- **Clipboard + Ctrl+V** for pasting - handles Unicode (Polish characters) correctly
- **STA thread for clipboard** - required by Windows clipboard API

## Configuration

Edit constants at the top of `Program.cs` and recompile:

| Constant | Default | Description |
|----------|---------|-------------|
| `WHISPER_MODEL` | `whisper-large-v3-turbo` | Whisper model to use |
| `HOTKEY_VK` | `0x91` (Scroll Lock) | Virtual key code for push-to-talk |
| `SAMPLE_RATE` | `16000` | Recording sample rate (Hz) |
| `MAX_RECORD_SEC` | `60` | Maximum recording length |

API key is read from `GROQ_API_KEY` environment variable at startup.

### Alternative hotkeys:

| Key | VK Code |
|-----|---------|
| Scroll Lock | `0x91` |
| Pause/Break | `0x13` |
| F24 | `0x87` |
| Right Ctrl | `0xA3` |

## Building

Requires only the .NET Framework 4.x compiler (included with Windows):

```
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /target:exe /out:dictation.exe /reference:System.Windows.Forms.dll /platform:x64 Program.cs
```

## Files

| File | Description |
|------|-------------|
| `dictation.exe` | Compiled executable |
| `Program.cs` | Source code |
| `start-dictation.vbs.example` | Template for silent launcher (sets env var, no window) |
| `dictation.log` | Runtime log (created on first run) |
| `dictation_recording.wav` | Last recording (created on first run) |

## Troubleshooting

Check `dictation.log` for diagnostics.

| Problem | Solution |
|---------|----------|
| No beep on key press | Scroll Lock might be captured by another app |
| "waveInOpen failed" | Check Windows microphone permissions (Settings > Privacy > Microphone) |
| "ERROR API" | Check API key, internet connection |
| Text not appearing | Make sure target window supports Ctrl+V paste |
| Polish chars broken | Should not happen (clipboard+Ctrl+V is Unicode-safe). File a bug. |

## Fullscreen Games

Works in fullscreen games. `WH_KEYBOARD_LL` hooks are not injected into the game process - Windows does a context switch to the hook's process. Scroll Lock is swallowed so the LED won't toggle.
