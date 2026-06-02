<div align="center">
  <img src="resources/logo.png" width="150" alt="Riff logo" />
</div>

<h1 align="center">Riff</h1>

<p align="center">
  <b>A single-file VBA audio engine for Microsoft Office with WASAPI playback, Media Foundation decoding, real-time DSP, oscillators, buses, meters, and WAV export.</b>
</p>

<p align="center">
  <img src="https://github.com/vbacollective/riff/actions/workflows/ci.yml/badge.svg" alt="CI" />
  <img src="https://github.com/vbacollective/riff/actions/workflows/release-assets.yml/badge.svg" alt="Release" />
  <img src="https://img.shields.io/badge/language-VBA-867DB1.svg" alt="Language" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D6.svg" alt="Platform" />
  <img src="https://img.shields.io/badge/arch-32%20%26%2064--bit-green.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/WASAPI-Shared%20Mode-blue.svg" alt="WASAPI" />
  <img src="https://img.shields.io/badge/Media%20Foundation-Decoding-orange.svg" alt="Media Foundation" />
  <img src="https://img.shields.io/badge/DSP-Studio%20Pipeline-critical.svg" alt="DSP" />
  <img src="https://img.shields.io/badge/dependencies-none-success.svg" alt="Dependencies" />
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" />
</p>

> [!IMPORTANT]
> Riff is Windows-only. It calls native Windows APIs including `kernel32`, `user32`, `ole32`, `oleaut32`, `winmm`, `mfplat`, `mfreadwrite`, and `shlwapi`.

## What Riff Does

Riff brings real audio playback and synthesis to VBA without DLLs, installers, ActiveX controls, typelibs, or external references. Import one `.bas` file into Excel, Word, PowerPoint, Access, Outlook, or another VBA host, then use simple global functions to load sounds, play voices, route buses, shape effects, render oscillators, and export WAV files.

The engine uses:

- WASAPI Shared Mode for output.
- Media Foundation for decoding formats supported by Windows.
- A runtime x86/x64 timer thunk for the audio callback.
- Direct COM vtable dispatch for WASAPI and Media Foundation interfaces.
- Static buffer and voice pools to avoid VBA object lifetime issues in the audio path.

## Repository Layout

| Path | Purpose |
|---|---|
| [`package/`](package/) | The importable Riff module. Import `package/Riff.bas` into your VBA project. |
| [`examples/`](examples/) | Practical example modules and copy-ready snippets. |
| [`docs/`](docs/) | API reference and internal architecture documentation. |
| [`resources/`](resources/) | Project visual assets used by documentation. |

## Install

1. Download or clone this repository.
2. Open the VBA editor in your Office host with `Alt+F11`.
3. Choose **File > Import File...**.
4. Select `package/Riff.bas`.
5. Save the workbook, document, database, or presentation as a macro-enabled file.

No **Tools > References** entries are required.

## Minimal Playback

This is the smallest complete pattern: start the engine, load a file, play it, and shut the engine down when the host closes.

```vb
Option Explicit

Public Sub PlayOneSound()
    If Not RiffOpen() Then
        MsgBox "Riff could not initialize the audio device.", vbCritical
        Exit Sub
    End If

    Dim bufferId As Long
    bufferId = RiffLoad("C:\Audio\click.wav")

    If bufferId < 0 Then
        MsgBox "The audio file could not be loaded.", vbExclamation
        Exit Sub
    End If

    RiffPlay bufferId
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RiffClose
End Sub
```

## Usage Examples

### Loop background music and play sound effects

Use one loaded buffer for background music and another for short effects. The same buffer can be played many times without decoding it again.

```vb
Option Explicit

Private musicBuffer As Long
Private explosionBuffer As Long
Private musicVoice As Long
Private gameAudioReady As Boolean

Public Sub StartGameAudio()
    If Not RiffOpen() Then Exit Sub

    gameAudioReady = False
    musicBuffer = RiffLoad("C:\GameAudio\theme.mp3")
    explosionBuffer = RiffLoad("C:\GameAudio\explosion.wav")

    If musicBuffer < 0 Or explosionBuffer < 0 Then
        MsgBox "One or more audio assets failed to load.", vbExclamation
        Exit Sub
    End If

    musicVoice = RiffPlay(musicBuffer)
    If musicVoice >= 0 Then
        RiffVoiceLoop(musicVoice) = True
        RiffVoiceVolume(musicVoice) = 0.35
    End If

    gameAudioReady = True
End Sub

Public Sub PlayExplosion()
    If Not gameAudioReady Then Exit Sub

    Dim voiceId As Long
    voiceId = RiffPlay(explosionBuffer)

    If voiceId >= 0 Then
        RiffVoiceVolume(voiceId) = 0.9
        RiffVoicePan(voiceId) = 0.15
        RiffVoiceLowPass(voiceId) = 0.85
    End If
End Sub
```

### Mix with audio buses

Buses are global volume groups. They are useful when your project needs independent music, sound-effect, voice, and UI levels.

```vb
Public Sub ConfigureBuses(ByVal musicVoice As Long, ByVal sfxVoice As Long)
    RiffVoiceBus(musicVoice) = RiffBusMusic
    RiffVoiceBus(sfxVoice) = RiffBusSfx

    RiffBusVolume(RiffBusMusic) = 0.45
    RiffBusVolume(RiffBusSfx) = 1!
End Sub
```

### Inspect failures

Functions that return `False` or `-1` update `RiffLastError`.

```vb
Dim bufferId As Long
bufferId = RiffLoad("C:\Audio\missing.wav")

If bufferId = -1 Then
    Debug.Print "Riff error:", RiffLastError
End If
```

### Create a synthesized alert tone

Oscillators do not need an audio file. They use the same voice controls and DSP chain as loaded buffers.

```vb
Public Sub PlayAlertTone()
    If Not RiffOpen() Then Exit Sub

    Dim voiceId As Long
    voiceId = RiffPlayOscillator(RiffWaveSine, 880)

    If voiceId >= 0 Then
        RiffVoiceVolume(voiceId) = 0.25
        RiffVoicePan(voiceId) = 0
        RiffVoiceReverbMix(voiceId) = 0.2
        RiffFadeOut voiceId, 0.8
    End If
End Sub
```

### Shape a radio-style voice effect

Riff properties are assigned per voice. Load and play a voice clip, then apply high-pass, low-pass, and distortion to narrow and roughen the sound.

```vb
Public Sub PlayRadioVoice()
    Dim bufferId As Long
    bufferId = RiffLoad("C:\Audio\dialog.wav")
    If bufferId < 0 Then Exit Sub

    Dim voiceId As Long
    voiceId = RiffPlay(bufferId)
    If voiceId < 0 Then Exit Sub

    RiffVoiceHighPass(voiceId) = 0.35
    RiffVoiceLowPass(voiceId) = 0.42
    RiffVoiceDistortion(voiceId) = 1.6
    RiffVoiceEqBass(voiceId) = 0.55
    RiffVoiceEqMid(voiceId) = 1.4
    RiffVoiceEqTreble(voiceId) = 0.75
End Sub
```

### Export audio to WAV

Loaded buffers and generated oscillators can be rendered to standard 16-bit stereo PCM WAV files.

```vb
Public Sub ExportExamples()
    Dim bufferId As Long
    bufferId = RiffLoad("C:\Audio\source.ogg")

    If bufferId >= 0 Then
        RiffExportBufferWav bufferId, "C:\Audio\source_export.wav"
    End If

    RiffRenderOscillatorWav RiffWaveSquare, 110, 2.5, "C:\Audio\square_110hz.wav"
End Sub
```

## Feature Summary

- 32 simultaneous polyphonic voices.
- 64 decoded audio buffers allocated with `VirtualAlloc`.
- File loading through `RiffLoad`.
- In-memory loading through `RiffLoadFromMemory`.
- Sine, square, sawtooth, and noise oscillators.
- BLEP correction for square and saw oscillators.
- 16 audio buses with independent volume.
- Per-voice DSP: reverb, chorus, flanger, delay, compressor, EQ, low-pass, high-pass, distortion, bitcrusher, sample-rate reduction, ring modulation, tremolo, auto-pan, stereo width, fade, pan, pitch, and looping.
- Master and per-voice peak meters.
- WAV export for loaded buffers and rendered oscillators.
- Office x86 and x64 support.

## Supported Hosts

Riff is designed for Windows VBA hosts, including:

- Excel
- Word
- PowerPoint
- Access
- Outlook
- Other VBA7-compatible hosts

Office 32-bit and 64-bit are both supported through conditional compilation.

## Documentation

- [API Reference](docs/API_REFERENCE.md): public functions, properties, parameter ranges, and operational guidance.
- [Architecture](docs/ARCHITECTURE.md): internal memory model, WASAPI path, Media Foundation decoding, timer thunk, DSP pipeline, and shutdown sequence.
- [Effect Cookbook](docs/EFFECT_COOKBOOK.md): copy-ready effect presets for polished application audio.
- [Troubleshooting](docs/TROUBLESHOOTING.md): practical diagnostics for initialization, loading, playback, and shutdown issues.
- [Examples](examples/README.md): detailed usage patterns and the showcase module.
- [Package Guide](package/README.md): import and integration guidance.

## Roadmap

Completed:

- WASAPI Shared Mode playback.
- Media Foundation decoding.
- In-memory audio loading.
- 32-voice polyphony.
- 64-buffer pool.
- 16 audio buses.
- Full per-voice DSP pipeline.
- Fade, loop, seek, pitch, and metering support.
- x86 and x64 timer thunks.
- WAV export and oscillator rendering.

Planned:

- Asynchronous loading.
- Dedicated ABI-safe direct vtable thunks for high-volume decode paths.
- macOS support through CoreAudio and AudioToolbox.

## License

MIT. See [LICENSE](LICENSE).
