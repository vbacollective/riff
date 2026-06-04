<div align="center">
  <img src="resources/logo.png" width="160" alt="Riff logo" />
</div>

<h1 align="center">Riff</h1>

<p align="center">
  <b>A high-performance, single-file audio engine for Microsoft Office.</b><br/>
  Real-time WASAPI playback, Media Foundation decoding, Studio DSP, adaptive buffering, and BLEP synthesis.
</p>

<p align="center">
  <img src="https://github.com/vbacollective/riff/actions/workflows/ci.yml/badge.svg" alt="CI" />
  <img src="https://github.com/vbacollective/riff/actions/workflows/release-assets.yml/badge.svg" alt="Release" />
  <img src="https://img.shields.io/badge/version-v1.0.8-blue.svg" alt="Version" />
  <img src="https://img.shields.io/badge/language-VBA-867DB1.svg" alt="Language" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D6.svg" alt="Platform" />
  <img src="https://img.shields.io/badge/arch-32%20%26%2064--bit-green.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/WASAPI-Shared%20Mode-blue.svg" alt="WASAPI" />
  <img src="https://img.shields.io/badge/Media%20Foundation-Decoding-orange.svg" alt="Media Foundation" />
  <img src="https://img.shields.io/badge/DSP-Studio%20Pipeline-critical.svg" alt="DSP" />
  <img src="https://img.shields.io/badge/dependencies-none-success.svg" alt="Dependencies" />
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" />
</p>

**Riff** is a complete, production-grade audio engine contained within a single `.bas` module. It allows VBA developers to integrate professional-quality audio playback, synthesis, routing, and DSP into Microsoft Office applications without external DLLs, ActiveX controls, installers, or additional references.

Whether you are building interactive dashboards in Excel, immersive presentations in PowerPoint, educational tools in Word, or automation systems in Access, Riff provides a practical real-time audio layer powered by the Windows Audio Session API (WASAPI), Media Foundation, and a native timer-driven DSP loop.

Riff is designed for developers who want game-like audio behavior inside Office: responsive UI sounds, background music, routed buses, soft limiting, fades, presets, procedural oscillators, noise generation, and safe cleanup when the host application closes.

## Key Capabilities

- **Zero Dependencies:** No DLLs to ship or register. Import `Riff.bas` and run.
- **Single-File Distribution:** The entire engine lives in one VBA module.
- **WASAPI Playback:** Shared-mode Windows audio output with low-latency rendering.
- **Media Foundation Decoding:** Loads common formats supported by the system, including WAV, MP3, AAC, FLAC, WMA, and more.
- **WAV Fast Path:** Compatible WAV files bypass Media Foundation and load through a direct RIFF parser for much faster startup.
- **Unified Playback API:** `RiffPlay` now accepts bus, loop, volume, and pan parameters directly.
- **Duplicate Prevention:** `RiffPlayOnce` prevents music or repeated ambience from stacking accidentally.
- **Adaptive Buffering:** Dynamically increases render queue safety during host stalls and returns to low latency when stable.
- **Burst Protection:** Voice stealing, per-buffer caps, and per-bus caps reduce stutter when many SFX are triggered rapidly.
- **Studio DSP Pipeline:** Independent per-voice effects including Reverb, Delay, Chorus, Flanger, Compressor, EQ, Filters, Distortion, Bitcrusher, Ring Modulation, Tremolo, Auto-Pan, and Stereo Width.
- **Effect Presets:** Ready-to-use presets such as Lo-Fi, Radio, Underwater, Echo, Small Room, Hall, Cathedral, Telephone, and more.
- **Real-Time Synthesis:** BLEP-corrected oscillators for sine, square, and saw waveforms.
- **Noise Generation:** White, pink, and brown noise for procedural ambience, wind, rain, static, rumble, and retro effects.
- **Audio Routing:** 16 global buses for Music, SFX, UI, Voice, and auxiliary groups.
- **Mixer Controls:** Bus volume, mute, solo, fades, peak meters, and master peak monitoring.
- **Smoothing:** Volume, pan, and pitch smoothing reduce clicks during parameter changes.
- **Soft Clipping:** Master soft clipper helps prevent harsh digital clipping when many voices overlap.
- **Diagnostics:** Render counters, underrun counters, clipping counters, buffer status, and adaptive queue information.
- **WAV Export:** Export loaded buffers and generated oscillators as standard PCM WAV files.
- **Architecture Aware:** Compatible with both 32-bit and 64-bit Office through `#If VBA7` / `#If Win64` declarations.

## Getting Started

### Installation

1. [Download](https://github.com/vbacollective/riff/releases) the latest `Riff.bas`.
2. Open the VBA Editor with `Alt + F11`.
3. Choose **File > Import File...** and select `Riff.bas`.
4. No external references are required.
5. Save your Office document as a macro-enabled file, such as `.xlsm`, `.pptm`, `.docm`, or `.accdb`.

### Minimal Implementation

Initialize the engine, load an asset, and play it.

```vb
Public Sub PlaySound()
    If Not RiffOpen() Then
        MsgBox "Riff failed to initialize. Error: " & RiffLastError
        Exit Sub
    End If

    Dim buf As Long
    buf = RiffLoad("C:\Audio\click.wav")

    If buf < 0 Then
        MsgBox "Failed to load audio. Error: " & RiffLastError
        Exit Sub
    End If

    Dim voice As Long
    voice = RiffPlay(buf, RiffBusSfx, False, 0.8, 0)

    If voice < 0 Then
        Debug.Print "No free voice. Error:", RiffLastError
    End If
End Sub
```

### Essential Cleanup

Always close the engine when your document, workbook, or presentation is closing.

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RiffClose
End Sub
```

For PowerPoint, call cleanup from your slideshow termination flow:

```vb
Public Sub OnSlideShowTerminate(ByVal Pres As Presentation)
    RiffClose
End Sub
```

> [!IMPORTANT]
> Riff uses native callbacks and WASAPI COM interfaces. Always call `RiffClose` before resetting the VBA project, closing the document, or ending a slideshow.

## Basic Usage

### Load Assets Once

Audio loading is synchronous. Load your assets during startup, not during gameplay or button-spam interactions.

```vb
Private sndClick As Long
Private sndExplosion As Long
Private sndMusic As Long

Public Sub AudioLoad()
    If Not RiffOpen() Then Exit Sub

    sndClick = RiffLoad(ActivePresentation.Path & "\audio\click.wav")
    sndExplosion = RiffLoad(ActivePresentation.Path & "\audio\explosion.wav")
    sndMusic = RiffLoad(ActivePresentation.Path & "\audio\music.wav")

    RiffBusVolume(RiffBusUi) = 0.9
    RiffBusVolume(RiffBusSfx) = 0.85
    RiffBusVolume(RiffBusMusic) = 0.45
End Sub
```

### Play UI and SFX

```vb
Public Sub PlayClick()
    RiffPlay sndClick, RiffBusUi, False, 0.7, 0
End Sub

Public Sub PlayExplosion()
    Dim pan As Single
    pan = (Rnd() * 2!) - 1!

    RiffPlay sndExplosion, RiffBusSfx, False, 1!, pan
End Sub
```

### Play Music Without Duplicating It

```vb
Private musicVoice As Long

Public Sub PlayMusic()
    musicVoice = RiffPlayOnce(sndMusic, RiffBusMusic, True, 0.5, 0)
End Sub

Public Sub StopMusic()
    If RiffVoiceActive(musicVoice) Then
        RiffVoiceFadeTo musicVoice, 0, 500
    End If
End Sub
```

### Generate a Procedural Sound

```vb
Public Sub PlayBeep()
    Dim v As Long
    v = RiffPlayOscillator(RiffWaveSine, 880, RiffBusUi, 0.25, 0)

    If v >= 0 Then
        RiffFadeOut v, 0.15
    End If
End Sub
```

### Generate Noise

```vb
Public Sub PlayRainLayer()
    Dim v As Long
    v = RiffPlayNoise(RiffWavePinkNoise, RiffBusSfx, 0.08, 0)

    If v >= 0 Then
        RiffVoiceLoop(v) = True
        RiffVoiceLowPass(v) = 0.55
        RiffVoiceReverbMix(v) = 0.2
    End If
End Sub
```

## Professional Audio Routing

Riff supports logical summing through 16 independent audio buses. Buses allow you to group related sounds, such as Music, SFX, UI, Voice, or Ambience, and control them as a single unit.

This is more efficient and cleaner than iterating through active voices manually. It also enables mixer-like behavior such as lowering music during dialogue, muting UI sounds, soloing a bus for debugging, or fading whole categories.

### Signal Hierarchy

The final volume of a sound is determined by:

```text
Master Volume × Bus Volume × Voice Volume × Fade/Smoothing
```

```vb
Public Sub SetupGameMixer()
    RiffMasterVolume = 1!

    RiffBusVolume(RiffBusMusic) = 0.45
    RiffBusVolume(RiffBusSfx) = 0.85
    RiffBusVolume(RiffBusUi) = 0.9
    RiffBusVolume(RiffBusVoice) = 1!
End Sub
```

### Bus Fade

```vb
Public Sub EnterPauseMenu()
    RiffBusFadeTo RiffBusMusic, 0.2, 500
    RiffBusFadeTo RiffBusSfx, 0.5, 300
End Sub

Public Sub LeavePauseMenu()
    RiffBusFadeTo RiffBusMusic, 0.45, 500
    RiffBusFadeTo RiffBusSfx, 0.85, 300
End Sub
```

### Mute and Solo

```vb
Public Sub ToggleMusicMute(ByVal muted As Boolean)
    RiffBusMuted(RiffBusMusic) = muted
End Sub

Public Sub DebugSoloSfx()
    RiffBusSolo(RiffBusSfx) = True
End Sub

Public Sub ClearSolo()
    RiffBusSolo(RiffBusSfx) = False
End Sub
```

### Peak Meter

```vb
Public Sub PrintBusPeaks()
    Dim l As Single
    Dim r As Single

    RiffBusGetPeak RiffBusMusic, l, r
    Debug.Print "Music peak:", l, r

    RiffMasterGetPeak l, r
    Debug.Print "Master peak:", l, r
End Sub
```

## Unified Playback API

The recommended public API is intentionally compact:

```vb
voice = RiffPlay(bufferHandle, busID, looped, volume, pan)
voice = RiffPlayOnce(bufferHandle, busID, looped, volume, pan)
voice = RiffPlayOscillator(waveType, frequencyHz, busID, volume, pan)
voice = RiffPlayNoise(noiseType, busID, volume, pan)
```

### `RiffPlay`

```vb
Dim v As Long
v = RiffPlay(sndExplosion, RiffBusSfx, False, 0.9, -0.2)
```

### `RiffPlayOnce`

Use this for music, ambience, menu loops, and anything that should not duplicate.

```vb
musicVoice = RiffPlayOnce(sndMusic, RiffBusMusic, True, 0.5, 0)
```

### `RiffPlayOscillator`

```vb
Dim osc As Long
osc = RiffPlayOscillator(RiffWaveSquare, 220, RiffBusSfx, 0.2, 0)
RiffFadeOut osc, 0.4
```

### `RiffPlayNoise`

```vb
Dim wind As Long
wind = RiffPlayNoise(RiffWaveBrownNoise, RiffBusSfx, 0.12, 0)
RiffVoiceLowPass(wind) = 0.35
```

### Compatibility Wrappers

Older function names are kept for compatibility:

```vb
RiffPlayBus bufferHandle, busID
RiffPlayBusOnce bufferHandle, busID, looped
RiffPlayOscillatorBus waveType, frequencyHz, busID
RiffPlayNoiseBus noiseType, busID
```

They now forward internally to the unified playback path.

## Effect Presets

Presets provide fast, musical starting points for common sound design situations.

```vb
RiffVoiceApplyPreset voice, RiffFxLoFi, 0.55
RiffVoiceApplyPreset voice, RiffFxRadio, 0.6
RiffVoiceApplyPreset voice, RiffFxUnderwater, 0.7
RiffVoiceApplyPreset voice, RiffFxSmallRoom, 0.4
```

### Typical Presets

| Preset | Use Case |
|:---|:---|
| `RiffFxDry` | Clears effect-heavy coloration and returns toward a clean signal. |
| `RiffFxSmallRoom` | Subtle room reflections for close spaces. |
| `RiffFxHall` | Wider ambience for music or narration. |
| `RiffFxCathedral` | Large washed-out reverb tail. |
| `RiffFxEcho` | Musical delay/echo effect. |
| `RiffFxLoFi` | Tape/sampler-style degradation without destroying the sound. |
| `RiffFxRadio` | Band-limited radio or speaker effect. |
| `RiffFxTelephone` | Narrow voice effect. |
| `RiffFxUnderwater` | Muffled filtered sound with movement. |
| `RiffFxRobot` | Ring-modulated robotic coloration. |
| `RiffFxWide` | Enhanced stereo width. |
| `RiffFxMono` | Mono collapse for compatibility or retro effects. |

### Preset Amount

`amount` usually ranges from `0.0` to `1.0`.

```vb
RiffVoiceApplyPreset v, RiffFxLoFi, 0.3  ' light coloration
RiffVoiceApplyPreset v, RiffFxLoFi, 0.8  ' stronger effect
RiffVoiceApplyPreset v, RiffFxDry, 1.0   ' clear preset-style coloration
```

## Manual DSP Helpers

You can use the individual properties directly, or use helper functions to set common groups of parameters.

```vb
RiffVoiceSetReverb v, 0.35, 0.7
RiffVoiceSetDelay v, 0.28, 0.45, 0.5
RiffVoiceSetChorus v, 0.5, 1.2
RiffVoiceSetFlanger v, 0.6, 0.35, 0.4
RiffVoiceSetFilter v, 0.45, 0.05
RiffVoiceClearEffects v
```

### Smooth Voice Changes

Use smoothing helpers to avoid clicks and sudden jumps.

```vb
RiffVoiceVolumeTo musicVoice, 0.2, 500
RiffVoicePanTo voice, -0.5, 150
RiffVoicePitchTo voice, 1.2, 100
```

## Noise and Oscillators

Riff supports both tonal oscillators and procedural noise.

| Waveform | Description | Common Uses |
|:---|:---|:---|
| `RiffWaveSine` | Pure tone | UI beeps, tests, soft synth tones |
| `RiffWaveSquare` | Hollow retro wave | Chiptune, alarms, retro UI |
| `RiffWaveSawtooth` | Bright harmonic wave | Synth leads, sweeps, engine-like sounds |
| `RiffWaveWhiteNoise` | Equal random energy | Static, glitch, impacts, noise bursts |
| `RiffWavePinkNoise` | Natural balanced noise | Rain, wind, fire, ambience |
| `RiffWaveBrownNoise` | Dark low-frequency noise | Rumble, thunder, earthquake, machinery |
| `RiffWaveNoise` | Compatibility alias | Equivalent to white noise |

```vb
RiffPlayOscillator RiffWaveSine, 440, RiffBusSfx, 0.2, 0
RiffPlayNoise RiffWavePinkNoise, RiffBusSfx, 0.08, 0
```

## WAV Fast Path

`RiffLoad` now attempts a direct WAV fast path before falling back to Media Foundation.

When possible, compatible WAV files are parsed directly by Riff:

```text
WAV file -> RIFF parser -> VirtualAlloc buffer -> Riff buffer pool
```

This avoids Media Foundation startup, source reader negotiation, COM loops, and format conversion overhead for simple WAV assets.

### Recommended Runtime Format

For fastest loading, use WAV files close to the active output mix:

```text
PCM16 stereo 48 kHz
PCM16 stereo 44.1 kHz
Float32 stereo 48 kHz
```

Small UI sounds and SFX should generally be shipped as WAV if load speed matters.

```vb
Dim click As Long
click = RiffLoad(ActivePresentation.Path & "\audio\click.wav")
```

If the WAV format is unsupported by the fast path, Riff automatically falls back to the Media Foundation decode path.

## Adaptive Buffering

Office hosts can occasionally stall when switching tabs, opening menus, recalculating, rendering slides, or interacting with the VBA editor. Because Riff is implemented in pure VBA with a native callback bridge, those stalls can affect how often the render loop gets serviced.

Riff v1.0.8 includes adaptive buffering to reduce audible dropouts:

```text
Normal state:
    target queue stays low for responsive playback

Stall or underrun risk:
    target queue rises automatically for safety

Stable state:
    target queue gradually returns to low latency
```

### Diagnostics

```vb
Debug.Print "Adaptive queue:", RiffAdaptiveQueueMs
Debug.Print "Underruns:", RiffUnderrunCount
Debug.Print "Last padding:", RiffLastPaddingFrames
Debug.Print "Frames available:", RiffLastFramesAvailable
Debug.Print "Frames written:", RiffLastFramesWritten
```

Reset counters before a test:

```vb
RiffResetAdaptiveStats
RiffResetDiagnostics
```

## Burst Safety

Triggering the same sound many times in a short time can overwhelm any small mixer. Riff includes safety controls to prevent rapid SFX spam from creating 32 overlapping copies of the same sound.

```vb
RiffVoiceStealingEnabled = True
RiffMaxVoicesPerBuffer = 6
RiffMaxVoicesPerBus = 20
```

### Diagnostics

```vb
Debug.Print "Active voices:", RiffActiveVoiceCount()
Debug.Print "Click instances:", RiffBufferVoiceCount(sndClick, RiffBusUi)
Debug.Print "SFX bus voices:", RiffBusVoiceCount(RiffBusSfx)
```

Recommended values for UI-heavy Office projects:

```vb
RiffMaxVoicesPerBuffer = 4
RiffMaxVoicesPerBus = 16
```

## Feature Summary

| Category | Features |
|:---|:---|
| **Core** | Single `.bas`, 32 voices, 64 buffers, 16 buses, x86/x64 support |
| **I/O** | Media Foundation decoding, in-memory loading, WAV fast path, WAV export |
| **Playback** | Unified `RiffPlay`, `RiffPlayOnce`, seeking, looping, fades, stop/reset behavior |
| **Routing** | Bus volume, mute, solo, fade, peak meters, master volume |
| **Stability** | Adaptive buffering, burst protection, voice stealing, cleanup safeguards |
| **Synthesis** | BLEP sine/square/saw, white noise, pink noise, brown noise |
| **Dynamics** | Compressor, soft clipping, distortion, bitcrusher |
| **Spatial** | Reverb, delay, stereo width, pan, auto-pan |
| **Modulation** | Chorus, flanger, tremolo, ring modulation |
| **Filters** | Biquad low-pass, high-pass, 3-band EQ |
| **Diagnostics** | Underruns, render errors, clipped samples, buffer state, active voices |
| **Export** | Loaded buffer export and oscillator render to PCM WAV |

## v1.0.8 Highlights

- Unified playback API with bus, loop, volume, and pan parameters.
- Added `RiffPlayOnce` for duplicate-safe music and ambience.
- Reworked bus system with volume, mute, solo, fades, and peak meters.
- Added polished FX presets and helper APIs.
- Added white, pink, and brown noise.
- Improved noise generation with deterministic per-voice RNG.
- Added smoothing for voice volume, pan, and pitch.
- Added master soft clipper to reduce harsh clipping.
- Added adaptive buffering for better stability when Office or Windows stalls.
- Added render diagnostics, underrun counters, clipping counters, and queue stats.
- Added WAV fast path for faster `RiffLoad` on compatible WAV assets.
- Optimized decode-path vtable calls.
- Fixed decode buffer growth failure handling.
- Improved stop behavior with WASAPI reset for more responsive stopping.
- Added timer suspend/wake and safer cleanup behavior.
- Added voice stealing and burst limits for rapid SFX triggering.
- Preserved compatibility wrappers for older API calls.

## Documentation

- [**API Reference**](docs/API_REFERENCE.md) – Detailed guide to every function, property, enum, and practical pattern.
- [**Architecture**](docs/ARCHITECTURE.md) – Deep dive into WASAPI, Media Foundation, thunks, callback flow, buffers, and DSP.
- [**Effect Cookbook**](docs/EFFECT_COOKBOOK.md) – Ready-to-use recipes for radio voices, underwater states, ambience, retro effects, and more.
- [**Troubleshooting**](docs/TROUBLESHOOTING.md) – Fixes for initialization, device, stutter, cleanup, and host-specific issues.
- [**Examples**](examples/README.md) – Practical demos and integration patterns.

## Roadmap

### Current Version (v1.0.8)

- [x] WASAPI Shared Mode playback.
- [x] Media Foundation decoding.
- [x] WAV fast path loader.
- [x] Unified playback API.
- [x] 32-voice polyphonic mixer.
- [x] 16-bus routing system.
- [x] Bus mute, solo, fade, and peak meters.
- [x] Full per-voice Studio DSP pipeline.
- [x] Effect presets.
- [x] White, pink, and brown noise.
- [x] BLEP oscillators.
- [x] Adaptive buffering.
- [x] Burst-safe voice management.
- [x] x86/x64 native thunk driver.
- [x] Offline WAV export.
- [x] Diagnostics and render counters.

### Planned

- [ ] Optional native decode backend for faster MP3/OGG loading.
- [ ] Optional streaming API for long music tracks.
- [ ] Background preload helpers.
- [ ] More musical preset packs.
- [ ] Additional master bus processors.
- [ ] Higher-level asset registry, such as `RiffLoadAs` and `RiffPlayKey`.
- [ ] macOS support through CoreAudio/AudioToolbox if the project expands beyond Windows.

## Performance Notes

For best performance in Office:

- Prefer WAV for short SFX and UI sounds.
- Preload assets at startup with `RiffLoad`.
- Avoid decoding MP3/OGG during interaction-heavy moments.
- Use `RiffPlayOnce` for music and ambience.
- Use bus volume/fades instead of changing many voices one by one.
- Keep effect-heavy processing for important voices only.
- Use `RiffMaxVoicesPerBuffer` to prevent repeated button clicks from stacking too many copies.
- Always call `RiffClose` on exit.

> [!NOTE]
> Riff is a pure VBA engine with a native callback bridge. It is highly capable for Office, but it is still hosted inside Excel, PowerPoint, Word, or Access. If the host application or the entire system stalls hard enough, audio scheduling can be affected. Adaptive buffering reduces this, but a fully independent audio thread would require a native backend.

## License

MIT. Designed for freedom, integration, and serious audio experimentation inside Microsoft Office.
