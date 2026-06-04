<div align="center">
  <img src="resources/logo.png" width="160" alt="Riff logo" />
</div>

<h1 align="center">Riff</h1>

<p align="center">
  <b>A high-performance, single-file audio engine for Microsoft Office.</b><br/>
  Real-time WASAPI playback, Media Foundation decoding, Studio DSP, and BLEP synthesis.
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

**Riff** is a complete, production-grade audio engine contained within a single `.bas` module. It allows VBA developers to integrate professional-quality audio playback and synthesis into Microsoft Office applications without external DLLs, ActiveX controls, or complex installers.

Whether you are building interactive dashboards in Excel, immersive presentations in PowerPoint, or standalone tools in Access, Riff provides a sample-accurate DSP pipeline and low-latency playback powered by the Windows Audio Session API (WASAPI).

## Key Capabilities

- **Zero Dependencies:** No DLLs to ship or register. Just import `Riff.bas`.
- **High Performance:** Multi-threaded architecture using a native machine-code thunk for a 10ms DSP callback.
- **Pro Audio Stack:** WASAPI Shared Mode for low-latency output and Media Foundation for broad format support (WAV, MP3, AAC, FLAC).
- **Studio DSP Pipeline:** 32 polyphonic voices, each with independent Reverb, Chorus, Flanger, Delay, Compressor, 3-Band EQ, Filters, Distortion, and more.
- **Real-time Synthesis:** BLEP-corrected oscillators (Sine, Square, Saw, Noise) for UI beeps or retro sound design.
- **Audio Routing:** 16 global audio buses for logical volume grouping (Music, SFX, UI).
- **Advanced Control:** Precise looping, sample-accurate fades, pitch shifting, and peak metering.
- **Architecture Aware:** Fully compatible with both 32-bit and 64-bit Office (x86/x64).

## Getting Started

### Installation

1. [Download](https://github.com/vbacollective/riff/releases) the latest `Riff.bas`.
2. Open the VBA Editor (`Alt + F11`) in your Office host.
3. **File > Import File...** and select `Riff.bas`.
4. No external references are required. Save your file as a macro-enabled format (`.xlsm`, `.pptm`, etc.).

### Minimal Implementation

Initialize the engine, load an asset, and play it. Remember to shut down the engine when your document closes.

```vb
Public Sub PlaySound()
    ' Initialize hardware
    If Not RiffOpen() Then Exit Sub

    ' Load and play
    Dim buf As Long:  buf = RiffLoad("C:\Audio\click.wav")
    Dim voice As Long: voice = RiffPlay(buf)

    If voice >= 0 Then
        RiffVoiceVolume(voice) = 0.8
    End If
End Sub

' Essential cleanup to prevent host crashes on project reset
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RiffClose
End Sub
```

## Professional Audio Routing

Riff supports logical summing through 16 independent audio buses. Buses allow you to group related sounds—such as "Music," "Ambient Effects," or "UI Feedback"—and control them as a single unit. This is more efficient than iterating through active voices and allows for complex mixing scenarios like ducking music during dialogue or muting specific categories of sound via a settings menu.

### Signal Hierarchy
The final volume of any sound is determined by a multiplicative hierarchy:
**Master Volume** × **Bus Volume** × **Voice Volume**

```vb
Public Sub SetupGameMixer()
    ' Route voices to functional groups
    RiffVoiceBus(bgmVoice) = RiffBusMusic
    RiffVoiceBus(sfxVoice) = RiffBusSfx

    ' Control entire categories independently
    ' Bus volume ranges from 0.0 (muted) to 2.0 (boosted)
    RiffBusVolume(RiffBusMusic) = 0.4  ' Lower music for a focus state
    RiffBusVolume(RiffBusSfx) = 1.0    ' Keep sound effects at full scale
End Sub
```

## Feature Summary

| Category | Features |
|:---|:---|
| **Core** | 32 Voices, 64 Buffers, 16 Buses, Master Peak Meters |
| **I/O** | Media Foundation Decoding, In-Memory Loading, WAV Export |
| **Synthesis** | Sine, Square, Saw, Noise (BLEP Band-limited) |
| **Dynamics** | Compressor (Auto-envelope), Distortion, Bitcrusher |
| **Spatial** | Freeverb-style Reverb, Delay/Echo, Stereo Width, Panning |
| **Modulation** | Chorus, Flanger, Tremolo, Auto-Pan, Ring Modulator |
| **Filters** | Biquad Low-Pass, High-Pass, 3-Band Parametric EQ |
| **Playback** | Sample-accurate Fades, Pitch/Speed, Looping, Seeking |

## Documentation

- [**API Reference**](docs/API_REFERENCE.md) – Detailed guide to every function and property.
- [**Architecture**](docs/ARCHITECTURE.md) – Deep dive into thunks, WASAPI, and the DSP loop.
- [**Effect Cookbook**](docs/EFFECT_COOKBOOK.md) – Ready-to-use recipes for radio voices, underwater states, and more.
- [**Troubleshooting**](docs/TROUBLESHOOTING.md) – Fixes for common initialization or runtime issues.
- [**Examples**](examples/README.md) – Practical patterns and interactive demos.

## Roadmap

### Current Version (v1.0.7)
- [x] WASAPI & Media Foundation Integration.
- [x] Complete Studio DSP Pipeline.
- [x] x86/x64 Native Thunk Driver.
- [x] Offline WAV Export.

### Planned
- [ ] Asynchronous background loading.
- [ ] Direct VTable optimization for decode paths.
- [ ] macOS support (CoreAudio/AudioToolbox).

## License

MIT. Designed for freedom and integration.
