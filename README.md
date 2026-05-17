<div align="center">
  <img src="resources/logo.png" width="150" />
</div>

<h1 align="center">Riff - VBA Audio Engine</h1>

<p align="center">
  <b>A complete WASAPI audio engine for Microsoft Office. No DLLs, no dependencies, no installation.</b>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/language-VBA-867DB1.svg" alt="Language" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D6.svg" alt="Platform" />
  <img src="https://img.shields.io/badge/arch-32%20%26%2064--bit-green.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/WASAPI-Shared%20Mode-blue.svg" alt="WASAPI" />
  <img src="https://img.shields.io/badge/Media%20Foundation-Decoding-orange.svg" alt="Media Foundation" />
  <img src="https://img.shields.io/badge/Polyphony-32%20Voices-blueviolet.svg" alt="Polyphony" />
  <img src="https://img.shields.io/badge/DSP-Studio%20Pipeline-critical.svg" alt="DSP" />
  <img src="https://img.shields.io/badge/Assembly-x86%20%26%20x64%20Thunks-red.svg" alt="Assembly" />
  <img src="https://img.shields.io/badge/dependencies-none-success.svg" alt="Dependencies" />
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" />
</p>

> [!NOTE]
> **Supported Applications**
>
> Excel, Word, PowerPoint, Access, Outlook, and any Microsoft Office VBA host.

> [!IMPORTANT]
> **Platform**
>
> ![](resources/svg/windows.svg) Currently Windows only. Riff relies on `kernel32`, `user32`, `ole32`, `mfplat`, `mfreadwrite`, `winmm` and `shlwapi`. Mac support is planned.

## Table of Contents

- [What is Riff](#what-is-riff)
- [Quick Start](#quick-start)
- [Architecture](#architecture)
  - [Why a Standard Module](#why-a-standard-module)
  - [The Assembly Thunk](#the-assembly-thunk)
  - [COM VTable Dispatch](#com-vtable-dispatch)
  - [The DSP Pipeline](#the-dsp-pipeline)
- [Features](#features)
- [API Reference](#api-reference)
- [Examples](#examples)
- [Compatibility](#compatibility)
- [Roadmap](#roadmap)
- [License](#license)

## What is Riff

Riff is a single `.bas` module that brings real audio to any Microsoft Office VBA host. Load audio files from disk or directly from memory, synthesize waveforms, and apply a full studio DSP pipeline per voice, all in real time, with zero external dependencies and no configuration required.

It talks directly to WASAPI through hand-rolled COM vtable calls, decodes any format Windows supports through Media Foundation, and drives the audio callback through a runtime-compiled x86/x64 assembly thunk that fires independently of the VBA execution thread.

A single file drop into any VBA project is all it takes. No references need to be enabled in **Tools > References**.

## Quick Start

Import `Riff.bas` into your VBA project via **File > Import File** in the VBA editor.

### Play a sound

```vb
RiffOpen

Dim buf As Long
buf = RiffLoad("C:\sounds\explosion.wav")
RiffPlay buf
```

### Play with effects

```vb
Dim buf As Long
Dim voice As Long

buf = RiffLoad("C:\sounds\music.ogg")
voice = RiffPlay(buf)

RiffVoiceLoop(voice) = True
RiffVoiceVolume(voice) = 0.8
RiffVoiceReverbMix(voice) = 0.4
RiffVoiceReverbTime(voice) = 0.75
```

### Synthesize a waveform

```vb
Dim voice As Long
voice = RiffPlayOscillator(0, 440) ' Sine wave at 440 Hz

RiffVoiceFlangerDepth(voice) = 0.6
RiffVoiceFlangerRate(voice) = 0.3
```

### Load audio from memory

```vb
Dim data() As Byte
data = SomeByteArrayContainingAnAudioFile()

Dim buf As Long
buf = RiffLoadFromMemory(data)
RiffPlay buf
```

## Architecture

### Why a Standard Module

This is a deliberate design choice. Several concrete constraints of the VBA environment make a procedural standard module the right foundation for an audio engine at this level.

**Zero COM overhead.** Every VBA class is a COM object. Method dispatch, reference counting, and heap allocation all accumulate at audio thread frequency. A standard `.bas` module bypasses the COM layer entirely and keeps the hot path clean.

**Data-oriented design.** Riff manages 32 polyphonic voices and 64 static buffers as statically allocated User-Defined Types in contiguous memory. This is more cache-friendly and eliminates heap fragmentation across long Office sessions.

**Native Win32 alignment.** Audio at this level requires direct memory access via `RtlMoveMemory`, raw pointer arithmetic, and COM vtable calls through `DispCallFunc`. A procedural module provides the flat memory model that makes this possible without intermediate copies.

**No lifecycle to manage.** State is accessed globally through simple integer handles. There are no object scopes to worry about and no risk of a variable falling out of scope while audio is playing.

### The Assembly Thunk

The core of the Riff engine is a machine code thunk compiled at runtime and injected into executable memory via `VirtualAlloc`. This thunk is what makes the audio loop possible inside VBA.

**How it works:**

1. On `RiffOpen`, `VirtualAlloc` allocates a block of memory with `PAGE_EXECUTE_READWRITE` permissions.
2. Raw x86 or x64 opcodes are written into that block via `RtlMoveMemory`, with function pointers patched inline.
3. `SetTimer` registers this thunk as a Win32 timer callback firing every 15ms.
4. On each tick, the thunk checks `EbMode` from `vbe7.dll` to detect if the VBA runtime is still alive. If not, it calls `KillTimer` and exits safely without crashing the host application.
5. If the runtime is alive, it calls `RiffTimerCallback`, which runs the DSP pipeline and pushes audio to WASAPI.

The x64 and x86 thunks use different calling conventions. x64 follows the Microsoft x64 ABI with arguments in `RCX`, `RDX`, `R8`, `R9`. x86 uses `stdcall` with stack cleanup via `ret 10h`. Both are selected transparently through `#If Win64`.

This is the same pattern used in professional VBA subclassing libraries, adapted here to drive a real-time audio engine.

### COM VTable Dispatch

Riff does not use any high-level wrappers for WASAPI or Media Foundation. It calls COM interfaces directly through their virtual function tables using `DispCallFunc` from `oleaut32.dll`.

The `vCall` function at the core of this approach accepts a COM interface pointer, a zero-based vtable index, and a `ParamArray` of arguments, and dispatches directly to the underlying C++ method:

```vb
' Equivalent to calling IAudioClient::Initialize
vCall rCtx.AudioClient, 3, AUDCLNT_SHAREMODE_SHARED, 0&, hnsDur, hnsPer, rCtx.MixFormatPtr, pNullPtr
```

This eliminates all dependency on typelib-bound interface wrappers and makes the entire WASAPI stack self-contained inside the module. The same technique is used for Media Foundation's `IMFSourceReader` during audio decoding.

### The DSP Pipeline

Each of the 32 polyphonic voices runs a full per-voice DSP chain on every audio frame. The pipeline processes samples in this order:

**Source** (buffer PCM or oscillator) > **Bitcrusher** > **Sample Rate Reduction** > **Distortion** > **Low Pass Filter** > **High Pass Filter** > **3-Band EQ** > **Ring Modulator** > **Tremolo** > **Stereo Width** > **Flanger** > **Chorus** > **Delay** > **Reverb** > **Compressor** > **AutoPan** > **Volume / Pan / Bus** > **Fade** > **Master Mix**

All modulated effects (Chorus, Flanger, Tremolo, AutoPan, RingMod) use LFO phase accumulators that persist across frames, producing continuous and smooth modulation without clicks or resets.

The ring buffer backing Chorus, Flanger, Delay, and Reverb is a single contiguous 1D array of `32 * 192000` floats. This avoids the Column-Major 2D array wipe issue inherent to VBA's memory layout when using `RtlZeroMemory` on multi-dimensional arrays.

Riff supports both 32-bit float and 16-bit integer WASAPI output formats and detects the active format automatically from the device's mix format at initialization.

## Features

- **32 polyphonic voices** playing simultaneously with independent DSP state
- **64 static audio buffers** decoded into physical memory via `VirtualAlloc`
- **In-memory loading** via `RiffLoadFromMemory`, enabling audio embedded directly in the VBA project
- **Built-in oscillators**: Sine, Square, Sawtooth, Triangle, Noise
- **8 audio buses** with independent volume control for grouping voices (music, SFX, voice, etc.)
- **Per-voice DSP pipeline** with Reverb, Chorus, Flanger, Delay, Compressor, 3-Band EQ, Low Pass, High Pass, Distortion, Bitcrusher, Sample Rate Reduction, Ring Modulator, Tremolo, AutoPan, Stereo Width
- **Fade in / Fade out** with frame-accurate interpolation
- **Loop regions** with sub-second precision via `RiffSetLoopRegionSec`
- **Pitch shifting** via playback rate control
- **VU meters** at voice and master level via peak amplitude tracking
- **32-bit float and 16-bit integer** WASAPI output, auto-detected
- **x86 and x64** support via `#If VBA7` and `#If Win64` conditional compilation
- **IDE-safe timer thunk** with `EbMode` liveness check to prevent crashes on VBE reset

## API Reference

### Initialization

| Function | Description |
|---|---|
| `RiffOpen()` | Initializes WASAPI, Media Foundation, and the timer thunk. Returns True on success. |
| `RiffClose()` | Shuts down the engine, stops all voices, and frees all memory. |
| `RiffIsInitialized` | Returns True if the engine is running. |

### Buffers

| Function | Description |
|---|---|
| `RiffLoad(filePath)` | Decodes an audio file from disk into memory. Returns a buffer handle (0-63) or -1. |
| `RiffLoadFromMemory(audioData())` | Decodes audio from a Byte array. Returns a buffer handle or -1. |
| `RiffUnload(bufferHandle)` | Frees a buffer and stops any voices using it. |
| `RiffBufferDurationSec(bufferHandle)` | Returns the duration of a loaded buffer in seconds. |

### Playback

| Function | Description |
|---|---|
| `RiffPlay(bufferHandle)` | Plays a buffer on a free voice. Returns a voice handle (0-31) or -1. |
| `RiffPlayOscillator(waveType, frequencyHz)` | Starts a synthetic oscillator on a free voice. |
| `RiffPause(voiceHandle)` | Pauses a voice without releasing it. |
| `RiffResume(voiceHandle)` | Resumes a paused voice. |
| `RiffStop(voiceHandle)` | Stops and frees a voice immediately. |
| `RiffStopAll()` | Stops all active voices. |
| `RiffFadeIn(voiceHandle, durationSec)` | Smoothly fades a voice in over the given duration. |
| `RiffFadeOut(voiceHandle, durationSec)` | Smoothly fades a voice out and stops it. |
| `RiffSetLoopRegionSec(voiceHandle, startSec, endSec)` | Constrains playback to loop within a region. |

### Voice Properties

| Property | Description |
|---|---|
| `RiffVoiceIsPlaying(v)` | True if the voice is actively playing. |
| `RiffVoiceIsPaused(v)` | True if the voice is paused. |
| `RiffVoiceVolume(v)` | Individual volume, 0.0 to 1.0. |
| `RiffVoicePitch(v)` | Playback speed and pitch. 1.0 is normal, 2.0 is one octave up. |
| `RiffVoicePan(v)` | Stereo pan. -1.0 left, 0.0 center, 1.0 right. |
| `RiffVoiceLoop(v)` | Enables or disables looping. |
| `RiffVoicePositionSec(v)` | Current playback position in seconds. Readable and writable. |
| `RiffVoiceBus(v)` | Routes the voice to a bus (0 to 7). |
| `RiffVoiceGetPeak(v, L, R)` | Returns the current peak amplitude for VU metering. |

### Master and Buses

| Property | Description |
|---|---|
| `RiffMasterVolume` | Global output volume, 0.0 to 1.0. |
| `RiffBusVolume(busID)` | Volume for a specific bus, 0.0 to 2.0. |
| `RiffMasterGetPeak(L, R)` | Returns the master output peak for VU metering. |

### DSP Effects

| Property | Description |
|---|---|
| `RiffVoiceLowPass(v)` | Low pass filter. Lower values muffle high frequencies. |
| `RiffVoiceHighPass(v)` | High pass filter. Removes low frequencies. |
| `RiffVoiceEqBass(v)` | Low shelf gain. 1.0 is flat. |
| `RiffVoiceEqMid(v)` | Mid band gain. 1.0 is flat. |
| `RiffVoiceEqTreble(v)` | High shelf gain. 1.0 is flat. |
| `RiffVoiceDistortion(v)` | Digital clipping. Values above 1.0 add saturation. |
| `RiffVoiceBitDepth(v)` | Bit depth reduction. 4 sounds like Game Boy, 8 like early PC. |
| `RiffVoiceSampleRateReduction(v)` | Downsampling factor. Creates robotic, lo-fi artifacts. |
| `RiffVoiceStereoWidth(v)` | Stereo field width. Below 1.0 narrows, above 1.0 exaggerates. |
| `RiffVoiceReverbMix(v)` | Reverb wet amount. |
| `RiffVoiceReverbTime(v)` | Reverb decay time. Higher values simulate larger spaces. |
| `RiffVoiceDelayTime(v)` | Echo interval in seconds. |
| `RiffVoiceDelayFeedback(v)` | Echo decay factor. |
| `RiffVoiceDelayMix(v)` | Echo wet amount. |
| `RiffVoiceChorusDepth(v)` | Chorus intensity. |
| `RiffVoiceChorusRate(v)` | Chorus LFO rate in Hz. |
| `RiffVoiceFlangerDepth(v)` | Flanger intensity. |
| `RiffVoiceFlangerRate(v)` | Flanger sweep rate in Hz. |
| `RiffVoiceFlangerFeedback(v)` | Flanger resonance. |
| `RiffVoiceTremoloRate(v)` | Tremolo LFO rate in Hz. |
| `RiffVoiceTremoloDepth(v)` | Tremolo intensity. |
| `RiffVoiceAutoPanRate(v)` | AutoPan LFO rate in Hz. |
| `RiffVoiceAutoPanDepth(v)` | AutoPan intensity. |
| `RiffVoiceRingModFreq(v)` | Ring modulator oscillator frequency in Hz. |
| `RiffVoiceRingModMix(v)` | Ring modulator wet amount. |
| `RiffVoiceCompressorThreshold(v)` | Level at which compression begins. |
| `RiffVoiceCompressorRatio(v)` | Compression strength above the threshold. |

## Examples

### Radio voice effect

```vb
Dim voice As Long
voice = RiffPlay(buf)
RiffVoiceHighPass(voice) = 0.3
RiffVoiceLowPass(voice) = 0.4
RiffVoiceDistortion(voice) = 1.8
```

### Underwater effect

```vb
Dim voice As Long
voice = RiffPlay(buf)
RiffVoiceLowPass(voice) = 0.08
RiffVoiceReverbMix(voice) = 0.6
RiffVoiceReverbTime(voice) = 0.8
RiffVoiceChorusDepth(voice) = 0.3
RiffVoiceChorusRate(voice) = 0.4
```

### Game Boy oscillator

```vb
Dim voice As Long
voice = RiffPlayOscillator(1, 440)
RiffVoiceBitDepth(voice) = 4
RiffVoiceSampleRateReduction(voice) = 8
```

### Game audio with buses

```vb
RiffOpen

Dim bufMusic As Long, bufSFX As Long
bufMusic = RiffLoad("music.ogg")
bufSFX = RiffLoad("explosion.wav")

Dim vMusic As Long, vSFX As Long
vMusic = RiffPlay(bufMusic)
vSFX = RiffPlay(bufSFX)

RiffVoiceBus(vMusic) = 0
RiffVoiceBus(vSFX) = 1
RiffVoiceLoop(vMusic) = True

RiffBusVolume(0) = 0.5
RiffBusVolume(1) = 1.0
```

### Loop with intro that does not repeat

```vb
Dim voice As Long
voice = RiffPlay(buf)
RiffVoiceLoop(voice) = True
RiffSetLoopRegionSec voice, 4.2, 38.7
```

### VU meter on a UserForm

```vb
Private Sub Timer_Tick()
    Dim pL As Single, pR As Single
    RiffMasterGetPeak pL, pR

    Dim pct As Long
    pct = CLng(((pL + pR) / 2) * 100)

    VUBar.Width = pct * 2

    If pct > 80 Then
        VUBar.BackColor = RGB(255, 50, 50)
    ElseIf pct > 50 Then
        VUBar.BackColor = RGB(255, 200, 0)
    Else
        VUBar.BackColor = RGB(50, 200, 50)
    End If
End Sub
```

## Compatibility

Riff uses only native Windows DLLs present on every version of Windows since Vista: `kernel32.dll`, `user32.dll`, `ole32.dll`, `oleaut32.dll`, `winmm.dll`, `mfplat.dll`, `mfreadwrite.dll`, and `shlwapi.dll`. No third-party installers, no COM registration, no ActiveX controls. Dropping the `.bas` file into a VBA project is all it takes.

### Operating System

| Version | Support |
|---|---|
| ![](resources/svg/windows.svg) Windows Vista | Supported |
| ![](resources/svg/windows.svg) Windows 7 | Supported |
| ![](resources/svg/windows.svg) Windows 8 / 8.1 | Supported |
| ![](resources/svg/windows.svg) Windows 10 | Supported |
| ![](resources/svg/windows.svg) Windows 11 | Supported |

### Office and VBA Host

| Environment | Support |
|---|---|
| Excel 32-bit | Supported |
| Excel 64-bit | Supported |
| Word 32-bit | Supported |
| Word 64-bit | Supported |
| PowerPoint 32-bit | Supported |
| PowerPoint 64-bit | Supported |
| Access 32-bit | Supported |
| Access 64-bit | Supported |
| Any VBA7 host (Office 2010+) | Supported |
| VBA6 (Office 2007 and earlier) | Supported |

32-bit and 64-bit compatibility is handled transparently through `#If VBA7` and `#If Win64` conditional compilation across all API declarations and assembly thunks.

## Roadmap

### ![](resources/svg/completed.svg) Completed

- [x] WASAPI Shared Mode output with automatic format detection (32-bit float and 16-bit integer)
- [x] Media Foundation decoding via `IMFSourceReader` with COM vtable dispatch
- [x] In-memory audio loading via `SHCreateMemStream` and `MFCreateSourceReaderFromByteStream`
- [x] 32 polyphonic voices with independent DSP state
- [x] 64 static audio buffers in physical memory via `VirtualAlloc`
- [x] 8 audio buses with independent volume control
- [x] Built-in oscillators: Sine, Square, Sawtooth, Triangle, Noise
- [x] Full per-voice DSP pipeline: Reverb, Chorus, Flanger, Delay, Compressor, EQ, Low Pass, High Pass, Distortion, Bitcrusher, Sample Rate Reduction, Ring Modulator, Tremolo, AutoPan, Stereo Width
- [x] Frame-accurate fade in / fade out
- [x] Loop regions with sample-aligned precision
- [x] Pitch shifting via playback rate
- [x] VU peak metering at voice and master level
- [x] IDE-safe timer thunk with `EbMode` liveness guard
- [x] x86 and x64 assembly thunks with correct calling conventions
- [x] `#If VBA7` and `#If Win64` full conditional compilation

### ![](resources/svg/planning.svg) Planned

- [ ] `RiffLoadAsync` with Win32 thread and callback on completion
- [ ] Direct vtable calls inside the decoding loop to eliminate `DispCallFunc` overhead
- [ ] Biquad filters for higher quality EQ and Low/High Pass
- [ ] Band-limited oscillators (BLEP) to eliminate aliasing at high frequencies
- [ ] Freeverb-style reverb for significantly improved spatial quality
- [ ] macOS support via CoreAudio and AudioToolbox

## License

MIT, free for personal and commercial use. See [LICENSE](LICENSE).
