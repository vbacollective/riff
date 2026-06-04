# Riff API Reference

This document describes the complete public interface of **Riff.bas**, a high-performance, COM-based WASAPI audio engine for VBA. Riff provides real-time audio playback, Media Foundation decoding, a sample-accurate Studio DSP Pipeline, BLEP-corrected synthesis, and flexible audio routing.

## Table of Contents

- [Core Concepts](#core-concepts)
- [Initialization and Lifecycle](#initialization-and-lifecycle)
- [Diagnostics and Engine State](#diagnostics-and-engine-state)
- [Buffer Management](#buffer-management)
- [Playback and Voice Control](#playback-and-voice-control)
- [Audio Buses and Mixing](#audio-buses-and-mixing)
- [Voice Properties](#voice-properties)
- [DSP Pipeline and Effects](#dsp-pipeline-and-effects)
- [High-Level Effect Presets](#high-level-effect-presets)
- [Offline Rendering and Export](#offline-rendering-and-export)
- [Enumerations](#enumerations)

## Core Concepts

### Handles and IDs
- **Buffer Handles (`0` to `63`):** Represent decoded PCM data stored in physical memory. A buffer must be loaded before it can be played.
- **Voice Handles (`0` to `31`):** Represent active playback channels. A voice handle is returned by playback functions and remains valid until the voice stops or is stopped manually.
- **Bus IDs (`0` to `15`):** Represent logical mixing groups. Every voice is routed to a bus (default: `RiffBusMain`). Bus volume is applied after voice volume and before master volume.

### Memory Safety
Riff operates at a low level using `VirtualAlloc` and native thunks. 

> [!WARNING]
> Always call `RiffClose` before host shutdown or project reset. Failing to do so can leave native timers active, leading to crashes or "Out of memory" errors when the project is re-run.

## Initialization and Lifecycle

### RiffOpen
`Public Function RiffOpen() As Boolean`

Initializes the engine, including Media Foundation, WASAPI device acquisition, and the DSP render timer. 
- **Returns:** `True` if initialization succeeded. `False` if failed (check `RiffLastError`).
- **Note:** It is safe to call `RiffOpen` multiple times; subsequent calls return `True` immediately if already running.

### RiffClose
`Public Sub RiffClose()`

Shuts down the engine completely. Stops all active voices, unloads all buffers, and releases hardware interfaces. 
- **Lifecycle:** Essential to call this in `Workbook_BeforeClose` or equivalent events.

### RiffSuspend
`Public Sub RiffSuspend()`

Manually stops the render timer to save CPU while keeping all buffers and voice state in memory. 

### RiffWake
`Public Function RiffWake() As Boolean`

Restarts the render timer after suspension. Returns `True` if the timer is running.

## Diagnostics and Engine State

### RiffIsInitialized
`Public Property Get RiffIsInitialized() As Boolean`
Checks if the engine is currently active.

### RiffLastError
`Public Property Get RiffLastError() As RiffErrorCode`
Returns the most recent error code. Successful API calls clear this value.

### RiffAutoSuspendTimer
`Public Property Get/Let RiffAutoSuspendTimer() As Boolean`
If `True`, the engine will automatically call `RiffSuspend` after 50 consecutive idle ticks (approx. 750ms of silence) to save CPU and VBE stability.

### RiffMasterVolume
`Public Property Get/Let RiffMasterVolume() As Single`
Sets the global master gain multiplier. Valid range: `0.0` to `1.0`.

### RiffMasterGetPeak
`Public Sub RiffMasterGetPeak(ByRef peakLeft As Single, ByRef peakRight As Single)`
Retrieves the instantaneous peak amplitude (0.0 to 1.0+) of the final master mix. Values decay over time.

## Buffer Management

### RiffLoad
`Public Function RiffLoad(ByVal filePath As String) As Long`
Decodes an audio file into memory. Supports standard formats (WAV, MP3, AAC, FLAC, etc.).
- **Returns:** Buffer handle (`0` to `63`) or `-1` on failure.

### RiffLoadFromMemory
`Public Function RiffLoadFromMemory(ByRef audioData() As Byte) As Long`
Decodes audio from a `Byte` array. Useful for assets stored in cells, custom properties, or embedded resources.

### RiffUnload
`Public Sub RiffUnload(ByVal bufferHandle As Long)`
Frees the physical memory used by the buffer and stops any voices playing from it.

### RiffBufferDurationSec
`Public Property Get RiffBufferDurationSec(ByVal bufferHandle As Long) As Single`
Returns the length of the loaded buffer in seconds.

## Playback and Voice Control

### RiffPlay
`Public Function RiffPlay(ByVal bufferHandle As Long, Optional ByVal busID As RiffBusId = RiffBusMain, Optional ByVal looped As Boolean = False, Optional ByVal volume As Single = 1.0, Optional ByVal pan As Single = 0.0) As Long`
Spawns a new voice for the specified buffer.
- **Parameters:**
    - `busID`: Routes the voice to a specific mixer bus.
    - `looped`: If `True`, the voice wraps back to start (or loop region) on completion.
- **Returns:** Voice handle (`0` to `31`) or `-1` if the pool is full.

### RiffPlayOnce
`Public Function RiffPlayOnce(ByVal bufferHandle As Long, ...args) As Long`
Starts playback only if the same buffer is not already active on the target bus.

### RiffPlayOscillator
`Public Function RiffPlayOscillator(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ...args) As Long`
Spawns a synth voice using the engine's internal oscillators.

### RiffPlayNoise
`Public Function RiffPlayNoise(Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, ...args) As Long`
Spawns a noise generator (White, Pink, or Brown noise).

### RiffPause / RiffResume / RiffStop
Voice-level playback control. `RiffStop` immediately frees the voice slot.

### RiffFadeIn / RiffFadeOut
`Public Sub RiffFadeIn(ByVal voiceHandle As Long, ByVal durationSec As Single)`
`Public Sub RiffFadeOut(ByVal voiceHandle As Long, ByVal durationSec As Single)`
Transitions volume smoothly. `RiffFadeOut` stops the voice automatically when finished.

### RiffSetLoopRegionSec
`Public Sub RiffSetLoopRegionSec(ByVal voiceHandle As Long, ByVal startSec As Single, ByVal endSec As Single)`
Defines a sub-region for looping. Boundaries are sample-aligned.

## Audio Buses and Mixing

Audio buses act as logical summing groups or "sub-mixes." Instead of managing the volume of dozens of individual voices, you can route them to functional buses (e.g., `RiffBusMusic`, `RiffBusSfx`, `RiffBusUi`) and control the group's gain, mute state, or solo status in a single call.

### The Signal Chain
Riff uses a multiplicative volume hierarchy. If the Master volume is 50%, the Bus volume is 50%, and the Voice volume is 100%, the resulting sound will play at 25% of full scale.

1. **Voice Volume:** Set via `RiffVoiceVolume`.
2. **Bus Volume:** Set via `RiffBusVolume`.
3. **Master Volume:** Set via `RiffMasterVolume`.

### Bus Management Functions

#### RiffBusVolume
`Public Property Get/Let RiffBusVolume(ByVal busID As RiffBusId) As Single`
Sets the gain multiplier for an entire bus. The valid range is `0.0` (silence) to `2.0` (+6dB boost).

#### RiffBusMuted / RiffBusSolo
`Public Property Get/Let RiffBusMuted(ByVal busID As RiffBusId) As Boolean`
`Public Property Get/Let RiffBusMuted(ByVal busID As RiffBusId, ByVal value As Boolean)`
`Public Property Get/Let RiffBusSolo(ByVal busID As RiffBusId) As Boolean`
`Public Property Get/Let RiffBusSolo(ByVal busID As RiffBusId, ByVal value As Boolean)`
Standard mixer controls. Soloing a bus effectively mutes all other buses unless they are also soloed.

#### RiffBusFadeTo
`Public Sub RiffBusFadeTo(ByVal busID As RiffBusId, ByVal targetVolume As Single, Optional ByVal durationMs As Long = 250)`
Transitions a bus volume over time. This is processed entirely within the native audio callback and does not block the VBA thread.

#### RiffBusGetPeak
`Public Sub RiffBusGetPeak(ByVal busID As RiffBusId, ByRef peakLeft As Single, ByRef peakRight As Single)`
Retrieves the post-fader output level for a specific bus, ideal for driving UI meters.

## Voice Properties

- **Volume (`Single`):** Per-voice gain (0.0 to 2.0).
- **Pitch (`Double`):** Speed and frequency multiplier (0.1 to 8.0).
- **Pan (`Single`):** Stereo position (-1.0 Left, 1.0 Right).
- **PositionSec (`Single`):** Current playback position. Writing seeks to a new time.

## DSP Pipeline and Effects

Every voice passes through a fixed Studio DSP chain. Controls with a mix or depth of `0.0` are bypassed for performance.

### Filters & EQ
- **RiffVoiceSetFilter:** Sets Low-Pass (0.01-1.0) and High-Pass (0.0-0.99) cutoffs.
- **EQ Bands:** `RiffVoiceEqBass`, `RiffVoiceEqMid`, `RiffVoiceEqTreble` (0.05 to 5.0 gain).

### Spatial & Modulation
- **Reverb:** `RiffVoiceSetReverb(mix, roomTime)`. Uses Freeverb logic.
- **Delay:** `RiffVoiceSetDelay(time, feedback, mix)`.
- **Chorus / Flanger:** `RiffVoiceSetChorus`, `RiffVoiceSetFlanger`.
- **Stereo Width:** `RiffVoiceStereoWidth` (0.0 Mono to 5.0 Extra-Wide).

### Dynamics & Texture
- **Compressor:** `RiffVoiceCompressorThreshold`, `RiffVoiceCompressorRatio`.
- **Distortion:** `RiffVoiceDistortion` (multiplier-based hard clip).
- **Bitcrusher:** `RiffVoiceBitDepth` (2 to 32), `RiffVoiceSampleRateReduction` (1 to 20).
- **Ring Modulator:** `RiffVoiceRingModFreq`, `RiffVoiceRingModMix`.

## High-Level Effect Presets

### RiffVoiceApplyPreset
`Public Sub RiffVoiceApplyPreset(ByVal voiceHandle As Long, ByVal preset As RiffEffectPreset, Optional ByVal amount As Single = 1.0)`
Instantly configures the entire DSP matrix for a specific character. 
- **Example:** `RiffVoiceApplyPreset v, RiffFxUnderwater, 0.7` sets filter cutoffs, EQ, and chorus to simulate being submerged.

## Offline Rendering and Export

### RiffExportBufferWav
`Public Function RiffExportBufferWav(ByVal bufferHandle As Long, ByVal filePath As String) As Boolean`
Saves a decoded buffer to a high-quality 16-bit stereo PCM WAV file.

### RiffRenderOscillatorWav
`Public Function RiffRenderOscillatorWav(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ByVal durationSec As Single, ByVal filePath As String) As Boolean`
Generates a synth file without using real-time playback voices.

## Enumerations

### RiffWaveType
Selects the oscillator or noise algorithm.
- `RiffWaveSine`, `RiffWaveSquare`, `RiffWaveSawtooth`, `RiffWaveWhiteNoise`, `RiffWavePinkNoise`, `RiffWaveBrownNoise`.

### RiffBusId
Mixing bus slots.
- `RiffBusMain`, `RiffBusSfx`, `RiffBusMusic`, `RiffBusVoice`, `RiffBusUi`, `RiffBusAux1` to `RiffBusAux11`.

### RiffEffectPreset
Character-based presets for `RiffVoiceApplyPreset`.
- `RiffFxDry`, `RiffFxSmallRoom`, `RiffFxHall`, `RiffFxCathedral`, `RiffFxSlapback`, `RiffFxEcho`, `RiffFxChorus`, `RiffFxFlanger`, `RiffFxLoFi`, `RiffFxRadio`, `RiffFxUnderwater`, `RiffFxWide`, `RiffFxRobot`, `RiffFxAmbient`.
