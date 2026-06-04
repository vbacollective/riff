# Riff API Reference

This document provides a comprehensive and exhaustive guide to the public interface of **Riff.bas**, a high-performance, single-file WASAPI audio engine for VBA. It covers every public function, property, and enumeration required to integrate professional audio playback, synthesis, and real-time DSP into Microsoft Office applications.

## Table of Contents

- [Core Concepts](#core-concepts)
- [Signal Hierarchy & Routing](#signal-hierarchy--routing)
- [Initialization and Lifecycle](#initialization-and-lifecycle)
- [Engine Configuration & State](#engine-configuration--state)
- [Audio Mixer Buses](#audio-mixer-buses)
- [Buffer Management](#buffer-management)
- [Playback Entry Points](#playback-and-voice-spawning)
- [Synthesis and Oscillators](#synthesis-and-oscillators)
- [Voice Playback Control](#voice-playback-control)
- [DSP Pipeline Matrix](#dsp-pipeline-matrix)
- [Effect Presets](#effect-presets)
- [Offline Rendering and Export](#offline-rendering-and-export)
- [Enumerations](#enumerations)

## Core Concepts

### Numeric Handles
Riff uses zero-based `Long` handles to represent resources. These handles are used as indices into optimized internal arrays to ensure high-performance, sample-accurate processing.
- **Buffer Handles (0 to 63):** Represent decoded, uncompressed PCM data residing in system memory.
- **Voice Handles (0 to 31):** Represent active playback channels. A voice is spawned when you call a playback function and is recycled once it finishes or is stopped.
- **Bus IDs (0 to 15):** Represent logical mixing groups.

### Resource Limits
- **Voices:** 32 simultaneous polyphonic channels.
- **Buffers:** 64 statically allocated audio slots.
- **Buses:** 16 independent mixer buses.

## Signal Hierarchy & Routing

Riff implements a multiplicative gain hierarchy that allows for granular control over the final output without altering individual asset states.

### The Signal Chain
The final amplitude of any sample reaching the hardware is calculated as:
**Output = Sample × Voice Volume × Bus Volume × Master Volume**

1. **Voice Volume:** Set per instance (e.g., an individual explosion).
2. **Bus Volume:** Set per category (e.g., all Sound Effects).
3. **Master Volume:** Set for the entire application.

## Initialization and Lifecycle

### RiffOpen
`Public Function RiffOpen() As Boolean`
Initializes the Media Foundation decoding subsystem, acquires the default WASAPI render device, and starts the high-resolution (10ms) DSP timer thunk.
- **Returns:** `True` if initialization succeeded. `False` if failed. Check `RiffLastError` for COM or hardware errors.
- **Note:** Safe to call multiple times; subsequent calls return `True` immediately if already running.

### RiffClose
`Public Sub RiffClose()`
Shuts down the engine completely. This stops the render timer, releases all WASAPI/COM interfaces, and frees all memory allocated via `VirtualAlloc`.
- **CRITICAL:** You **must** call this in your host's shutdown event (e.g., `Workbook_BeforeClose`). Failure to close the engine leaves a native timer running in the background, which will likely crash the host on project reset.

### RiffSuspend
`Public Sub RiffSuspend()`
Pauses the native render timer to save CPU cycles. All buffers and voice states are preserved in memory. Useful when the application is minimized or inactive.

### RiffWake
`Public Function RiffWake() As Boolean`
Restarts the render timer after suspension.
- **Returns:** `True` if the timer is successfully running.

## Engine Configuration & State

### RiffIsInitialized
`Public Property Get RiffIsInitialized() As Boolean`
Indicates whether the engine has been successfully opened and is ready to process audio.

### RiffLastError
`Public Property Get RiffLastError() As RiffErrorCode`
Returns the code of the most recent failure. Successful API calls reset this to `RiffErrorNone`.

### RiffAutoSuspendTimer
`Public Property Get/Let RiffAutoSuspendTimer() As Boolean`
When enabled (`True`), the engine automatically suspends its render timer after 50 consecutive idle ticks (approx. 750ms of silence). It automatically wakes up when a new playback command is issued.

### RiffMasterVolume
`Public Property Get/Let RiffMasterVolume() As Single`
Sets the final gain multiplier for the entire engine.
- **Range:** `0.0` (Mute) to `1.0` (Unity). Defaults to `1.0`.

### RiffMasterGetPeak
`Public Sub RiffMasterGetPeak(ByRef peakLeft As Single, ByRef peakRight As Single)`
Retrieves the instantaneous peak amplitude of the master output.
- **Parameters:** `peakLeft` and `peakRight` are passed by reference and updated with values from `0.0` to `1.0+` (clipping).

### Engine Constants
- `RiffMaxVoices`: Returns `32`.
- `RiffMaxBuffers`: Returns `64`.
- `RiffMaxBuses`: Returns `16`.

## Audio Mixer Buses

Buses act as logical summing groups for routing and mixing.

### RiffBusVolume
`Public Property Get/Let RiffBusVolume(ByVal busID As RiffBusId) As Single`
Sets the volume for an entire bus.
- **Range:** `0.0` to `2.0` (+6dB boost). Defaults to `1.0`.

### RiffBusMuted
`Public Property Get/Let RiffBusMuted(ByVal busID As RiffBusId) As Boolean`
Silences a bus without altering its `RiffBusVolume` setting.

### RiffBusSolo
`Public Property Get/Let RiffBusSolo(ByVal busID As RiffBusId) As Boolean`
When one or more buses are in Solo mode, only those buses (and any others also in Solo) will be audible. All other non-soloed buses are implicitly muted.

### RiffBusFadeTo
`Public Sub RiffBusFadeTo(ByVal busID As RiffBusId, ByVal targetVolume As Single, Optional ByVal durationMs As Long = 250)`
Smoothly transitions a bus volume. This is processed entirely in the native background thread and is non-blocking.

### RiffBusGetPeak
`Public Sub RiffBusGetPeak(ByVal busID As RiffBusId, ByRef peakLeft As Single, ByRef peakRight As Single)`
Retrieves the post-fader output level for a specific bus.

### RiffBusReset
`Public Sub RiffBusReset(ByVal busID As RiffBusId)`
Restores a bus to its default state (Volume 1.0, Unmuted, No Solo, No Fades).

## Buffer Management

Buffers store decoded audio data in memory for instant playback.

### RiffLoad
`Public Function RiffLoad(ByVal filePath As String) As Long`
Decodes an audio file into a free buffer slot. Supports WAV, MP3, AAC, FLAC, and OGG (on Windows 10+).
- **Returns:** Buffer handle (`0-63`) or `-1` on failure.

### RiffLoadFromMemory
`Public Function RiffLoadFromMemory(ByRef audioData() As Byte) As Long`
Decodes audio from a raw binary array.
- **Note:** The array must contain a valid file format (e.g., the contents of an MP3 file), not raw PCM samples.

### RiffUnload
`Public Sub RiffUnload(ByVal bufferHandle As Long)`
Kills any voices using the buffer and frees the associated physical memory.

### RiffBufferDurationSec
`Public Property Get RiffBufferDurationSec(ByVal bufferHandle As Long) As Single`
Returns the total length of the buffer in seconds.

## Playback and Voice Spawning

These functions create a new **Voice** to play a loaded **Buffer**.

### RiffPlay
`Public Function RiffPlay(ByVal bufferHandle As Long, Optional ByVal busID As RiffBusId = RiffBusMain, Optional ByVal looped As Boolean = False, Optional ByVal volume As Single = 1.0, Optional ByVal pan As Single = 0.0) As Long`
The primary playback function.
- **Returns:** Voice handle (`0-31`) or `-1` if the pool is full.

### RiffPlayOnce
`Public Function RiffPlayOnce(ByVal bufferHandle As Long, ...)`
Starts playback **only if** the same buffer is not already playing on the target bus. Useful for preventing "phasing" sounds or duplicate UI clicks.

### RiffPlayBus
`Public Function RiffPlayBus(ByVal bufferHandle As Long, ByVal busID As RiffBusId, ...)`
A compatibility wrapper for routing a buffer directly to a bus.

## Synthesis and Oscillators

Oscillators generate sound mathematically and do not require a loaded buffer. They use the same DSP pipeline as buffers.

### RiffPlayOscillator
`Public Function RiffPlayOscillator(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ...args) As Long`
Spawns a periodic synth voice (Sine, Square, Saw).

### RiffPlayNoise
`Public Function RiffPlayNoise(Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, ...args) As Long`
Spawns a noise generator.
- **Types:** White (random), Pink (1/f), or Brown (1/f²).

## Voice Playback Control

Once a voice handle is obtained, you can manipulate its playback state in real-time.

### Status and Identification
- **RiffVoiceIsPlaying:** Returns `True` if the voice is active and not paused.
- **RiffVoiceIsPaused:** Returns `True` if the voice is held in a paused state.
- **RiffFindPlayingVoice:** Finds an active voice handle for a given buffer.

### Transport Control
- **RiffPause / RiffResume:** Suspends or continues playback.
- **RiffStop:** Immediately kills the voice and clears its DSP state.
- **RiffStopAll:** Kills all active voices across all buses.

### Envelopes and Regions
- **RiffFadeIn / RiffFadeOut:** Applies a linear gain ramp over a specified duration in seconds.
- **RiffSetLoopRegionSec:** Defines the start and end points for looping within a buffer.

### Real-time Parameters
- **RiffVoiceVolume:** Per-voice gain (`0.0` to `2.0`).
- **RiffVoicePitch:** Pitch and speed multiplier (`0.1` to `8.0`). `1.0` is normal.
- **RiffVoicePan:** Stereo position from `-1.0` (Full Left) to `1.0` (Full Right).
- **RiffVoicePositionSec:** Gets or sets the current playback cursor in seconds.

## DSP Pipeline Matrix

Every voice passes through a fixed chain of DSP stages. Setting a "Mix" or "Depth" property to `0.0` bypasses that stage for performance.

### 1. Bitcrusher & Distortion
- **RiffVoiceBitDepth:** Simulates lower bit depths (`2` to `32`).
- **RiffVoiceSampleRateReduction:** Simple integer downsampling (`1` to `20`).
- **RiffVoiceDistortion:** Multiplier-based hard clipping.

### 2. Filters & EQ
- **RiffVoiceLowPass / RiffVoiceHighPass:** Resonance-compensated Biquad filters.
- **RiffVoiceSetFilter:** Sets both Low-Pass and High-Pass simultaneously.
- **RiffVoiceEqBass / Mid / Treble:** Gain multipliers for a 3-band parametric EQ.

### 3. Modulation Effects
- **RiffVoiceRingModFreq / Mix:** Metallic, robotic ring modulation.
- **RiffVoiceTremoloRate / Depth:** Periodic volume oscillation.
- **RiffVoiceAutoPanRate / Depth:** Periodic stereo movement.
- **RiffVoiceChorusRate / Depth:** Shimmering, thick ensemble effect.
- **RiffVoiceFlangerRate / Depth / Feedback:** Resonant, sweeping "jet" effect.

### 4. Spatial Effects
- **RiffVoiceReverbMix / Time:** High-quality Freeverb implementation.
- **RiffVoiceDelayTime / Feedback / Mix:** Standard echo/delay. Max delay is ~1.9s.
- **RiffVoiceStereoWidth:** Controls the side-channel gain. `0.0` is Mono, `1.0` is standard Stereo, `5.0` is ultra-wide.

### 5. Dynamics
- **RiffVoiceCompressorThreshold / Ratio:** Real-time gain reduction with automatic attack/release.

## Effect Presets

### RiffVoiceApplyPreset
`Public Sub RiffVoiceApplyPreset(ByVal voiceHandle As Long, ByVal preset As RiffEffectPreset, Optional ByVal amount As Single = 1.0)`
A powerful high-level function that instantly configures multiple DSP stages to achieve a specific character.
- **Presets:** `RiffFxRadio`, `RiffFxUnderwater`, `RiffFxCathedral`, `RiffFxLoFi`, `RiffFxRobot`, `RiffFxAmbient`, etc.
- **Note:** This resets all other DSP properties to neutral before applying the preset.

## Offline Rendering and Export

### RiffExportBufferWav
`Public Function RiffExportBufferWav(ByVal bufferHandle As Long, ByVal filePath As String) As Boolean`
Writes the contents of a decoded buffer to a standard 16-bit stereo PCM WAV file.

### RiffRenderOscillatorWav
`Public Function RiffRenderOscillatorWav(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ByVal durationSec As Single, ByVal filePath As String) As Boolean`
Generates a synthesized audio asset directly to disk without requiring a real-time voice.

## Enumerations

### RiffBusId
Defines the standard mixer slots.
`RiffBusMain`, `RiffBusSfx`, `RiffBusMusic`, `RiffBusVoice`, `RiffBusUi`, `RiffBusAux1`...`RiffBusAux11`.

### RiffWaveType
Defines oscillator and noise algorithms.
`RiffWaveSine`, `RiffWaveSquare`, `RiffWaveSawtooth`, `RiffWaveWhiteNoise`, `RiffWavePinkNoise`, `RiffWaveBrownNoise`.

### RiffEffectPreset
Defines the built-in DSP configurations for `RiffVoiceApplyPreset`.
`RiffFxDry`, `RiffFxSmallRoom`, `RiffFxHall`, `RiffFxCathedral`, `RiffFxSlapback`, `RiffFxEcho`, `RiffFxChorus`, `RiffFxFlanger`, `RiffFxLoFi`, `RiffFxRadio`, `RiffFxUnderwater`, `RiffFxWide`, `RiffFxRobot`, `RiffFxAmbient`.

### RiffErrorCode
Diagnostic values returned by `RiffLastError`.
`RiffErrorNone`, `RiffErrorNotInitialized`, `RiffErrorNoFreeBuffer`, `RiffErrorNoFreeVoice`, `RiffErrorInvalidBuffer`, `RiffErrorInvalidVoice`, `RiffErrorInvalidBus`, `RiffErrorInvalidArgument`, `RiffErrorFileNotFound`, `RiffErrorComFailure`, `RiffErrorMemoryAllocation`, `RiffErrorDecodeFailed`, `RiffErrorUnsupportedFormat`.
