# Riff API Reference

This document describes the complete public interface of **Riff.bas**, a high-performance, COM-based WASAPI audio engine for VBA (x86/x64 compatible). It implements real-time audio playback through Windows Audio Session API (WASAPI), decodes audio files via Media Foundation, and provides a full Studio DSP Pipeline including Freeverb-style Reverb, Chorus, Flanger, Compressor, Biquad 3-Band EQ, Bitcrusher, Ring Modulator, Auto-Pan, Delay, BLEP Oscillators, In-Memory Loading, WAV Export, Audio Buses, and Peak Meters.

## Table of Contents

- [Core Concepts](#core-concepts)
  - [Engine Initialization](#engine-initialization)
  - [Buffer Pool](#buffer-pool)
  - [Voice Pool and Polyphony](#voice-pool-and-polyphony)
  - [Audio Buses](#audio-buses)
  - [WAV Export](#wav-export)
  - [DSP Pipeline](#dsp-pipeline)
  - [Timer and Callback Architecture](#timer-and-callback-architecture)
  - [Platform Compatibility](#platform-compatibility)
- [Initialization and Teardown](#initialization-and-teardown)
- [Global Settings and Asset Management](#global-settings-and-asset-management)
- [Export and Offline Rendering](#export-and-offline-rendering)
- [Playback and Voice Actions](#playback-and-voice-actions)
- [Voice Properties](#voice-properties)
- [DSP Filters and Effects](#dsp-filters-and-effects)
  - [Bitcrusher](#bitcrusher)
  - [Distortion](#distortion)
  - [Low-Pass and High-Pass Filters](#low-pass-and-high-pass-filters)
  - [3-Band EQ](#3-band-eq)
  - [Compressor](#compressor)
  - [Stereo Width](#stereo-width)
  - [Tremolo](#tremolo)
  - [Auto-Pan](#auto-pan)
  - [Ring Modulator](#ring-modulator)
  - [Chorus](#chorus)
  - [Flanger](#flanger)
  - [Reverb](#reverb)
  - [Delay](#delay)
- [Practical Patterns](#practical-patterns)
- [Operational Caveats](#operational-caveats)

## Core Concepts

### Engine Initialization

Riff initializes the full audio stack in a single call to `RiffOpen`. Internally, it performs the following sequence:

1. Allocates and compiles a native machine-code thunk via `VirtualAlloc` with `PAGE_EXECUTE_READWRITE`.
2. Calls `MFStartup` to initialize the Windows Media Foundation runtime.
3. Creates the WASAPI device enumerator, acquires the default audio output device, and obtains an `IAudioClient` interface.
4. Queries the device's native mix format (`GetMixFormat`) to determine sample rate, bit depth, and channel count.
5. Initializes the audio client in Shared Mode with a 150 ms buffer.
6. Obtains the `IAudioRenderClient` and starts the audio stream.
7. Allocates a contiguous 1D ring buffer array (`rRingBuf`) of 32 × 192,000 `Single` samples for spatial effects.
8. Sets the system timer resolution to 1 ms via `timeBeginPeriod`.
9. Starts a `SetTimer` callback at 15 ms intervals to drive the DSP loop.

> [!IMPORTANT]
> All playback, DSP, and asset operations silently exit if `RiffOpen` has not been called or if it returned `False`. Always check the return value before proceeding.

```vb
If Not RiffOpen() Then
    MsgBox "Riff failed to initialize. Check audio device availability."
    Exit Sub
End If
```

### Buffer Pool

Riff maintains an internal pool of up to **64 static audio buffers** (indices `0` to `63`). Each buffer holds raw PCM data decoded from an audio file and allocated in physical memory via `VirtualAlloc`.

`RiffLoad` and `RiffLoadFromMemory` decode a source file (WAV, MP3, or any format supported by Media Foundation) into a free slot and return its integer handle. This handle is used to instantiate playback voices.

> [!NOTE]
> Decoded audio is converted to the device's native mix format at load time. The engine performs no format conversion at playback, which is what enables zero-latency voice spawning. Export helpers convert loaded buffers back to standard 16-bit stereo PCM WAV when writing files.

### Voice Pool and Polyphony

Riff maintains a pool of **32 polyphonic voices** (indices `0` to `31`). Each voice is an independent playback channel with its own complete DSP pipeline state. A voice is activated by calling `RiffPlay` or `RiffPlayOscillator` and is automatically freed when playback finishes.

Each call to `RiffPlay` scans for the first inactive voice and returns its integer handle. This handle is used to control volume, pitch, pan, looping, fades, and every DSP parameter for that specific instance.

> [!CAUTION]
> If all 32 voices are active, `RiffPlay` and `RiffPlayOscillator` return `-1` and no sound is produced. Monitor active voice count in high-density scenarios.

```vb
Dim v As Long
v = RiffPlay(myBuffer)

If v = -1 Then
    Debug.Print "No voice available"
Else
    RiffVoiceVolume(v) = 0.8
    RiffVoicePan(v) = -0.5
End If
```

### Audio Buses

Riff provides **8 audio buses** (indices `0` to `7`) as an additional global volume layer applied per voice. Every voice routes to Bus 0 by default. You can assign voices to different buses and control each bus volume independently, enabling logical groups such as Music, SFX, Voice, and UI.

```vb
' Assign a voice to Bus 1 (SFX group)
RiffVoiceBus(sfxVoice) = 1

' Lower the SFX bus without touching other groups
RiffBusVolume(1) = 0.4
```

Bus volume is applied as a multiplier after the individual voice volume and before the master volume. The effective gain for a frame is:

```
finalGain = VoiceVolume × BusVolume × MasterVolume
```

### WAV Export

Riff can write audio back to disk as standard **16-bit stereo PCM WAV**. Loaded buffers can be exported with `RiffExportBufferWav`, and oscillators can be rendered directly with `RiffRenderOscillatorWav`.

Export is intentionally conservative: the output format is always stereo PCM16 for maximum compatibility with DAWs, video editors, game tools, and Windows media players. When exporting a loaded buffer, Riff reads the current decoded mix format and converts supported source layouts to stereo PCM16.

```vb
Dim buf As Long
buf = RiffLoad("C:\Sounds\theme.mp3")

If buf >= 0 Then
    RiffExportBufferWav buf, "C:\Sounds\theme_export.wav"
End If
```

### DSP Pipeline

Every voice processes audio through a fixed pipeline applied in the following order per sample:

1. **Source Read**: Buffer lookup with pitch-adjusted index or oscillator generation.
2. **Sample Rate Reduction**: Bitcrusher downsampling (hold N frames).
3. **Bit Depth Reduction**: Bitcrusher quantization.
4. **Distortion**: Soft-clip multiplication and hard clamp.
5. **Low-Pass Filter**: Biquad low-pass filtering for smoother cutoff behavior.
6. **High-Pass Filter**: Biquad high-pass filtering for controlled low-frequency removal.
7. **3-Band EQ**: Biquad bass, mid, and treble bands using per-voice filter state.
8. **Ring Modulator**: Sine oscillator multiplication with wet mix blend.
9. **Tremolo**: Volume LFO modulation.
10. **Stereo Width**: Mid/side processing.
11. **Flanger**: Short comb filter with LFO-swept delay tap from the ring buffer.
12. **Chorus**: Modulated delay tap blend from the ring buffer.
13. **Echo/Delay**: Fixed-time delay tap with feedback from the ring buffer.
14. **Reverb**: Freeverb-style comb network with damping and stereo spread.
15. **Ring Buffer Write**: Stores the pre-compressor wet signal for spatial effect feedback.
16. **Compressor**: Envelope follower with threshold and ratio gain reduction.
17. **Auto-Pan**: LFO-modulated pan position applied to final L/R gain.
18. **Volume and Master Gain**: Voice, bus, and master multipliers applied.
19. **Fade**: Smooth linear fade-in or fade-out multiplier.
20. **Mix Accumulation**: Output added to the shared mix buffer.

> [!NOTE]
> The ring buffer used for Flanger, Chorus, Delay, and Reverb is per-voice and 192,000 samples long (approximately 4 seconds at 48 kHz). Effects with long delay times may produce silence initially until the buffer fills.

### Timer and Callback Architecture

Riff uses a native machine-code thunk (`InitThunks`) compiled at runtime into executable memory. This thunk is registered with `SetTimer` at a 15 ms interval and calls `RiffTimerCallback` on each tick. The thunk also checks `EbMode` from the VBA runtime to determine whether VBA is in a break or design state and kills the timer automatically if so, preventing crashes on project reset.

> [!WARNING]
> Always call `RiffClose` before resetting the VBA project or closing the workbook. The thunk holds a pointer into the VBA runtime. Resetting without cleanup can crash the host application.

### Platform Compatibility

Riff is fully compatible with both **32-bit VBA** (Office x86) and **64-bit VBA** (Office x64). All pointer types are declared conditionally with `#If VBA7` / `#If Win64` compiler directives. The DSP engine paths for 16-bit PCM and 32-bit float WASAPI output are implemented, with conservative handling for unsupported formats to avoid writing invalid output data.

## Initialization and Teardown

### RiffOpen

```vb
Public Function RiffOpen() As Boolean
```

Initializes the complete audio engine: Media Foundation, WASAPI device acquisition, the DSP timer, and the ring buffer. Calling this function when the engine is already running is safe and returns `True` immediately.

**Returns:** `True` if the engine started successfully. `False` if any internal step failed (device unavailable, Media Foundation error, timer allocation failure, etc.).

```vb
If Not RiffOpen() Then
    MsgBox "Audio engine failed to start."
    Exit Sub
End If
```

> [!IMPORTANT]
> Call `RiffOpen` once at workbook open or before first use. Do not call it repeatedly in a loop.

### RiffClose

```vb
Public Sub RiffClose()
```

Shuts down the audio engine completely. Stops all active voices, kills the DSP timer, frees all static buffers from physical memory, releases all WASAPI and Media Foundation COM interfaces, and shuts down `MFShutdown`. Also restores the system timer resolution via `timeEndPeriod`.

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RiffClose
End Sub
```

> [!WARNING]
> Failing to call `RiffClose` before a VBA project reset or workbook close may leave the native thunk timer active, which can crash Excel.

### RiffIsInitialized

```vb
Public Property Get RiffIsInitialized() As Boolean
```

Returns `True` if the engine has been successfully initialized and is currently running.

```vb
If Not RiffIsInitialized Then
    RiffOpen
End If
```

## Global Settings and Asset Management

### RiffMasterVolume

```vb
Public Property Get RiffMasterVolume() As Single
Public Property Let RiffMasterVolume(ByVal value As Single)
```

Gets or sets the global master output volume applied to all voices across all buses. Valid range is `0.0` (silent) to `1.0` (full volume). Values outside this range are clamped automatically.

```vb
RiffMasterVolume = 0.75
```

### RiffBusVolume

```vb
Public Property Get RiffBusVolume(ByVal busID As Long) As Single
Public Property Let RiffBusVolume(ByVal busID As Long, ByVal value As Single)
```

Gets or sets the volume multiplier for the specified audio bus. `busID` must be between `0` and `7`. Valid range is `0.0` to `2.0` (values above `1.0` amplify). Values out of range are clamped.

```vb
' Silence the music bus
RiffBusVolume(0) = 0.0

' Boost the SFX bus slightly
RiffBusVolume(1) = 1.2
```

### RiffMasterGetPeak

```vb
Public Sub RiffMasterGetPeak(ByRef peakLeft As Single, ByRef peakRight As Single)
```

Retrieves the current peak amplitude of the master output mix for both left and right channels. Values range from `0.0` to `1.0` (or above `1.0` if clipping occurs). Peak values decay at approximately 10% per timer tick. Use this to drive VU meters or clipping indicators.

```vb
Dim pL As Single, pR As Single
RiffMasterGetPeak pL, pR
Debug.Print "Master Peak L/R:", pL, pR
```

### RiffLoad

```vb
Public Function RiffLoad(ByVal filePath As String) As Long
```

Decodes an audio file from disk into physical memory using Media Foundation. Supports any format that Media Foundation can decode on the target system (WAV, MP3, AAC, FLAC, WMA, and others depending on installed codecs). The audio is resampled and converted to the device's native mix format at load time.

**Parameters:**

`filePath`: Full path to the audio file on disk.

**Returns:** A buffer handle (`0` to `63`) on success, or `-1` if the file was not found, could not be decoded, or the buffer pool is full.

```vb
Dim buf As Long
buf = RiffLoad("C:\Sounds\explosion.wav")

If buf = -1 Then
    Debug.Print "Failed to load audio file."
End If
```

> [!NOTE]
> Loading is synchronous and may take a moment for large files. Decode all assets at startup rather than during real-time playback.

### RiffLoadFromMemory

```vb
Public Function RiffLoadFromMemory(ByRef audioData() As Byte) As Long
```

Decodes audio directly from a `Byte` array in memory, without requiring any file on disk. Internally wraps the array in an `IStream` via `SHCreateMemStream`, then passes it to Media Foundation for decoding. Useful for embedding audio in the workbook as a VBA constant or Base64-decoded resource.

**Parameters:**

`audioData`: A `Byte` array containing the complete binary content of a supported audio file.

**Returns:** A buffer handle (`0` to `63`) on success, or `-1` on failure.

```vb
Dim audioBytes() As Byte
' ... populate audioBytes from a resource or network source ...
Dim buf As Long
buf = RiffLoadFromMemory(audioBytes)
```

### RiffUnload

```vb
Public Sub RiffUnload(ByVal bufferHandle As Long)
```

Removes a loaded audio buffer from the pool, frees its physical memory via `VirtualFree`, and immediately deactivates any voice currently playing from that buffer. The buffer slot becomes available for a new `RiffLoad` call.

```vb
RiffUnload myBuffer
```

> [!WARNING]
> All voices referencing `bufferHandle` are stopped immediately when `RiffUnload` is called, regardless of their current playback state.

### RiffBufferDurationSec

```vb
Public Property Get RiffBufferDurationSec(ByVal bufferHandle As Long) As Single
```

Returns the total duration in seconds of a loaded static buffer, calculated from the buffer's byte length and the device's average bytes-per-second rate.

```vb
Debug.Print "Duration:", RiffBufferDurationSec(myBuffer), "seconds"
```

## Export and Offline Rendering

### RiffExportBufferWav

```vb
Public Function RiffExportBufferWav(ByVal bufferHandle As Long, ByVal filePath As String) As Boolean
```

Exports a loaded buffer to a standard **16-bit stereo PCM WAV** file. The source buffer must already be loaded with `RiffLoad` or `RiffLoadFromMemory`. Mono sources are duplicated to stereo during export. Stereo sources are preserved as stereo.

**Parameters:**

`bufferHandle`: A valid buffer handle from the static buffer pool.

`filePath`: Destination path for the WAV file. Existing files may be overwritten by VBA's binary file output path.

**Returns:** `True` if the WAV file was written successfully. `False` if the engine is not initialized, the buffer handle is invalid, the buffer is empty, or the output path cannot be written.

```vb
Dim buf As Long
buf = RiffLoad("C:\Sounds\voice.ogg")

If buf >= 0 Then
    If Not RiffExportBufferWav(buf, "C:\Sounds\voice_export.wav") Then
        Debug.Print "Export failed."
    End If
End If
```

> [!NOTE]
> Export writes PCM16 stereo WAV for compatibility. It does not encode MP3, OGG, AAC, or FLAC. Use an external encoder if compressed output is required.

### RiffRenderOscillatorWav

```vb
Public Function RiffRenderOscillatorWav(ByVal waveType As Long, ByVal frequencyHz As Single, ByVal durationSec As Single, ByVal filePath As String) As Boolean
```

Renders a generated oscillator directly to a **16-bit stereo PCM WAV** file without creating a playback voice. This is useful for generating test tones, UI beeps, retro SFX, and waveform assets directly from VBA.

**Parameters:**

`waveType`: Oscillator waveform. `0` = Sine, `1` = Square, `2` = Sawtooth, `3` = Noise. Square and sawtooth are band-limited with BLEP to reduce high-frequency aliasing.

`frequencyHz`: Oscillator frequency in Hz. Values below `1.0` are normalized to `440.0`.

`durationSec`: Render duration in seconds. Must be greater than `0`.

`filePath`: Destination path for the WAV file.

**Returns:** `True` if the WAV file was written successfully. `False` if the engine is not initialized, the duration is invalid, or the output path cannot be written.

```vb
RiffRenderOscillatorWav 0, 440, 2, "C:\Tones\sine_a4.wav"
RiffRenderOscillatorWav 1, 220, 1, "C:\Tones\square_a3.wav"
RiffRenderOscillatorWav 2, 110, 1, "C:\Tones\saw_a2.wav"
```

> [!NOTE]
> Offline oscillator rendering writes a dry oscillator signal. It does not run the per-voice real-time DSP chain, because it does not allocate a `RiffVoice`.

## Playback and Voice Actions

### RiffPlay

```vb
Public Function RiffPlay(ByVal bufferHandle As Long) As Long
```

Spawns a new voice on the first available slot in the voice pool and begins playback of the specified buffer from the beginning. All DSP parameters are reset to neutral defaults.

**Parameters:**

`bufferHandle`: A valid buffer handle obtained from `RiffLoad` or `RiffLoadFromMemory`.

**Returns:** A voice handle (`0` to `31`) on success, or `-1` if the buffer is invalid or all 32 voices are occupied.

```vb
Dim v As Long
v = RiffPlay(explosionBuffer)

If v >= 0 Then
    RiffVoiceVolume(v) = 0.9
    RiffVoicePan(v) = 0.3
End If
```

### RiffPlayOscillator

```vb
Public Function RiffPlayOscillator(ByVal waveType As Long, ByVal frequencyHz As Single) As Long
```

Spawns a voice that generates audio mathematically from a waveform oscillator instead of reading from a buffer. The oscillator runs indefinitely until stopped manually, since it has no natural end point. Square and sawtooth oscillators use BLEP band-limiting to reduce aliasing at high frequencies.

**Parameters:**

`waveType`: Selects the waveform shape. Valid values:

| Value | Waveform |
|:---|:---|
| `0` | Sine: smooth, pure tone |
| `1` | Square: hollow, buzzy tone at half amplitude |
| `2` | Sawtooth: bright, aggressive ramp wave |
| `3` | Noise: random white noise |

`frequencyHz`: The pitch in Hz. Minimum value is `1.0`. Defaults to `440.0` Hz (concert A) if a value below `1.0` is passed.

**Returns:** A voice handle (`0` to `31`) on success, or `-1` if all voices are occupied.

```vb
Dim osc As Long
osc = RiffPlayOscillator(0, 440.0) ' Sine at A4

RiffVoiceVolume(osc) = 0.5
RiffVoiceReverbMix(osc) = 0.3

' Stop after 2 seconds
Application.OnTime Now + TimeValue("00:00:02"), "StopOscillator"
```

> [!NOTE]
> Oscillators do not support loop regions or `RiffVoicePositionSec` (those properties are silently ignored). Use `RiffStop` to end them.

### RiffPause

```vb
Public Sub RiffPause(ByVal voiceHandle As Long)
```

Temporarily suspends processing of the specified voice. The voice remains allocated in the pool and retains all its DSP state, position, and settings. No audio is output while paused.

```vb
RiffPause myVoice
```

### RiffResume

```vb
Public Sub RiffResume(ByVal voiceHandle As Long)
```

Resumes a previously paused voice from its current position.

```vb
RiffResume myVoice
```

### RiffStop

```vb
Public Sub RiffStop(ByVal voiceHandle As Long)
```

Immediately halts and frees the specified voice. The slot is returned to the pool and becomes available for a new `RiffPlay` call.

```vb
RiffStop myVoice
```

### RiffStopAll

```vb
Public Sub RiffStopAll()
```

Immediately halts and frees all 32 voices in the pool simultaneously.

```vb
RiffStopAll
```

### RiffFadeIn

```vb
Public Sub RiffFadeIn(ByVal voiceHandle As Long, ByVal durationSec As Single)
```

Applies a linear fade-in to the voice over the specified duration. The fade multiplier starts at `0.0` and reaches `1.0` at the end of `durationSec`. The voice's `Volume` property is not modified; the fade is an independent multiplier applied at the final gain stage.

```vb
Dim v As Long
v = RiffPlay(ambienceBuffer)
RiffVoiceLoop(v) = True
RiffFadeIn v, 3.0  ' Fade in over 3 seconds
```

### RiffFadeOut

```vb
Public Sub RiffFadeOut(ByVal voiceHandle As Long, ByVal durationSec As Single)
```

Applies a linear fade-out to the voice over the specified duration. The multiplier goes from `1.0` to `0.0`, after which the voice is automatically stopped and freed.

```vb
RiffFadeOut myVoice, 2.0  ' Fade out over 2 seconds, then stop
```

### RiffSetLoopRegionSec

```vb
Public Sub RiffSetLoopRegionSec(ByVal voiceHandle As Long, ByVal startSec As Single, ByVal endSec As Single)
```

Constrains looped playback to a specific region within the buffer, defined in seconds. When `RiffVoiceLoop` is `True` and the playback position reaches `endSec`, it wraps back to `startSec` rather than the beginning of the file. Boundaries are aligned to the device's block alignment automatically.

```vb
RiffVoiceLoop(v) = True
RiffSetLoopRegionSec v, 1.5, 4.8  ' Loop between 1.5s and 4.8s
```

> [!NOTE]
> This function has no effect on oscillator voices. The `loopEnd` boundary is clamped to the buffer's total length if `endSec` exceeds it.

## Voice Properties

### RiffVoiceIsPlaying

```vb
Public Property Get RiffVoiceIsPlaying(ByVal voiceHandle As Long) As Boolean
```

Returns `True` if the voice is active and not paused. Returns `False` if the voice has finished, been stopped, or is paused.

```vb
Do While RiffVoiceIsPlaying(v)
    DoEvents
Loop
Debug.Print "Playback finished."
```

### RiffVoiceIsPaused

```vb
Public Property Get RiffVoiceIsPaused(ByVal voiceHandle As Long) As Boolean
```

Returns `True` if the voice is currently in a paused state.

### RiffVoiceBus

```vb
Public Property Get RiffVoiceBus(ByVal voiceHandle As Long) As Long
Public Property Let RiffVoiceBus(ByVal voiceHandle As Long, ByVal value As Long)
```

Gets or sets the audio bus this voice routes to. Valid range is `0` to `7`. Defaults to `0`.

```vb
RiffVoiceBus(v) = 2  ' Route to bus 2 (e.g., "Voice" group)
```

### RiffVoiceGetPeak

```vb
Public Sub RiffVoiceGetPeak(ByVal voiceHandle As Long, ByRef peakLeft As Single, ByRef peakRight As Single)
```

Retrieves the current instantaneous peak amplitude for this voice's left and right channels after all DSP and gain stages. Use for per-voice VU meters. Values decay at approximately 10% per timer tick.

```vb
Dim pL As Single, pR As Single
RiffVoiceGetPeak v, pL, pR
```

### RiffVoiceLoop

```vb
Public Property Get RiffVoiceLoop(ByVal voiceHandle As Long) As Boolean
Public Property Let RiffVoiceLoop(ByVal voiceHandle As Long, ByVal value As Boolean)
```

Gets or sets whether the voice loops automatically when it reaches its end (or loop region end). Defaults to `False`.

```vb
RiffVoiceLoop(v) = True
```

### RiffVoicePositionSec

```vb
Public Property Get RiffVoicePositionSec(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoicePositionSec(ByVal voiceHandle As Long, ByVal value As Single)
```

Gets or sets the current playback position in seconds. Writing to this property performs a seek: the internal byte position is recalculated and aligned to the device's block size. Clamped to the valid buffer range.

```vb
' Skip to the 5-second mark
RiffVoicePositionSec(v) = 5.0

Debug.Print "Current position:", RiffVoicePositionSec(v)
```

> [!NOTE]
> Has no effect on oscillator voices.

### RiffVoiceVolume

```vb
Public Property Get RiffVoiceVolume(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceVolume(ByVal voiceHandle As Long, ByVal value As Single)
```

Gets or sets the individual volume for this voice. Valid range is `0.0` to `1.0`. Defaults to `1.0`.

```vb
RiffVoiceVolume(v) = 0.6
```

### RiffVoicePitch

```vb
Public Property Get RiffVoicePitch(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoicePitch(ByVal voiceHandle As Long, ByVal value As Single)
```

Gets or sets the playback speed and pitch multiplier. `1.0` is normal speed. `2.0` doubles the speed and raises pitch by one octave. `0.5` halves the speed and lowers pitch by one octave. Minimum value is `0.1`. Defaults to `1.0`.

```vb
RiffVoicePitch(v) = 1.5  ' 50% faster, higher pitch
```

> [!NOTE]
> Pitch shifting is achieved via index stepping (resampling), not a pitch-preserving algorithm. Changing pitch also changes playback duration.

### RiffVoicePan

```vb
Public Property Get RiffVoicePan(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoicePan(ByVal voiceHandle As Long, ByVal value As Single)
```

Gets or sets the stereo panning position. `-1.0` is fully left, `0.0` is center, `1.0` is fully right. Panning attenuates the opposite channel linearly. Defaults to `0.0`.

```vb
RiffVoicePan(v) = -0.8  ' Mostly left
```

## DSP Filters and Effects

### Bitcrusher

The Bitcrusher consists of two independent effects that can be used separately or together to simulate retro, lo-fi, or degraded audio.

#### RiffVoiceBitDepth

```vb
Public Property Get RiffVoiceBitDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceBitDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

Sets the effective bit depth used to quantize the audio signal. At `32` (default), no quantization is applied. Lower values produce an increasingly stepped, crunchy sound. Minimum value is `2`. The internal quantization step count is computed as `2 ^ value`.

```vb
RiffVoiceBitDepth(v) = 8   ' 8-bit retro quality
RiffVoiceBitDepth(v) = 4   ' Extreme lo-fi
RiffVoiceBitDepth(v) = 32  ' Disabled (full quality)
```

#### RiffVoiceSampleRateReduction

```vb
Public Property Get RiffVoiceSampleRateReduction(ByVal voiceHandle As Long) As Long
Public Property Let RiffVoiceSampleRateReduction(ByVal voiceHandle As Long, ByVal value As Long)
```

Holds each audio sample for `N` consecutive frames before advancing, simulating a lower effective sample rate. `1` (default) disables the effect. Higher values produce robotic, stepped artifacts.

```vb
RiffVoiceSampleRateReduction(v) = 4  ' Simulate ~12 kHz at 48 kHz device rate
```

### Distortion

#### RiffVoiceDistortion

```vb
Public Property Get RiffVoiceDistortion(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDistortion(ByVal voiceHandle As Long, ByVal value As Single)
```

Amplifies the signal by the given multiplier before hard-clipping it to the `[-1.0, 1.0]` range, producing digital distortion. `1.0` (default) applies no clipping. Values above `1.0` increase harmonic saturation.

```vb
RiffVoiceDistortion(v) = 3.0  ' Moderate overdrive
RiffVoiceDistortion(v) = 10.0 ' Heavy clipping / fuzz
```

### Low-Pass and High-Pass Filters

#### RiffVoiceLowPass

```vb
Public Property Get RiffVoiceLowPass(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceLowPass(ByVal voiceHandle As Long, ByVal value As Single)
```

Controls the normalized cutoff amount for the per-voice biquad low-pass filter. `1.0` (default) bypasses the filter and passes the full spectrum. Lower values increasingly attenuate high frequencies with a smoother, higher-quality response than the previous one-pole filter. Valid range: `0.01` to `1.0`.

```vb
RiffVoiceLowPass(v) = 0.05  ' Very muffled
RiffVoiceLowPass(v) = 0.3   ' Warm, dark tone
```

#### RiffVoiceHighPass

```vb
Public Property Get RiffVoiceHighPass(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceHighPass(ByVal voiceHandle As Long, ByVal value As Single)
```

Controls the normalized cutoff amount for the per-voice biquad high-pass filter. `0.0` (default) bypasses the filter. Higher values increasingly attenuate bass and low-mid content, producing a thinner or telephone-like quality. Valid range: `0.0` to `0.99`.

```vb
RiffVoiceHighPass(v) = 0.8  ' Thin, radio-like sound
```

### 3-Band EQ

The 3-Band EQ uses biquad filters with independent per-voice state for bass, mid, and treble processing. Each band gain is a linear multiplier, where `1.0` is flat (no boost or cut).

#### RiffVoiceEqBass

```vb
Public Property Get RiffVoiceEqBass(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceEqBass(ByVal voiceHandle As Long, ByVal value As Single)
```

Gain for the low-frequency shelf. Valid range: `0.0` to `5.0`. Default is `1.0`.

#### RiffVoiceEqMid

```vb
Public Property Get RiffVoiceEqMid(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceEqMid(ByVal voiceHandle As Long, ByVal value As Single)
```

Gain for the mid-frequency band. Valid range: `0.0` to `5.0`. Default is `1.0`.

#### RiffVoiceEqTreble

```vb
Public Property Get RiffVoiceEqTreble(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceEqTreble(ByVal voiceHandle As Long, ByVal value As Single)
```

Gain for the high-frequency shelf. Valid range: `0.0` to `5.0`. Default is `1.0`.

```vb
' Scoop the mids for a V-shaped guitar tone
RiffVoiceEqBass(v)   = 1.5
RiffVoiceEqMid(v)    = 0.4
RiffVoiceEqTreble(v) = 1.8
```

> [!NOTE]
> The EQ is bypassed internally when all three bands are exactly `1.0`, saving CPU.

### Compressor

The compressor is an envelope-follower-based dynamics processor with fast attack and slow release coefficients baked in (`0.01` and `0.001` respectively).

#### RiffVoiceCompressorThreshold

```vb
Public Property Get RiffVoiceCompressorThreshold(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceCompressorThreshold(ByVal voiceHandle As Long, ByVal value As Single)
```

The amplitude level above which gain reduction begins. Valid range: `0.01` to `1.0`. Default is `1.0` (effectively disabled). Lower values engage the compressor earlier.

#### RiffVoiceCompressorRatio

```vb
Public Property Get RiffVoiceCompressorRatio(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceCompressorRatio(ByVal voiceHandle As Long, ByVal value As Single)
```

The ratio of input gain change to output gain change above the threshold. `1.0` (default) applies no compression. `4.0` means for every 4 dB above threshold, only 1 dB passes. Valid range: `1.0` to `20.0`.

```vb
' Gentle compression for voice
RiffVoiceCompressorThreshold(v) = 0.7
RiffVoiceCompressorRatio(v) = 3.0

' Heavy limiting for SFX
RiffVoiceCompressorThreshold(v) = 0.5
RiffVoiceCompressorRatio(v) = 15.0
```

### Stereo Width

#### RiffVoiceStereoWidth

```vb
Public Property Get RiffVoiceStereoWidth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceStereoWidth(ByVal voiceHandle As Long, ByVal value As Single)
```

Adjusts the perceived width of the stereo field using mid/side processing. `1.0` (default) is unmodified. `0.0` collapses to mono. Values above `1.0` exaggerate width. Valid range: `0.0` to `5.0`.

```vb
RiffVoiceStereoWidth(v) = 0.0  ' Mono
RiffVoiceStereoWidth(v) = 2.5  ' Extra-wide stereo
```

### Tremolo

Tremolo modulates the output volume with a sine LFO.

#### RiffVoiceTremoloRate

```vb
Public Property Get RiffVoiceTremoloRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceTremoloRate(ByVal voiceHandle As Long, ByVal value As Single)
```

The frequency of the volume oscillation LFO in Hz. `0.0` (default) disables the effect. Valid range: `0.0` to `20.0`.

#### RiffVoiceTremoloDepth

```vb
Public Property Get RiffVoiceTremoloDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceTremoloDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

The intensity of the volume oscillation. `0.0` (default) is no effect. `1.0` causes the volume to swing from full to near-silence. Valid range: `0.0` to `1.0`.

```vb
RiffVoiceTremoloRate(v)  = 5.0  ' 5 Hz wobble
RiffVoiceTremoloDepth(v) = 0.6  ' Moderate depth
```

### Auto-Pan

Auto-Pan modulates the stereo pan position with a sine LFO, causing the sound to move across the stereo field automatically.

#### RiffVoiceAutoPanRate

```vb
Public Property Get RiffVoiceAutoPanRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceAutoPanRate(ByVal voiceHandle As Long, ByVal value As Single)
```

The speed of the panning LFO in Hz. `0.0` (default) disables the effect. Valid range: `0.0` to `20.0`.

#### RiffVoiceAutoPanDepth

```vb
Public Property Get RiffVoiceAutoPanDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceAutoPanDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

The maximum pan displacement applied by the LFO. `0.0` (default) is no movement. `1.0` sweeps fully left to right. Valid range: `0.0` to `1.0`.

```vb
RiffVoiceAutoPanRate(v)  = 0.5  ' Slow sweep
RiffVoiceAutoPanDepth(v) = 0.8  ' Wide movement
```

> [!NOTE]
> Auto-Pan offsets the voice's static `RiffVoicePan` value. Both properties interact: a voice panned to `0.5` with `AutoPanDepth` of `1.0` will sweep from `-0.5` to `1.0`, not from `-1.0` to `1.0`.

### Ring Modulator

Ring modulation multiplies the audio signal by a sine oscillator, producing metallic, bell-like, or inharmonic timbres.

#### RiffVoiceRingModFreq

```vb
Public Property Get RiffVoiceRingModFreq(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceRingModFreq(ByVal voiceHandle As Long, ByVal value As Single)
```

The frequency in Hz of the ring modulator carrier oscillator. `0.0` (default) disables the effect. Minimum value: `0.0`.

#### RiffVoiceRingModMix

```vb
Public Property Get RiffVoiceRingModMix(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceRingModMix(ByVal voiceHandle As Long, ByVal value As Single)
```

The wet/dry blend of the ring modulator. `0.0` (default) is fully dry. `1.0` is fully ring-modulated. Valid range: `0.0` to `1.0`.

```vb
RiffVoiceRingModFreq(v) = 120.0  ' Carrier at 120 Hz
RiffVoiceRingModMix(v)  = 0.5    ' 50% blend
```

### Chorus

Chorus blends a modulated delayed copy of the signal with the original, creating a thicker, ensemble-like sound.

#### RiffVoiceChorusDepth

```vb
Public Property Get RiffVoiceChorusDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceChorusDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

The wet mix level of the chorus effect. `0.0` (default) disables the effect. `1.0` applies maximum chorus blend. Valid range: `0.0` to `1.0`.

#### RiffVoiceChorusRate

```vb
Public Property Get RiffVoiceChorusRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceChorusRate(ByVal voiceHandle As Long, ByVal value As Single)
```

The LFO rate governing the modulation of the chorus delay tap in Hz. Default is `1.5`. Valid range: `0.1` to `10.0`.

```vb
RiffVoiceChorusDepth(v) = 0.4
RiffVoiceChorusRate(v)  = 1.2
```

### Flanger

Flanger creates a comb-filter sweep effect by mixing the signal with a very short, LFO-modulated copy of itself. Produces the characteristic jet-plane or swirling sound.

#### RiffVoiceFlangerDepth

```vb
Public Property Get RiffVoiceFlangerDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceFlangerDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

The wet blend amount of the flanger. `0.0` (default) disables the effect. Valid range: `0.0` to `1.0`.

#### RiffVoiceFlangerRate

```vb
Public Property Get RiffVoiceFlangerRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceFlangerRate(ByVal voiceHandle As Long, ByVal value As Single)
```

The LFO sweep rate in Hz. Default is `0.5`. Valid range: `0.1` to `10.0`.

#### RiffVoiceFlangerFeedback

```vb
Public Property Get RiffVoiceFlangerFeedback(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceFlangerFeedback(ByVal voiceHandle As Long, ByVal value As Single)
```

The amount of processed signal fed back into the flanger input, increasing resonance and depth. `0.0` (default) is no feedback. Values near `0.95` produce very pronounced, ringing sweeps. Valid range: `0.0` to `0.95`.

```vb
RiffVoiceFlangerDepth(v)    = 0.7
RiffVoiceFlangerRate(v)     = 0.3
RiffVoiceFlangerFeedback(v) = 0.6
```

### Reverb

Reverb simulates spatial reflections using a Freeverb-style comb network with damping and stereo spread. It is still lightweight enough for the VBA real-time callback, but produces smoother spatial tails than the previous simple tap-based reverb.

#### RiffVoiceReverbMix

```vb
Public Property Get RiffVoiceReverbMix(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceReverbMix(ByVal voiceHandle As Long, ByVal value As Single)
```

The wet level of the reverb blend. `0.0` (default) disables the effect. `1.0` is fully wet. Valid range: `0.0` to `1.0`.

#### RiffVoiceReverbTime

```vb
Public Property Get RiffVoiceReverbTime(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceReverbTime(ByVal voiceHandle As Long, ByVal value As Single)
```

Controls the decay rate of the reverb tails by setting the feedback coefficient of the comb filter. `0.0` produces a very short, tight decay. `0.95` produces long, cavernous reverb. Valid range: `0.0` to `0.95`. Default is `0.5`.

```vb
' Small room
RiffVoiceReverbMix(v)  = 0.2
RiffVoiceReverbTime(v) = 0.3

' Cathedral
RiffVoiceReverbMix(v)  = 0.6
RiffVoiceReverbTime(v) = 0.9
```

### Delay

Delay repeats the audio signal after a fixed time interval, with a feedback path for multiple decaying echoes.

#### RiffVoiceDelayTime

```vb
Public Property Get RiffVoiceDelayTime(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDelayTime(ByVal voiceHandle As Long, ByVal value As Single)
```

The interval between echoes in seconds. `0.0` (default) disables the effect. The delay time is converted to a sample offset and aligned to stereo frame boundaries. Valid range: `0.0` to `1.0`.

#### RiffVoiceDelayFeedback

```vb
Public Property Get RiffVoiceDelayFeedback(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDelayFeedback(ByVal voiceHandle As Long, ByVal value As Single)
```

The amount of the delayed signal fed back into the delay line. `0.0` produces a single echo. Values near `0.95` produce many repeating, slowly-decaying echoes. Valid range: `0.0` to `0.95`.

#### RiffVoiceDelayMix

```vb
Public Property Get RiffVoiceDelayMix(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDelayMix(ByVal voiceHandle As Long, ByVal value As Single)
```

The blend of the delay output into the main signal. `0.0` (default) disables the effect. Valid range: `0.0` to `1.0`.

```vb
' Slapback echo at 120 BPM (0.5 beat = 250ms)
RiffVoiceDelayTime(v)     = 0.25
RiffVoiceDelayFeedback(v) = 0.3
RiffVoiceDelayMix(v)      = 0.4
```

## Practical Patterns

### One-Shot Sound Effect

```vb
Sub PlayExplosion()
    Dim v As Long
    v = RiffPlay(explosionBuffer)

    If v >= 0 Then
        RiffVoiceVolume(v) = 1.0
        RiffVoicePan(v) = (Rnd() * 2) - 1  ' Random pan
        RiffVoiceDistortion(v) = 1.5
        RiffVoiceReverbMix(v) = 0.15
    End If
End Sub
```

### Looping Ambient Background

```vb
Sub StartAmbience()
    Dim v As Long
    v = RiffPlay(forestBuffer)

    RiffVoiceLoop(v) = True
    RiffVoiceVolume(v) = 0.4
    RiffVoiceReverbMix(v) = 0.3
    RiffVoiceReverbTime(v) = 0.6
    RiffVoiceChorusDepth(v) = 0.2
    RiffFadeIn v, 4.0
End Sub
```

### Synthesized UI Beep

```vb
Sub PlayUIBeep()
    Dim v As Long
    v = RiffPlayOscillator(0, 880.0)  ' Sine at 880 Hz

    RiffVoiceVolume(v) = 0.3
    RiffVoiceBus(v) = 3  ' Route to UI bus
    RiffFadeOut v, 0.15  ' Quick fade out
End Sub
```

### Per-Voice VU Meter

```vb
Sub UpdateVUMeters()
    Dim pL As Single, pR As Single

    RiffVoiceGetPeak musicVoice, pL, pR
    Sheet1.Shapes("MeterL").Width = pL * 200
    Sheet1.Shapes("MeterR").Width = pR * 200
End Sub
```

### Retro Game Audio

```vb
Sub PlayRetroEffect()
    Dim v As Long
    v = RiffPlayOscillator(1, 220.0)  ' Square wave

    RiffVoiceBitDepth(v) = 8
    RiffVoiceSampleRateReduction(v) = 3
    RiffVoicePitch(v) = 1.0
    RiffVoiceVolume(v) = 0.7
    RiffFadeOut v, 0.3
End Sub
```

### Dynamic Music Bus

```vb
Sub SetMusicIntensity(ByVal intensity As Single)
    ' intensity: 0.0 (calm) to 1.0 (full action)
    RiffBusVolume(0) = intensity        ' Main music track
    RiffBusVolume(1) = 1.0 - intensity  ' Soft ambient layer
End Sub
```

### Export a Loaded Buffer

```vb
Sub ExportLoadedAudio()
    If Not RiffOpen() Then Exit Sub

    Dim buf As Long
    buf = RiffLoad("C:\Sounds\line_reading.mp3")

    If buf >= 0 Then
        If RiffExportBufferWav(buf, "C:\Sounds\line_reading.wav") Then
            Debug.Print "Export complete."
        Else
            Debug.Print "Export failed."
        End If
    End If
End Sub
```

### Render a Test Tone

```vb
Sub RenderTone()
    If Not RiffOpen() Then Exit Sub

    RiffRenderOscillatorWav 0, 440, 3, "C:\Tones\a4_sine.wav"
    RiffRenderOscillatorWav 1, 220, 1, "C:\Tones\a3_square.wav"
End Sub
```

## Operational Caveats

> [!WARNING]
> Always call `RiffClose` before resetting the VBA project or closing the workbook. The native timer thunk points into the VBA runtime. Resetting without cleanup will crash Excel.

> [!WARNING]
> Riff operates entirely on the Windows timer thread calling back into the VBA runtime via a machine-code thunk. The DSP callback fires every 15 ms regardless of what your VBA code is doing. Do not access `rVoices` or `rCtx` directly from user code: all interactions must go through the public API.

> [!CAUTION]
> The voice pool is limited to 32 simultaneous voices. In high-density scenarios (many rapid one-shots), consider stopping or reusing voices explicitly rather than relying solely on natural playback completion.

> [!CAUTION]
> The buffer pool is limited to 64 entries. Unload buffers that are no longer needed with `RiffUnload` to free physical memory and pool slots.

> [!NOTE]
> Audio loading via `RiffLoad` and `RiffLoadFromMemory` is synchronous and may introduce a brief pause for large or heavily compressed files. Decode all assets during an initialization phase rather than during gameplay or interaction.

> [!NOTE]
> WAV export is also synchronous. Large buffers may briefly block the host while PCM data is converted and written to disk.

> [!NOTE]
> The 3-Band EQ processing is bypassed automatically when all three band gains are exactly `1.0`. Similarly, each DSP stage with a depth or mix parameter set to `0.0` is skipped. Setting parameters back to their defaults recovers the CPU cost of those stages.

> [!NOTE]
> The Delay, Chorus, Flanger, and Reverb effects all share the same 192,000-sample ring buffer per voice. Very long delay times combined with deep reverb on the same voice may produce audible interactions between the two effect paths.
