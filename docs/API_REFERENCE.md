# Riff API Reference

**Riff.bas** is a high-performance VBA audio engine for Windows Office hosts. It uses Media Foundation for decoding, WASAPI shared-mode output for playback, and a real-time DSP pipeline written directly in VBA with a small native timer thunk. It supports x86 and x64 VBA, SafeArray-backed dynamic audio buffers, polyphonic voices, expandable audio buses, synthetic oscillators, white/pink/brown noise, WAV export, peak meters, adaptive buffering, musical preset packs, master bus processors, and a full per-voice effects chain.

This reference includes the current stability/performance pass: dynamic buffer/voice/bus pools, manual pre-reservation APIs for large projects, anti-accumulation voice caps, finite procedural one-shots, Hz-based filter helpers, lazy temporal-buffer preparation, faster `RiffPlay`/preset setup, warm-silence underrun protection, the VBE-safe idle timer that prevents IntelliSense from being held in a permanent running state, optional VBE foreground protection, and stop/reset-safe editor cleanup for manual VBE interruptions.

```vb
voice = RiffPlay(bufferHandle, RiffBusSfx, False, 0.9, 0!)
music = RiffPlayOnce(musicBuffer, RiffBusMusic, True, 0.45, 0!)
noise = RiffPlayNoise(RiffWavePinkNoise, RiffBusSfx, 0.08, 0!, 0.06!)
wind = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.05!, 0!)
osc = RiffPlayOscillator(RiffWaveSine, 440!, RiffBusUi, 0.25, 0!, 0.12!)
```

## Table of Contents

- [Core Concepts](#core-concepts)
  - [Mental Model](#mental-model)
  - [Engine Lifecycle](#engine-lifecycle)
  - [Buffer Pool](#buffer-pool)
  - [Voice Pool](#voice-pool)
  - [Dynamic Resource Pools](#dynamic-resource-pools)
  - [Anti-Accumulation and Burst Safety](#anti-accumulation-and-burst-safety)
  - [Audio Buses](#audio-buses)
  - [Adaptive Buffering](#adaptive-buffering)
  - [Timer and Host Behavior](#timer-and-host-behavior)
  - [VBE-Safe Idle Timer](#vbe-safe-idle-timer)
  - [Stop/Reset-Safe Editor Cleanup](#stopreset-safe-editor-cleanup)
  - [VBE Edit Protection for Live Code Changes](#vbe-edit-protection-for-live-code-changes)
  - [Performance Model](#performance-model)
  - [DSP Pipeline](#dsp-pipeline)
  - [Oscillators and Noise](#oscillators-and-noise)
  - [WAV Export](#wav-export)
  - [Compatibility Strategy](#compatibility-strategy)
- [Enums](#enums)
  - [RiffWaveType](#riffwavetype)
  - [RiffBusId](#riffbusid)
  - [RiffEffectPreset](#riffeffectpreset)
  - [RiffMasterPreset](#riffmasterpreset)
  - [RiffErrorCode](#rifferrorcode)
- [Initialization and Teardown](#initialization-and-teardown)
- [Diagnostics](#diagnostics)
  - [Dynamic Pool Reservation API](#dynamic-pool-reservation-api)
  - [Runtime Voice Counters](#runtime-voice-counters)
- [Global Settings](#global-settings)
- [Master Bus Processors](#master-bus-processors)
- [Audio Buses](#audio-buses-api)
- [Asset Loading and Memory](#asset-loading-and-memory)
- [Export and Offline Rendering](#export-and-offline-rendering)
- [Playback](#playback)
- [Voice State and Transport](#voice-state-and-transport)
- [Voice Properties](#voice-properties)
- [Effect Helpers and Presets](#effect-helpers-and-presets)
- [Musical Preset Packs](#musical-preset-packs)
- [DSP Parameters](#dsp-parameters)
  - [Bitcrusher](#bitcrusher)
  - [Ring Modulator](#ring-modulator)
  - [Auto-Pan](#auto-pan)
  - [3-Band EQ](#3-band-eq)
  - [Compressor](#compressor)
  - [Flanger](#flanger)
  - [Distortion](#distortion)
  - [Low-Pass and High-Pass Filters](#low-pass-and-high-pass-filters)
  - [Stereo Width](#stereo-width)
  - [Tremolo](#tremolo)
  - [Chorus](#chorus)
  - [Reverb](#reverb)
  - [Delay](#delay)
- [Practical Recipes](#practical-recipes)
- [Best Practices](#best-practices)
- [Performance Benchmarks](#performance-benchmarks)
- [Troubleshooting](#troubleshooting)
- [Complete Public API Index](#complete-public-api-index)

## Core Concepts

### Mental Model

Riff has four main layers:

1. **Engine**: `RiffOpen`, `RiffClose`, `RiffSuspend`, and `RiffWake` manage Media Foundation, WASAPI, the audio timer, and native memory.
2. **Buffers**: `RiffLoad` and `RiffLoadFromMemory` decode audio into reusable static buffers.
3. **Voices**: `RiffPlay`, `RiffPlayOnce`, `RiffPlayOscillator`, and `RiffPlayNoise` create active sound instances. A voice is the thing you control with volume, pan, pitch, loop, fade, and effects.
4. **Buses**: `RiffBusMusic`, `RiffBusSfx`, `RiffBusUi`, and others group voices for global volume/mute/solo/fade control.

The key distinction is:

```txt
Buffer = loaded audio asset
Voice  = one currently playing instance of a buffer, oscillator, or noise source
Bus    = group routing layer for volume/mute/solo/fade/peak control
```

A single buffer can be played many times at once, producing many voices. Each voice can have different DSP settings.

```vb
Dim coin As Long
Dim v1 As Long
Dim v2 As Long

coin = RiffLoad("C:\Game\Audio\coin.wav")

v1 = RiffPlay(coin, RiffBusSfx, False, 1!, -0.5!)
v2 = RiffPlay(coin, RiffBusSfx, False, 0.7!, 0.5!)
```

### Engine Lifecycle

A typical program opens Riff once, loads assets, plays sounds during interaction, then closes Riff when the workbook or presentation exits.

```vb
Private sndClick As Long
Private sndMusic As Long
Private musicVoice As Long

Public Sub AudioInit()
    If Not RiffOpen() Then
        MsgBox "Riff failed to initialize. Error: " & CStr(RiffLastError)
        Exit Sub
    End If

    sndClick = RiffLoad(ActivePresentation.Path & "\audio\click.wav")
    sndMusic = RiffLoad(ActivePresentation.Path & "\audio\music.mp3")

    RiffBusVolume(RiffBusMusic) = 0.45!
    RiffBusVolume(RiffBusSfx) = 0.9!

    musicVoice = -1
End Sub

Public Sub AudioShutdown()
    RiffClose
End Sub
```

### Buffer Pool

Riff stores decoded audio in a dynamic SafeArray-backed buffer pool. The engine starts with **64 buffer slots** for compatibility and speed, but the pool can grow automatically when more assets are loaded. A buffer handle is still a numeric integer, and loading returns `-1` on failure.

```vb
Dim explosion As Long
explosion = RiffLoad("C:\Game\explosion.wav")

If explosion = -1 Then
    Debug.Print "Load failed:", RiffLastError
End If
```

Loading is synchronous. Decode files during setup, not during frame-by-frame gameplay or UI animation. For very large projects, call `RiffReserveBuffers` during initialization so the buffer pool is prepared before the first big asset batch is loaded.

### Voice Pool

Riff starts with **32 active voice slots** for compatibility and low startup cost, but the voice pool can grow dynamically when more simultaneous voices are needed. A voice handle is still a numeric integer. Playback functions return `-1` only when the engine cannot allocate or grow a usable voice slot.

```vb
Dim v As Long
v = RiffPlay(explosion, RiffBusSfx)

If v <> -1 Then
    RiffVoiceVolume(v) = 0.8!
    RiffVoicePan(v) = 0.2!
End If
```

Each voice has its own source position, loop state, volume, pitch, pan, peak meters, lifetime counters, fade state, peak meters, and DSP state.


### Dynamic Resource Pools

Riff NEXT 1.2.1 removes the old hardcoded practical limits for buffers and voices. Internally, the main pools are SafeArray-backed and grow on demand:

```txt
Buffers -> start at 64, grow when more loaded assets are needed
Voices  -> start at 32, grow when more simultaneous playback slots are needed
Buses   -> start at 16 built-in buses, can be reserved/grown for custom routing
```

The public handles remain simple `Long` values. This means old projects do not need to change how they store buffer handles, voice handles, or bus IDs.

```vb
Dim sndHit As Long
Dim v As Long

sndHit = RiffLoad("C:\Game\Audio\hit.wav")
v = RiffPlay(sndHit, RiffBusSfx)
```

The engine still begins with the original capacities because that keeps small projects fast and avoids allocating more memory than necessary. Large projects can ask Riff to prepare bigger pools up front:

```vb
Public Sub AudioPreloadLargeProject()
    If Not RiffOpen() Then Exit Sub

    RiffReserveBuffers 500
    RiffReserveVoices 128
    RiffReserveBuses 32
End Sub
```

This does not mean infinite memory. It means the old fixed Riff-side limit is gone. Real limits now come from available memory, Office bitness, the host process, and how much decoded audio or DSP state the project keeps alive.

### Anti-Accumulation and Burst Safety

The current stable Riff build includes additional guardrails for games that trigger very short sounds repeatedly. This specifically targets the situation where a gameplay loop fires the same SFX, noise burst, or oscillator many times per second and the engine appears to slow down because active voices or DSP state pile up.

The safety model is:

```txt
Short repeated SFX -> optional cap per buffer and per bus
Noise one-shots    -> finite duration by default
Oscillator beeps   -> optional finite duration
Voice stealing     -> enabled by default and cleaned before reuse
Release/attack     -> micro-ramped to reduce clicks and pops
Preset values      -> clamped/sanitized after preset application
```

Relevant runtime controls:

```vb
RiffVoiceStealingEnabled = True

' 0 disables explicit burst caps. This is the NEXT 1.2.1 default.
RiffMaxVoicesPerBuffer = 0
RiffMaxVoicesPerBus = 0

' Optional game-style caps when you want stricter burst control.
RiffMaxVoicesPerBuffer = 4
RiffMaxVoicesPerBus = 18
```

Relevant diagnostics:

```vb
Debug.Print "Active voices:", RiffActiveVoiceCount()
Debug.Print "SFX voices:", RiffBusVoiceCount(RiffBusSfx)
Debug.Print "Explosion instances:", RiffBufferVoiceCount(sndExplosion, RiffBusSfx)
```

For gameplay code, prefer finite one-shots:

```vb
RiffPlay sndStep, RiffBusSfx, False, 0.35!
RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.15!, 0!, 0.04!
RiffPlayOscillator RiffWaveSquare, 880!, RiffBusUi, 0.18!, 0!, 0.06!
```

Use continuous noise only when you really want ambience or a looped procedural layer:

```vb
windVoice = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.04!)
RiffVoiceSetFilterHz windVoice, 1800!, 80!
' Later:
RiffFadeOut windVoice, 1.5!
```

### Audio Buses

Riff provides **16 built-in named buses** by default:

```txt
Main, Sfx, Music, Voice, Ui, Aux1 ... Aux11
```

The built-in enum covers IDs `0` to `15`. In NEXT 1.2.1, the internal bus pool can be reserved or grown beyond that for advanced custom routing, but the named `RiffBusId` values remain the recommended public routing layer for normal projects.

Buses are used to group voices. For example, all music voices can be routed to `RiffBusMusic`, all effects to `RiffBusSfx`, and all interface sounds to `RiffBusUi`.

```vb
RiffBusVolume(RiffBusMusic) = 0.35!
RiffBusVolume(RiffBusSfx) = 0.9!
RiffBusVolume(RiffBusUi) = 1!
```

The effective output gain is conceptually:

```txt
voice output × voice volume × bus volume × master volume
```

Buses also support mute, solo, fade, peak metering, and optional persistent voice presets that can be applied to all current and future voices on a bus.

### Master Bus Processors

Riff v1.0.9 adds a lightweight master processing stage after all voices and buses are mixed. This stage is useful for final polish and global mix shaping.

The master stage can apply:

- soft clipping
- low-pass and high-pass filtering
- 3-band EQ
- compression/glue
- drive/saturation
- stereo width
- output gain
- named master presets

Use master processors carefully. They affect the entire mix, not just one voice or bus.

```vb
RiffMasterProcessorEnabled = True
RiffMasterApplyPreset RiffMasterFxGlue, 0.7
RiffMasterOutputGain = 0.95!
```

For clean output:

```vb
RiffMasterClearProcessors
```

### Adaptive Buffering

The adaptive-buffer build tries to balance low latency and stability. Fixed low buffers respond quickly but can underrun when PowerPoint, Excel, the VBE, or the PC stalls. Fixed high buffers are stable but make stop/play feel delayed.

Riff adapts dynamically:

```txt
Normal state:      target queue ≈ 60 ms
Recovery state:    target queue ≈ 160 ms
Panic state:       target queue ≈ 200 ms
Stable recovery:   gradually returns toward 60 ms
```

The device buffer is initialized large enough to support the safe queue target, while normal playback still aims for low latency. Use diagnostics to inspect current behavior:

```vb
Debug.Print "Queue target:", RiffAdaptiveQueueMs
Debug.Print "Underruns:", RiffUnderrunCount
Debug.Print "Padding:", RiffLastPaddingFrames
Debug.Print "Available:", RiffLastFramesAvailable
Debug.Print "Written:", RiffLastFramesWritten
```

Use `RiffResetAdaptiveStats` before a test run.

```vb
RiffResetAdaptiveStats
```

### Timer and Host Behavior

Riff uses a Windows timer and a small native thunk to call back into the VBA runtime. This allows real-time-ish playback inside Office without requiring an external DLL. The thunk checks the VBE state to reduce the risk of callbacks continuing after a VBA reset or break.

The adaptive build also supports timer suspension:

- `RiffSuspend` stops the render timer without unloading buffers.
- `RiffWake` restarts the engine/timer path when needed.
- `RiffAutoSuspendTimer` can allow Riff to suspend the timer when idle.

Office is not a real-time audio host. If the UI freezes, Windows is under heavy load, or the VBE is in an unusual state, playback can still be affected. The adaptive queue reduces audible dropouts but cannot fully replace a native audio thread.

### VBE-Safe Idle Timer

The current performance build includes a VBE-safe idle policy for the render timer. Earlier performance builds could keep the render timer alive indefinitely when `RiffAutoSuspendTimer = False`, which kept the VBA project visually stuck in a running state. In the VBE this could make IntelliSense stop working and make the title bar flicker between normal and running states.

The updated behavior is:

```txt
Active audio                    -> render timer stays active
Short silence after gameplay SFX -> timer keeps WASAPI warm briefly
Longer idle period              -> timer stops automatically
Next Play/Wake call             -> timer restarts safely
```

This keeps the low-underrun behavior needed for rapid SFX bursts while avoiding the common Office/VBE problem where a timer callback keeps the editor in `Running` mode after the sound has finished.

Recommended setup for games:

```vb
Public Sub AudioInitForGame()
    If Not RiffOpen() Then Exit Sub

    ' Keeps short SFX responsive, but the VBE-safe idle guard
    ' still stops the timer after the engine is truly idle.
    RiffAutoSuspendTimer = False
End Sub
```

Recommended setup for tools, demos, editors, and normal Office workflows:

```vb
Public Sub AudioInitForOfficeTool()
    If Not RiffOpen() Then Exit Sub
    RiffAutoSuspendTimer = True
End Sub
```

If IntelliSense ever appears stuck after a manual break/reset during development, the current stop-safe build can also use `RiffEditorEmergencyStop` to kill leftover editor timers/callbacks without requiring a full asset unload.

### Stop/Reset-Safe Editor Cleanup

The VBE Stop/Reset button is a special case because it can reset VBA variables while an old Windows timer callback is still pending. In affected builds, this could leave a stale timer calling back into a reset project, which made the VBE title bar flicker and kept IntelliSense disabled until the user paused/stopped the project again.

The stop-safe build protects against this in two layers:

```txt
Normal idle after audio  -> VBE-safe idle policy stops the timer automatically
Manual VBE Stop/Reset   -> stale timer callback self-kills its received timer id
Unknown old timer id    -> callback treats it as orphaned and kills it
Emergency editor cleanup -> RiffEditorEmergencyStop kills/suspends editor-side timer state
```

This means a manual click on the VBE Stop/Reset button should no longer leave Riff in a half-running editor state. For normal application shutdown, still prefer `RiffClose`; for a development-only emergency cleanup, use `RiffEditorEmergencyStop`.

```vb
Public Sub DevAudioEmergencyStop()
    RiffEditorEmergencyStop
End Sub
```


### VBE Edit Protection for Live Code Changes

Editing live VBA code while a native timer/callback is running is one of the riskiest workflows in Office. The VBE may recompile or rearrange project state while Riff is still rendering audio through the timer path. In that exact situation, a host crash can also leave the workbook/presentation in a corrupted state.

NEXT 1.2.1 adds a manual editor-safe workflow:

```vb
RiffPrepareForVbeEdit
' Change code in the VBE here.
RiffResumeAfterVbeEdit
```

`RiffPrepareForVbeEdit` stops the render timer but keeps the engine, loaded buffers, and WASAPI resources alive. `RiffResumeAfterVbeEdit` restarts the timer after the edit window is safe again.

For convenience, you can wrap this into development macros:

```vb
Public Sub EntrarModoEdicaoSeguro()
    RiffPrepareForVbeEdit
End Sub

Public Sub SairModoEdicaoSeguro()
    If Not RiffResumeAfterVbeEdit() Then
        Debug.Print "Riff resume failed:", RiffLastError
    End If
End Sub
```

There is also an optional automatic foreground guard:

```vb
RiffEditorSafeMode = True
```

This mode suspends the render timer when the VBE receives focus. It is disabled by default because many developers test playback directly from the VBE, and an aggressive automatic guard can make it look like Riff opened but produced no sound.

### Performance Model

The performance build optimizes the common gameplay path: loading happens once, `RiffPlay` is called many times, and most short sounds are dry or only lightly processed. The main improvements are internal and do not change the public API.

Key performance changes:

- `RiffPlay` no longer clears the large per-voice temporal ring buffer for ordinary dry SFX.
- Delay, reverb, chorus, and flanger ring buffers are prepared lazily only when a temporal effect is actually used.
- `RiffVoiceApplyPreset` no longer pays the full temporal-buffer reset cost for dry/EQ/compressor/filter/radio-style presets.
- Voice allocation performs fewer repeated scans when choosing a free or stealable voice.
- Generated noise/oscillator sources use finite lifetime counters when a duration is provided, so they do not remain active accidentally.
- Preset values are sanitized after application to reduce invalid feedback/filter/width/compressor states.
- The render path skips neutral DSP stages where possible.

Expected practical impact from the internal benchmark suite:

```txt
Dry RiffPlay path:       around 11-13 us/call in the tested Office/VBA host
Noise one-shot:          around 14-15 us/call
Oscillator one-shot:     around 13-15 us/call
Preset/DSP setup path:   around 20 us/call after lazy temporal-buffer optimization
Game-loop benchmark:     0 failed handles, 0 final active voices, stable memory
```

These numbers are benchmark-dependent, but the important trend is that the regular SFX path is now much cheaper than earlier builds that cleared large DSP buffers during every play or preset setup.

### DSP Pipeline

Each active voice goes through a fixed per-sample DSP chain. The conceptual order is:

1. Source read: buffer, oscillator, or noise generator.
2. Pitch/source stepping.
3. Bitcrusher sample-rate hold.
4. Bit-depth reduction.
5. Distortion/saturation.
6. Low-pass and high-pass filters.
7. 3-band EQ.
8. Ring modulation.
9. Tremolo.
10. Stereo width.
11. Flanger.
12. Chorus.
13. Delay.
14. Reverb.
15. Compressor.
16. Auto-pan.
17. Voice gain, bus gain, master gain.
18. Fade gain.
19. Peak metering and mix accumulation.
20. Master bus processing: optional soft clip, filters, EQ, compression, drive, stereo width, and output gain.

Most DSP stages are skipped when their parameters are at neutral defaults. For example, delay is skipped when `RiffVoiceDelayMix = 0`, and EQ is skipped when all EQ bands are `1.0`.

The performance build also avoids clearing and preparing heavy temporal-effect memory unless it is needed. Delay, reverb, chorus, and flanger use a large per-voice ring buffer. Earlier builds could clear this buffer during normal voice reset or preset application, which made fast repeated SFX noticeably heavier. The current build marks temporal state as stale and prepares it lazily when a temporal effect first needs it.

This means a dry sound, a filtered UI click, a bitcrushed beep, or a radio-style preset does not pay the same setup cost as a voice with delay/reverb/chorus/flanger.

### Oscillators and Noise

Riff can synthesize audio without a file. Oscillators use `frequencyHz` to define pitch: low Hz is bass/low pitch, high Hz is bright/high pitch.

```vb
Dim beep As Long
beep = RiffPlayOscillator(RiffWaveSine, 880!, RiffBusUi, 0.25!, 0!, 0.08!)
```

`durationSec` is optional. When omitted or `0`, oscillators are continuous and must be stopped manually. For UI/game SFX, pass a short duration so the voice cleans itself automatically.

```vb
' Continuous oscillator; stop manually.
Dim hum As Long
hum = RiffPlayOscillator(RiffWaveSine, 55!, RiffBusSfx, 0.15!)
RiffStop hum

' Finite beep; auto-releases.
RiffPlayOscillator RiffWaveSquare, 1320!, RiffBusUi, 0.18!, 0!, 0.05!
```

Noise is also supported. In the stable gameplay build, `RiffPlayNoise` is a short one-shot by default. This prevents procedural noise from accidentally accumulating forever when used as a fast SFX.

```vb
' Short noise hit; auto-stops after the default noise duration.
RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.18!

' Explicit short duration.
RiffPlayNoise RiffWavePinkNoise, RiffBusSfx, 0.08!, 0!, 0.12!
```

For ambience, use the explicit loop helpers:

```vb
Dim rain As Long
rain = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.05!)
RiffVoiceSetFilterHz rain, 2600!, 120!

' Later:
RiffFadeOut rain, 2!
```

Supported source types:

- `RiffWaveSine`
- `RiffWaveSquare`
- `RiffWaveSawtooth`
- `RiffWaveNoise` / `RiffWaveWhiteNoise`
- `RiffWavePinkNoise`
- `RiffWaveBrownNoise`

### WAV Export

Riff can export loaded buffers and generated oscillators to 16-bit stereo PCM WAV.

```vb
RiffExportBufferWav sndMusic, "C:\Temp\music_export.wav"
RiffRenderOscillatorWav RiffWaveSawtooth, 110!, 1.5!, "C:\Temp\saw.wav"
RiffRenderOscillatorWav RiffWavePinkNoise, 0!, 3!, "C:\Temp\pink_noise.wav"
```

### Compatibility Strategy

The modern API is unified around optional bus, loop, volume, and pan arguments:

```vb
RiffPlay buffer, bus, looped, volume, pan
RiffPlayOnce buffer, bus, looped, volume, pan
RiffPlayOscillator wave, frequency, bus, volume, pan
RiffPlayNoise noiseType, bus, volume, pan
```

Older helper names remain available:

```vb
RiffPlayBus
RiffPlayBusOnce
RiffPlayOscillatorBus
RiffPlayNoiseBus
```

They are compatibility wrappers and should generally not be used in new examples unless you are preserving old code style.

## Enums

### RiffWaveType

```vb
Public Enum RiffWaveType
    RiffWaveSine = 0
    RiffWaveSquare = 1
    RiffWaveSawtooth = 2
    RiffWaveNoise = 3
    RiffWaveWhiteNoise = 3
    RiffWavePinkNoise = 4
    RiffWaveBrownNoise = 5
End Enum
```

| Value | Description |
|:---|:---|
| `RiffWaveSine` | Pure sine tone. Best for clean beeps, test tones, bass fundamentals. |
| `RiffWaveSquare` | Band-limited square wave. Good for retro UI sounds, chiptune, alarms. |
| `RiffWaveSawtooth` | Band-limited saw wave. Bright, buzzy, useful for synth leads and sweeps. |
| `RiffWaveNoise` | Compatibility alias for white noise. |
| `RiffWaveWhiteNoise` | Equal random energy across frequencies. Static, glitches, impacts. |
| `RiffWavePinkNoise` | More natural noise with reduced high-frequency harshness. Rain, wind, fire, ambience. |
| `RiffWaveBrownNoise` | Dark, low-heavy noise. Thunder rumble, engines, earthquakes, distant impacts. |

### RiffBusId

```vb
Public Enum RiffBusId
    RiffBusMain = 0
    RiffBusSfx = 1
    RiffBusMusic = 2
    RiffBusVoice = 3
    RiffBusUi = 4
    RiffBusAux1 = 5
    RiffBusAux2 = 6
    RiffBusAux3 = 7
    RiffBusAux4 = 8
    RiffBusAux5 = 9
    RiffBusAux6 = 10
    RiffBusAux7 = 11
    RiffBusAux8 = 12
    RiffBusAux9 = 13
    RiffBusAux10 = 14
    RiffBusAux11 = 15
End Enum
```

Recommended routing:

| Bus | Typical Use |
|:---|:---|
| `RiffBusMain` | Default/general output. |
| `RiffBusSfx` | Gameplay sound effects. |
| `RiffBusMusic` | Background music, ambience layers. |
| `RiffBusVoice` | Dialogue, narration, character voices. |
| `RiffBusUi` | Menu sounds, buttons, confirmation/cancel sounds. |
| `RiffBusAux1` ... `RiffBusAux11` | Custom groups: ambience, enemies, weather, cinematic, debug, etc. |

### RiffEffectPreset

```vb
Public Enum RiffEffectPreset
    RiffFxDry = 0
    RiffFxSmallRoom = 1
    RiffFxHall = 2
    RiffFxCathedral = 3
    RiffFxSlapback = 4
    RiffFxEcho = 5
    RiffFxChorus = 6
    RiffFxFlanger = 7
    RiffFxLoFi = 8
    RiffFxRadio = 9
    RiffFxUnderwater = 10
    RiffFxWide = 11
    RiffFxRobot = 12
    RiffFxAmbient = 13
    RiffFxWarmTape = 14
    RiffFxVHS = 15
    RiffFxDreamPad = 16
    RiffFxDarkCave = 17
    RiffFxTinySpeaker = 18
    RiffFxMegaphone = 19
    RiffFxGameBoy = 20
    RiffFxHorrorDrone = 21
    RiffFxWind = 22
    RiffFxRain = 23
    RiffFxCinematicBoom = 24
    RiffFxSoftFocus = 25
End Enum
```

Presets are convenience recipes applied to a voice. Use `RiffVoiceApplyPreset voice, preset, amount`. `amount` normally ranges from `0.0` to `1.0`. `0.0` should produce a dry/neutral result, while `1.0` applies the preset fully.

| Preset | Description |
|:---|:---|
| `RiffFxDry` | Clears voice effects back to neutral. |
| `RiffFxSmallRoom` | Short room reverb. |
| `RiffFxHall` | Medium hall reverb. |
| `RiffFxCathedral` | Long, spacious reverb. |
| `RiffFxSlapback` | Short echo useful for retro/voice effects. |
| `RiffFxEcho` | Longer delay with feedback. |
| `RiffFxChorus` | Gentle widening/thickening modulation. |
| `RiffFxFlanger` | Jet-like comb sweep. |
| `RiffFxLoFi` | Musical degraded/tape/old-sampler character. |
| `RiffFxRadio` | Band-limited voice/radio/phone tone. |
| `RiffFxUnderwater` | Muffled low-pass-heavy sound. |
| `RiffFxWide` | Stereo width enhancement. |
| `RiffFxRobot` | Ring-modulated robotic timbre. |
| `RiffFxAmbient` | Spacious reverb/chorus ambience. |
| `RiffFxWarmTape` | Warm tape-like coloration with gentle filtering and compression. |
| `RiffFxVHS` | Warbly degraded VHS-style tone. |
| `RiffFxDreamPad` | Soft wide chorus/reverb preset for pads and ambience. |
| `RiffFxDarkCave` | Dark, large, cave-like space. |
| `RiffFxTinySpeaker` | Small speaker coloration with narrow bandwidth. |
| `RiffFxMegaphone` | Mid-forward projected voice effect. |
| `RiffFxGameBoy` | Retro crunchy bit-reduced game tone. |
| `RiffFxHorrorDrone` | Dark modulated unsettling texture. |
| `RiffFxWind` | Airy filtered noise/wind coloration. |
| `RiffFxRain` | Soft filtered rain/noise ambience. |
| `RiffFxCinematicBoom` | Big low-heavy impact treatment. |
| `RiffFxSoftFocus` | Smooth, softened, gentle wide tone. |

### RiffMasterPreset

```vb
Public Enum RiffMasterPreset
    RiffMasterFxClean = 0
    RiffMasterFxGlue = 1
    RiffMasterFxWarm = 2
    RiffMasterFxBright = 3
    RiffMasterFxDark = 4
    RiffMasterFxRadio = 5
    RiffMasterFxCinematic = 6
    RiffMasterFxNight = 7
    RiffMasterFxSoftLimiter = 8
End Enum
```

Master presets apply to the final mixed output. They are used with `RiffMasterApplyPreset`.

| Preset | Description |
|:---|:---|
| `RiffMasterFxClean` | Neutral master stage. |
| `RiffMasterFxGlue` | Gentle compression and soft limiting for a cohesive mix. |
| `RiffMasterFxWarm` | Warm tonal shaping with slight saturation. |
| `RiffMasterFxBright` | Brighter master tone with controlled low end. |
| `RiffMasterFxDark` | Darker master tone for night/cave/low-energy scenes. |
| `RiffMasterFxRadio` | Global radio/band-limited output. |
| `RiffMasterFxCinematic` | Wider, fuller, slightly compressed cinematic shaping. |
| `RiffMasterFxNight` | Lower, softer, less bright nighttime mix. |
| `RiffMasterFxSoftLimiter` | Safety preset focused on limiting and clipping control. |

### RiffErrorCode

```vb
Public Enum RiffErrorCode
    RiffErrorNone = 0
    RiffErrorNotInitialized = 1
    RiffErrorNoFreeBuffer = 2
    RiffErrorNoFreeVoice = 3
    RiffErrorInvalidBuffer = 4
    RiffErrorInvalidVoice = 5
    RiffErrorInvalidBus = 6
    RiffErrorInvalidArgument = 7
    RiffErrorFileNotFound = 8
    RiffErrorComFailure = 9
    RiffErrorMemoryAllocation = 10
    RiffErrorDecodeFailed = 11
    RiffErrorUnsupportedFormat = 12
End Enum
```

| Error | Meaning |
|:---|:---|
| `RiffErrorNone` | Last operation completed without a Riff-level error. |
| `RiffErrorNotInitialized` | The engine is not open. Call `RiffOpen`. |
| `RiffErrorNoFreeBuffer` | Riff could not find or grow a free buffer slot. This usually means memory allocation failed or the engine state is invalid. |
| `RiffErrorNoFreeVoice` | Riff could not find, steal, or grow a usable voice slot. Check memory pressure, caps, and active voice counts. |
| `RiffErrorInvalidBuffer` | Buffer handle is outside range or inactive. |
| `RiffErrorInvalidVoice` | Voice handle is outside range or invalid for the operation. |
| `RiffErrorInvalidBus` | Bus id is outside the currently allocated bus capacity. Built-in named buses are still `0` to `15`. |
| `RiffErrorInvalidArgument` | Argument is empty, invalid, or outside accepted range. |
| `RiffErrorFileNotFound` | Source path does not exist. |
| `RiffErrorComFailure` | Windows/COM/WASAPI/Media Foundation call failed. |
| `RiffErrorMemoryAllocation` | Native memory allocation failed. |
| `RiffErrorDecodeFailed` | Media Foundation could not produce usable PCM. |
| `RiffErrorUnsupportedFormat` | Device/source format is unsupported by the renderer. |

## Initialization and Teardown

### RiffOpen

```vb
Public Function RiffOpen() As Boolean
```

Initializes Media Foundation, WASAPI shared-mode rendering, native memory, timer resolution, internal voice/buffer/bus state, and the adaptive render loop.

Returns `True` on success. Returns `False` on failure and sets `RiffLastError`.

```vb
Public Sub StartAudio()
    If Not RiffOpen() Then
        MsgBox "Audio failed. RiffLastError = " & CStr(RiffLastError), vbCritical
        Exit Sub
    End If
End Sub
```

Calling `RiffOpen` when already initialized is safe and returns `True`.

### RiffClose

```vb
Public Sub RiffClose()
```

Fully shuts down Riff. It stops voices, kills/suspends the render timer, releases loaded buffers, releases COM interfaces, shuts down Media Foundation, and restores the timer resolution.

Call it when the workbook/presentation closes, when the slide show ends, or before resetting the VBA project.

PowerPoint example:

```vb
Public Sub OnSlideShowTerminate(ByVal Pres As Presentation)
    RiffClose
End Sub
```

Excel example:

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RiffClose
End Sub
```

### RiffEditorEmergencyStop

```vb
Public Sub RiffEditorEmergencyStop()
```

Development/emergency cleanup helper for the VBA editor. It is intended for the rare case where a manual VBE Stop/Reset or an older test build leaves a stale render timer/callback alive after the project has been interrupted.

Unlike `RiffClose`, this helper is primarily editor-focused: it force-stops the render timer path and clears the state that can keep the VBE visually stuck in `Running` mode. Use it as a recovery tool while developing, not as the normal shutdown path for a finished workbook or presentation.

```vb
Public Sub AudioEmergencyResetForEditor()
    RiffEditorEmergencyStop
End Sub
```

Recommended normal shutdown still remains:

```vb
Public Sub AudioShutdown()
    RiffClose
End Sub
```

Use `RiffEditorEmergencyStop` when the editor itself is the problem: IntelliSense remains disabled, the title bar keeps flickering, or a previous manual Stop/Reset left Riff in a stale timer state.

### RiffIsInitialized

```vb
Public Property Get RiffIsInitialized() As Boolean
```

Returns `True` when Riff is currently initialized.

```vb
If Not RiffIsInitialized Then
    If Not RiffOpen() Then Exit Sub
End If
```

### RiffSuspend

```vb
Public Sub RiffSuspend()
```

Suspends the render timer without unloading decoded buffers. This is useful when you want to temporarily stop the audio loop but keep assets in memory.

```vb
Public Sub PauseAudioSystem()
    RiffStopAll
    RiffSuspend
End Sub
```


### RiffPrepareForVbeEdit

```vb
Public Sub RiffPrepareForVbeEdit()
```

Stops the render timer before editing live VBA code while Riff is initialized. It does not unload buffers, release WASAPI resources, or close the engine.

Use it when you need to change code in the VBE while audio may still be active or while the Riff timer may still be alive.

```vb
Public Sub BeforeEditingAudioCode()
    RiffPrepareForVbeEdit
End Sub
```

This is a development-safety helper. For final application shutdown, still use `RiffClose`.

### RiffResumeAfterVbeEdit

```vb
Public Function RiffResumeAfterVbeEdit() As Boolean
```

Restarts the render timer after a manual VBE edit pause. Returns `True` when the timer is running or was started successfully. Returns `False` and sets `RiffLastError` if the engine is not initialized or the timer path cannot be restored.

```vb
Public Sub AfterEditingAudioCode()
    If Not RiffResumeAfterVbeEdit() Then
        Debug.Print "Riff resume failed:", RiffLastError
    End If
End Sub
```

### RiffWake

```vb
Public Function RiffWake() As Boolean
```

Wakes/restarts the render path after `RiffSuspend`. If the engine is not initialized, it attempts to initialize it.

```vb
If RiffWake() Then
    RiffPlay sndClick, RiffBusUi
End If
```


### RiffEditorSafeMode

```vb
Public Property Get RiffEditorSafeMode() As Boolean
Public Property Let RiffEditorSafeMode(ByVal value As Boolean)
```

Enables or disables the optional VBE foreground guard. When enabled, Riff may suspend the render timer when the VBE is the foreground window. This reduces the chance of editing code while a native callback is still touching VBA state.

```vb
RiffEditorSafeMode = True
```

This property is disabled by default in NEXT 1.2.1 because normal tests launched from the VBE should still produce sound. Prefer the explicit `RiffPrepareForVbeEdit` / `RiffResumeAfterVbeEdit` workflow while validating a project.

### RiffEditorTimerSuspended

```vb
Public Property Get RiffEditorTimerSuspended() As Boolean
```

Returns `True` when the editor guard or manual edit workflow has suspended the render timer for VBE safety.

```vb
If RiffEditorTimerSuspended Then
    Debug.Print "Riff timer is paused for editor safety."
End If
```

### RiffAutoSuspendTimer

```vb
Public Property Get RiffAutoSuspendTimer() As Boolean
Public Property Let RiffAutoSuspendTimer(ByVal value As Boolean)
```

Controls whether Riff may automatically stop/suspend its timer after a period of silence. This reduces the chance of a timer keeping the VBE in a running state after one-shot sounds finish.

```vb
RiffAutoSuspendTimer = True
```

When `True`, Riff aggressively suspends the timer after idle periods, which is the safest mode for normal Office tools, demos, and editing workflows.

When `False`, Riff keeps the audio path warm for rapid repeated playback, which is useful for games with lots of short SFX. In the VBE-safe performance build, this no longer means the timer stays alive forever: after a short warm-idle period with no active voices, Riff still stops the timer automatically so the VBE can return to a normal editable state and IntelliSense can recover. In the stop-safe build, stale timer callbacks caused by the VBE Stop/Reset button are also self-terminated so the editor is not left flickering in a half-running state.

```vb
' Game-style usage: responsive bursts, still VBE-safe after idle.
RiffAutoSuspendTimer = False

' Office/editor-style usage: release the VBE as soon as practical.
RiffAutoSuspendTimer = True
```

## Diagnostics

### RiffLastError

```vb
Public Property Get RiffLastError() As RiffErrorCode
```

Returns the last Riff-level error code.

```vb
Dim b As Long
b = RiffLoad("C:\missing.wav")

If b = -1 Then
    Debug.Print "Load failed:", RiffLastError
End If
```

### RiffMaxVoices

```vb
Public Property Get RiffMaxVoices() As Long
```

Returns the current allocated voice capacity. In NEXT 1.2.1 this starts at `32`, but it may grow automatically or through `RiffReserveVoices`.

```vb
Dim i As Long
For i = 0 To RiffMaxVoices - 1
    If RiffVoiceIsPlaying(i) Then Debug.Print "Voice playing:", i
Next
```

### RiffMaxBuffers

```vb
Public Property Get RiffMaxBuffers() As Long
```

Returns the current allocated buffer capacity. In NEXT 1.2.1 this starts at `64`, but it may grow automatically or through `RiffReserveBuffers`.

### RiffMaxBuses

```vb
Public Property Get RiffMaxBuses() As Long
```

Returns the current allocated bus capacity. It starts at the built-in `16` named buses and can grow through `RiffReserveBuses` for advanced routing.


### Dynamic Pool Reservation API

These functions are optional. Riff grows pools automatically when possible, but pre-reserving capacity during startup is better for large games, presentations, or audio-heavy tools because it avoids growth work during gameplay or animation.

#### RiffReserveBuffers

```vb
Public Function RiffReserveBuffers(ByVal capacity As Long) As Boolean
```

Ensures the buffer pool can hold at least `capacity` decoded assets. Returns `True` when the requested capacity is available.

```vb
If Not RiffReserveBuffers(500) Then
    Debug.Print "Could not reserve buffers:", RiffLastError
End If
```

Use this before a large `RiffLoad` batch.

#### RiffReserveVoices

```vb
Public Function RiffReserveVoices(ByVal capacity As Long) As Boolean
```

Ensures the voice pool can hold at least `capacity` simultaneous voice slots. This is useful for projects with heavy SFX bursts, layered ambience, dialogue, and music playing at the same time.

```vb
If Not RiffReserveVoices(128) Then
    Debug.Print "Could not reserve voices:", RiffLastError
End If
```

Because every voice carries its own state and effect memory, do not reserve thousands of voices without a real need. Reserve based on the maximum simultaneous playback your project actually expects.

#### RiffReserveBuses

```vb
Public Function RiffReserveBuses(ByVal capacity As Long) As Boolean
```

Ensures the bus pool can hold at least `capacity` bus slots. The first 16 buses are the built-in named buses from `RiffBusId`. Extra bus IDs are for advanced users who want custom routing beyond the default enum.

```vb
If Not RiffReserveBuses(32) Then
    Debug.Print "Could not reserve buses:", RiffLastError
End If
```

Normal projects should keep using `RiffBusMusic`, `RiffBusSfx`, `RiffBusUi`, and the built-in aux buses unless they specifically need more routing groups.

### Runtime Voice Counters

The stable gameplay build exposes lightweight counters that are useful for finding accidental accumulation during development.

#### RiffActiveVoiceCount

```vb
Public Function RiffActiveVoiceCount() As Long
```

Returns the number of currently active voices across all buses and source types. This should return to `0` after finite one-shots finish and after cleanup.

```vb
Debug.Print "Active voices:", RiffActiveVoiceCount()
```

#### RiffBusVoiceCount

```vb
Public Function RiffBusVoiceCount(ByVal busID As RiffBusId) As Long
```

Returns the number of active voices currently routed to a bus.

```vb
Debug.Print "SFX active:", RiffBusVoiceCount(RiffBusSfx)
Debug.Print "Music active:", RiffBusVoiceCount(RiffBusMusic)
```

#### RiffBufferVoiceCount

```vb
Public Function RiffBufferVoiceCount(ByVal bufferHandle As Long, Optional ByVal busID As RiffBusId = RiffBusMain) As Long
```

Returns how many active voices are currently playing a specific decoded buffer. Use this to confirm that fast repeated effects are not multiplying beyond the configured cap.

```vb
Debug.Print "Step voices:", RiffBufferVoiceCount(sndStep, RiffBusSfx)
```

### Voice Stealing and Burst Caps

#### RiffVoiceStealingEnabled

```vb
Public Property Get RiffVoiceStealingEnabled() As Boolean
Public Property Let RiffVoiceStealingEnabled(ByVal value As Boolean)
```

Controls whether Riff may reuse an existing voice when a burst would otherwise exceed the active voice pool or per-source safety caps. It is recommended to keep this enabled for games.

```vb
RiffVoiceStealingEnabled = True
```

#### RiffMaxVoicesPerBuffer

```vb
Public Property Get RiffMaxVoicesPerBuffer() As Long
Public Property Let RiffMaxVoicesPerBuffer(ByVal value As Long)
```

Limits how many instances of the same loaded buffer can be active at once. This prevents a rapid-fire sound such as footsteps, bullets, UI clicks, or collision hits from piling up into a heavy mixer load. A value of `0` disables this explicit cap, which is the NEXT 1.2.1 default.

```vb
' Default/unlimited-by-cap behavior.
RiffMaxVoicesPerBuffer = 0

' Optional stricter gameplay cap.
RiffMaxVoicesPerBuffer = 4
```

#### RiffMaxVoicesPerBus

```vb
Public Property Get RiffMaxVoicesPerBus() As Long
Public Property Let RiffMaxVoicesPerBus(ByVal value As Long)
```

Limits the number of active voices routed to one bus. This helps protect the SFX bus from large gameplay bursts while still allowing music and UI buses to remain responsive. A value of `0` disables this explicit cap, which is the NEXT 1.2.1 default.

```vb
' Default/unlimited-by-cap behavior.
RiffMaxVoicesPerBus = 0

' Optional stricter gameplay cap.
RiffMaxVoicesPerBus = 18
```

### RiffAdaptiveQueueMs

```vb
Public Property Get RiffAdaptiveQueueMs() As Long
```

Returns the current adaptive queue target in milliseconds. Expected values are commonly around `60`, `160`, or `200`, with gradual recovery between states.

```vb
Debug.Print "Adaptive queue target:", RiffAdaptiveQueueMs & "ms"
```

### RiffUnderrunCount

```vb
Public Property Get RiffUnderrunCount() As Long
```

Returns the number of detected underrun-risk events. A rising value indicates that the host or system is failing to feed WASAPI quickly enough.

```vb
If RiffUnderrunCount > 0 Then
    Debug.Print "Audio underruns detected:", RiffUnderrunCount
End If
```

### RiffLastPaddingFrames

```vb
Public Property Get RiffLastPaddingFrames() As Long
```

Returns the last WASAPI padding value in frames. Padding is the amount of queued audio already waiting in the output buffer.

### RiffLastFramesAvailable

```vb
Public Property Get RiffLastFramesAvailable() As Long
```

Returns the number of frames that were available to write on the last render tick.

### RiffLastFramesWritten

```vb
Public Property Get RiffLastFramesWritten() As Long
```

Returns how many frames Riff wrote on the last render tick.

### RiffResetAdaptiveStats

```vb
Public Sub RiffResetAdaptiveStats()
```

Clears adaptive-buffer diagnostics so you can run a clean test.

```vb
Public Sub TestAudioStability()
    RiffResetAdaptiveStats
    RiffPlay sndMusic, RiffBusMusic, True, 0.4!
End Sub
```

### Diagnostic Snapshot Example

```vb
Public Sub PrintRiffDiagnostics()
    Debug.Print "Initialized:", RiffIsInitialized
    Debug.Print "Auto suspend:", RiffAutoSuspendTimer
    Debug.Print "Editor safe mode:", RiffEditorSafeMode
    Debug.Print "Editor timer suspended:", RiffEditorTimerSuspended
    Debug.Print "Buffer capacity:", RiffMaxBuffers
    Debug.Print "Voice capacity:", RiffMaxVoices
    Debug.Print "Bus capacity:", RiffMaxBuses
    Debug.Print "Adaptive queue ms:", RiffAdaptiveQueueMs
    Debug.Print "Underruns:", RiffUnderrunCount
    Debug.Print "Last padding frames:", RiffLastPaddingFrames
    Debug.Print "Last available frames:", RiffLastFramesAvailable
    Debug.Print "Last written frames:", RiffLastFramesWritten
End Sub
```

## Global Settings

### RiffMasterVolume

```vb
Public Property Get RiffMasterVolume() As Single
Public Property Let RiffMasterVolume(ByVal value As Single)
```

Global volume applied after voice and bus gain. Values are clamped internally.

```vb
RiffMasterVolume = 0.75!
```

Use this for the player's master audio setting.

```vb
Public Sub SetMasterFromSlider(ByVal sliderValue As Single)
    RiffMasterVolume = sliderValue / 100!
End Sub
```

### RiffMasterGetPeak

```vb
Public Sub RiffMasterGetPeak(ByRef peakLeft As Single, ByRef peakRight As Single)
```

Returns decaying master peak meters for left and right channels.

```vb
Dim l As Single, r As Single
RiffMasterGetPeak l, r
Debug.Print "Master peak:", Format(l, "0.00"), Format(r, "0.00")
```

## Master Bus Processors

Master processors run after all voices and buses have been mixed. They are designed for final polish, safety limiting, and global scene-level coloration. Unlike voice effects, master processors affect the entire output.

Use them sparingly. A master preset can make a whole project sound more cohesive, but aggressive settings can also exaggerate clipping or reduce clarity.

### RiffSoftClipEnabled

```vb
Public Property Get RiffSoftClipEnabled() As Boolean
Public Property Let RiffSoftClipEnabled(ByVal value As Boolean)
```

Enables or disables the master soft clipper. This is a safety stage intended to reduce harsh digital clipping when many voices overlap.

```vb
RiffSoftClipEnabled = True
```

### RiffMasterProcessorEnabled

```vb
Public Property Get RiffMasterProcessorEnabled() As Boolean
Public Property Let RiffMasterProcessorEnabled(ByVal value As Boolean)
```

Enables or disables the optional master processor chain.

```vb
RiffMasterProcessorEnabled = True
```

When disabled, the mix still uses normal master volume and core output handling, but optional master EQ/compression/drive/width processing is bypassed.

### RiffMasterLowPass

```vb
Public Property Get RiffMasterLowPass() As Single
Public Property Let RiffMasterLowPass(ByVal value As Single)
```

Controls global low-pass filtering. `1.0` is open/neutral. Lower values darken the entire mix.

```vb
RiffMasterLowPass = 0.85!
```

### RiffMasterHighPass

```vb
Public Property Get RiffMasterHighPass() As Single
Public Property Let RiffMasterHighPass(ByVal value As Single)
```

Controls global high-pass filtering. `0.0` is neutral. Higher values remove more low end.

```vb
RiffMasterHighPass = 0.05!
```

### RiffMasterEqBass / RiffMasterEqMid / RiffMasterEqTreble

```vb
Public Property Get RiffMasterEqBass() As Single
Public Property Let RiffMasterEqBass(ByVal value As Single)

Public Property Get RiffMasterEqMid() As Single
Public Property Let RiffMasterEqMid(ByVal value As Single)

Public Property Get RiffMasterEqTreble() As Single
Public Property Let RiffMasterEqTreble(ByVal value As Single)
```

Global 3-band tonal shaping. `1.0` is neutral.

```vb
RiffMasterEqBass = 1.1!
RiffMasterEqMid = 1!
RiffMasterEqTreble = 0.95!
```

### RiffMasterCompressorThreshold / RiffMasterCompressorRatio

```vb
Public Property Get RiffMasterCompressorThreshold() As Single
Public Property Let RiffMasterCompressorThreshold(ByVal value As Single)

Public Property Get RiffMasterCompressorRatio() As Single
Public Property Let RiffMasterCompressorRatio(ByVal value As Single)
```

Global compression controls. Use this for mix glue, dialogue/music balancing, and soft limiting behavior.

```vb
RiffMasterCompressorThreshold = 0.72!
RiffMasterCompressorRatio = 2.5!
```

### RiffMasterDrive

```vb
Public Property Get RiffMasterDrive() As Single
Public Property Let RiffMasterDrive(ByVal value As Single)
```

Applies global drive/saturation before final output gain.

```vb
RiffMasterDrive = 1.08!
```

Use subtle values for warmth. Heavy values can distort the full mix.

### RiffMasterStereoWidth

```vb
Public Property Get RiffMasterStereoWidth() As Single
Public Property Let RiffMasterStereoWidth(ByVal value As Single)
```

Controls global stereo width. `1.0` is neutral, `0.0` is mono, values above `1.0` widen the mix.

```vb
RiffMasterStereoWidth = 1.15!
```

### RiffMasterOutputGain

```vb
Public Property Get RiffMasterOutputGain() As Single
Public Property Let RiffMasterOutputGain(ByVal value As Single)
```

Final output gain after master processing.

```vb
RiffMasterOutputGain = 0.95!
```

### RiffMasterClearProcessors

```vb
Public Sub RiffMasterClearProcessors()
```

Resets master processing to a clean neutral state.

```vb
RiffMasterClearProcessors
```

### RiffMasterApplyPreset

```vb
Public Sub RiffMasterApplyPreset(ByVal preset As RiffMasterPreset, Optional ByVal amount As Single = 1!)
```

Applies a master processor preset.

```vb
RiffMasterApplyPreset RiffMasterFxGlue, 0.7!
RiffMasterApplyPreset RiffMasterFxCinematic, 0.55!
RiffMasterApplyPreset RiffMasterFxClean
```

### Practical Master Preset Usage

```vb
Public Sub SetNormalMix()
    RiffMasterApplyPreset RiffMasterFxClean
    RiffMasterVolume = 1!
End Sub

Public Sub SetCinematicMix()
    RiffMasterApplyPreset RiffMasterFxCinematic, 0.65!
    RiffBusFadeTo RiffBusMusic, 0.5!, 600
    RiffBusFadeTo RiffBusSfx, 0.85!, 300
End Sub

Public Sub SetNightMix()
    RiffMasterApplyPreset RiffMasterFxNight, 0.7!
    RiffBusFadeTo RiffBusMusic, 0.28!, 800
    RiffBusFadeTo RiffBusSfx, 0.55!, 300
End Sub
```

## Audio Buses API

### RiffBusVolume

```vb
Public Property Get RiffBusVolume(ByVal busID As RiffBusId) As Single
Public Property Let RiffBusVolume(ByVal busID As RiffBusId, ByVal value As Single)
```

Gets or sets a bus volume. Bus volume is a group gain multiplier. Values are clamped internally.

```vb
RiffBusVolume(RiffBusMusic) = 0.35!
RiffBusVolume(RiffBusSfx) = 0.9!
RiffBusVolume(RiffBusUi) = 1!
```

### RiffBusMuted

```vb
Public Property Get RiffBusMuted(ByVal busID As RiffBusId) As Boolean
Public Property Let RiffBusMuted(ByVal busID As RiffBusId, ByVal value As Boolean)
```

Mutes or unmutes a bus without changing its stored volume.

```vb
RiffBusMuted(RiffBusMusic) = True
RiffBusMuted(RiffBusMusic) = False
```

### RiffBusSolo

```vb
Public Property Get RiffBusSolo(ByVal busID As RiffBusId) As Boolean
Public Property Let RiffBusSolo(ByVal busID As RiffBusId, ByVal value As Boolean)
```

Solo mode lets one or more buses pass while non-soloed buses are effectively muted. If no bus is soloed, all non-muted buses pass normally.

```vb
RiffBusSolo(RiffBusMusic) = True
' Only soloed buses are heard.

RiffBusSolo(RiffBusMusic) = False
' Normal bus mix returns.
```

### RiffBusFadeTo

```vb
Public Sub RiffBusFadeTo(ByVal busID As RiffBusId, ByVal targetVolume As Single, Optional ByVal durationMs As Long = 250)
```

Smoothly fades a bus volume to a target value. Useful for music ducking, scene transitions, and pause menus.

```vb
' Duck music while dialogue plays.
RiffBusFadeTo RiffBusMusic, 0.18!, 500
RiffBusFadeTo RiffBusVoice, 1!, 200
```

```vb
' Restore after dialogue.
RiffBusFadeTo RiffBusMusic, 0.45!, 750
```

### RiffBusGetPeak

```vb
Public Sub RiffBusGetPeak(ByVal busID As RiffBusId, ByRef peakLeft As Single, ByRef peakRight As Single)
```

Gets decaying peak meters for a specific bus.

```vb
Dim l As Single, r As Single
RiffBusGetPeak RiffBusMusic, l, r
Debug.Print "Music bus peak:", l, r
```

### RiffBusReset

```vb
Public Sub RiffBusReset(ByVal busID As RiffBusId)
Public Sub RiffBusApplyPreset(ByVal busID As RiffBusId, ByVal preset As RiffEffectPreset, Optional ByVal amount As Single = 1!, Optional ByVal persistent As Boolean = True, Optional ByVal applyActive As Boolean = True)
Public Sub RiffBusClearEffects(ByVal busID As RiffBusId, Optional ByVal clearPersistent As Boolean = True)
Public Property Get RiffBusPresetEnabled(ByVal busID As RiffBusId) As Boolean
Public Property Let RiffBusPresetEnabled(ByVal busID As RiffBusId, ByVal value As Boolean)
Public Property Get RiffBusPreset(ByVal busID As RiffBusId) As RiffEffectPreset
Public Property Get RiffBusPresetAmount(ByVal busID As RiffBusId) As Single
```

Resets a bus to default state: normal volume, no mute, no solo, no pending fade, cleared peak values.

```vb
RiffBusReset RiffBusSfx
```

### RiffBusApplyPreset

```vb
Public Sub RiffBusApplyPreset( _
    ByVal busID As RiffBusId, _
    ByVal preset As RiffEffectPreset, _
    Optional ByVal amount As Single = 1!, _
    Optional ByVal persistent As Boolean = True, _
    Optional ByVal applyActive As Boolean = True _
)
```

Applies a voice effect preset to a whole bus.

When `applyActive` is `True`, the preset is applied to voices currently routed to the bus. When `persistent` is `True`, future voices routed to the bus automatically receive the preset as they are created or rerouted.

```vb
' Apply underwater treatment to current and future music/SFX voices.
RiffBusApplyPreset RiffBusMusic, RiffFxUnderwater, 0.55!
RiffBusApplyPreset RiffBusSfx, RiffFxUnderwater, 0.8!
```

Apply only to future voices:

```vb
RiffBusApplyPreset RiffBusVoice, RiffFxRadio, 0.65!, True, False
```

Apply only to currently active voices:

```vb
RiffBusApplyPreset RiffBusSfx, RiffFxSmallRoom, 0.4!, False, True
```

### RiffBusClearEffects

```vb
Public Sub RiffBusClearEffects(ByVal busID As RiffBusId, Optional ByVal clearPersistent As Boolean = True)
```

Clears effects from active voices on a bus. When `clearPersistent` is `True`, it also removes the stored persistent preset for future voices.

```vb
RiffBusClearEffects RiffBusMusic
RiffBusClearEffects RiffBusSfx
```

Keep the persistent preset but clear only active voices:

```vb
RiffBusClearEffects RiffBusVoice, False
```

### RiffBusPresetEnabled

```vb
Public Property Get RiffBusPresetEnabled(ByVal busID As RiffBusId) As Boolean
Public Property Let RiffBusPresetEnabled(ByVal busID As RiffBusId, ByVal value As Boolean)
```

Gets or sets whether the persistent bus preset is enabled.

```vb
RiffBusPresetEnabled(RiffBusMusic) = False
RiffBusPresetEnabled(RiffBusMusic) = True
```

### RiffBusPreset

```vb
Public Property Get RiffBusPreset(ByVal busID As RiffBusId) As RiffEffectPreset
```

Returns the stored persistent preset for a bus.

```vb
Debug.Print RiffBusPreset(RiffBusMusic)
```

### RiffBusPresetAmount

```vb
Public Property Get RiffBusPresetAmount(ByVal busID As RiffBusId) As Single
```

Returns the stored persistent preset amount for a bus.

```vb
Debug.Print RiffBusPresetAmount(RiffBusMusic)
```

### Bus Pattern: Settings Menu

```vb
Public Sub ApplyAudioSettings(ByVal masterPct As Single, ByVal musicPct As Single, ByVal sfxPct As Single, ByVal uiPct As Single)
    RiffMasterVolume = masterPct / 100!
    RiffBusVolume(RiffBusMusic) = musicPct / 100!
    RiffBusVolume(RiffBusSfx) = sfxPct / 100!
    RiffBusVolume(RiffBusUi) = uiPct / 100!
End Sub
```

## Asset Loading and Memory

### RiffLoad

```vb
Public Function RiffLoad(ByVal filePath As String) As Long
```

Decodes an audio file from disk through Media Foundation into Riff's dynamic buffer pool. Returns a buffer handle or `-1`. If the current buffer capacity is full, Riff attempts to grow the pool automatically before failing.

Supported formats depend on Windows Media Foundation codecs installed on the system, commonly WAV, MP3, AAC/M4A, WMA, and some FLAC/other formats depending on Windows version and codecs.

```vb
Dim click As Long
click = RiffLoad(ActivePresentation.Path & "\audio\click.wav")

If click = -1 Then
    MsgBox "Could not load click.wav. Error: " & CStr(RiffLastError)
End If
```

### Large Asset Loading Pattern

For projects with many sound effects, voice lines, or music tracks, reserve buffer capacity before loading everything. This keeps the load phase predictable and avoids repeated SafeArray growth during the preload loop.

```vb
Public Sub AudioLoadLargeLibrary()
    If Not RiffOpen() Then Exit Sub

    If Not RiffReserveBuffers(500) Then
        Debug.Print "Buffer reserve failed:", RiffLastError
        Exit Sub
    End If

    ' Load your asset table here.
End Sub
```

### RiffLoadFromMemory

```vb
Public Function RiffLoadFromMemory(ByRef audioData() As Byte) As Long
```

Decodes audio from a byte array. This is useful when audio is embedded in a workbook/presentation or decoded from Base64.

```vb
Dim bytes() As Byte
Dim buf As Long

' Fill bytes() with the full file contents first.
buf = RiffLoadFromMemory(bytes)
```

The byte array must contain a complete encoded audio file, not raw PCM samples.

### RiffUnload

```vb
Public Sub RiffUnload(ByVal bufferHandle As Long)
```

Unloads a buffer and frees its native memory. Any active voices using this buffer are stopped.

```vb
RiffUnload sndExplosion
sndExplosion = -1
```

### RiffBufferDurationSec

```vb
Public Property Get RiffBufferDurationSec(ByVal bufferHandle As Long) As Single
```

Returns a buffer's duration in seconds.

```vb
Debug.Print "Music duration:", RiffBufferDurationSec(sndMusic)
```

### Loading Pattern: Asset Table

```vb
Private sndClick As Long
Private sndCancel As Long
Private sndExplosion As Long
Private sndMusic As Long

Public Function AudioLoadAll() As Boolean
    If Not RiffOpen() Then Exit Function

    sndClick = RiffLoad(ActivePresentation.Path & "\audio\ui_click.wav")
    sndCancel = RiffLoad(ActivePresentation.Path & "\audio\ui_cancel.wav")
    sndExplosion = RiffLoad(ActivePresentation.Path & "\audio\explosion.wav")
    sndMusic = RiffLoad(ActivePresentation.Path & "\audio\music.mp3")

    AudioLoadAll = (sndClick >= 0 And sndCancel >= 0 And sndExplosion >= 0 And sndMusic >= 0)
End Function
```

## Export and Offline Rendering

### RiffExportBufferWav

```vb
Public Function RiffExportBufferWav(ByVal bufferHandle As Long, ByVal filePath As String) As Boolean
```

Exports a loaded buffer as 16-bit stereo PCM WAV.

```vb
Dim ok As Boolean
ok = RiffExportBufferWav(sndMusic, ActivePresentation.Path & "\export\music.wav")

If Not ok Then Debug.Print "Export failed:", RiffLastError
```

### RiffRenderOscillatorWav

```vb
Public Function RiffRenderOscillatorWav(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ByVal durationSec As Single, ByVal filePath As String) As Boolean
```

Renders a generated waveform or noise source to a WAV file without creating a real-time voice.

```vb
RiffRenderOscillatorWav RiffWaveSine, 440!, 2!, "C:\Temp\a4.wav"
RiffRenderOscillatorWav RiffWaveSquare, 220!, 1!, "C:\Temp\square.wav"
RiffRenderOscillatorWav RiffWavePinkNoise, 0!, 5!, "C:\Temp\rain_noise.wav"
```

Offline rendering is dry. It does not apply per-voice DSP presets or bus settings.

## Playback

### RiffPlay

```vb
Public Function RiffPlay( _
    ByVal bufferHandle As Long, _
    Optional ByVal busID As RiffBusId = RiffBusMain, _
    Optional ByVal looped As Boolean = False, _
    Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, _
    Optional ByVal pan As Single = 0! _
) As Long
```

The main playback function. Creates a voice from a loaded buffer. The bus, loop, volume, and pan are applied before the voice becomes active.

```vb
Dim v As Long
v = RiffPlay(sndExplosion, RiffBusSfx, False, 0.9!, 0.1!)
```

Looping music:

```vb
musicVoice = RiffPlay(sndMusic, RiffBusMusic, True, 0.45!, 0!)
```

### RiffPlayOnce

```vb
Public Function RiffPlayOnce( _
    ByVal bufferHandle As Long, _
    Optional ByVal busID As RiffBusId = RiffBusMain, _
    Optional ByVal looped As Boolean = False, _
    Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, _
    Optional ByVal pan As Single = 0! _
) As Long
```

Plays a buffer only if the same buffer is not already active on the requested bus. If already playing, returns the existing voice handle instead of starting a duplicate.

```vb
Public Sub EnsureMusic()
    musicVoice = RiffPlayOnce(sndMusic, RiffBusMusic, True, 0.45!)
End Sub
```

This is the preferred function for background music and persistent ambiences.

### RiffPlayOscillator

```vb
Public Function RiffPlayOscillator( _
    ByVal waveType As RiffWaveType, _
    ByVal frequencyHz As Single, _
    Optional ByVal busID As RiffBusId = RiffBusMain, _
    Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, _
    Optional ByVal pan As Single = 0!, _
    Optional ByVal durationSec As Single = 0! _
) As Long
```

Creates an oscillator voice. `frequencyHz` controls pitch: `220` is lower, `440` is concert A, `880` is one octave higher. When `durationSec` is `0`, the oscillator is continuous and must be stopped manually. When `durationSec` is greater than `0`, the oscillator becomes a finite one-shot and auto-releases.

```vb
' Continuous tone.
Dim hum As Long
hum = RiffPlayOscillator(RiffWaveSine, 110!, RiffBusSfx, 0.12!)

' Finite UI beep.
RiffPlayOscillator RiffWaveSquare, 880!, RiffBusUi, 0.22!, 0!, 0.07!

' Higher confirmation sparkle.
RiffPlayOscillator RiffWaveSine, 1760!, RiffBusUi, 0.12!, 0!, 0.045!
```

The stable build applies a tiny attack/release safety envelope to new/generated voices. This reduces clicks when a fast SFX starts, stops, or is stolen/reused.

### RiffPlayNoise

```vb
Public Function RiffPlayNoise( _
    Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, _
    Optional ByVal busID As RiffBusId = RiffBusMain, _
    Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, _
    Optional ByVal pan As Single = 0!, _
    Optional ByVal durationSec As Single = RIFF_DEFAULT_NOISE_DURATION_SEC _
) As Long
```

Creates a procedural noise voice. Unlike earlier builds, `RiffPlayNoise` is a finite one-shot by default. This prevents accidental infinite noise voices in gameplay code.

```vb
' Short one-shot noise hit.
RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.18!

' Softer filtered dust/wind hit.
Dim n As Long
n = RiffPlayNoise(RiffWavePinkNoise, RiffBusSfx, 0.08!, 0!, 0.12!)
RiffVoiceSetFilterHz n, 1800!, 120!
```

Use `durationSec:=0` only when you intentionally want continuous noise:

```vb
Dim continuous As Long
continuous = RiffPlayNoise(RiffWavePinkNoise, RiffBusMusic, 0.04!, 0!, 0!)
RiffFadeOut continuous, 2!
```

For readability, prefer `RiffPlayNoiseLoop` for continuous ambience.

### RiffPlayNoiseOnBus

```vb
Public Function RiffPlayNoiseOnBus( _
    ByVal busID As RiffBusId, _
    Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, _
    Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, _
    Optional ByVal pan As Single = 0!, _
    Optional ByVal durationSec As Single = RIFF_DEFAULT_NOISE_DURATION_SEC _
) As Long
```

Bus-first helper for short noise one-shots. This exists to avoid ambiguity in call sites where the first argument should visually be the destination bus.

```vb
RiffPlayNoiseOnBus RiffBusSfx
RiffPlayNoiseOnBus RiffBusUi, RiffWaveWhiteNoise, 0.05!, 0!, 0.025!
```

### RiffPlayNoiseLoop

```vb
Public Function RiffPlayNoiseLoop( _
    Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, _
    Optional ByVal busID As RiffBusId = RiffBusMain, _
    Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, _
    Optional ByVal pan As Single = 0! _
) As Long
```

Creates continuous procedural noise for ambience or synthesis. Stop or fade out the returned voice manually.

```vb
Dim rain As Long
rain = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.045!)
RiffVoiceSetFilterHz rain, 2400!, 100!
RiffVoiceApplyPreset rain, RiffFxRain, 0.55!
```

### RiffPlayNoiseLoopOnBus

```vb
Public Function RiffPlayNoiseLoopOnBus( _
    ByVal busID As RiffBusId, _
    Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, _
    Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, _
    Optional ByVal pan As Single = 0! _
) As Long
```

Bus-first continuous noise helper.

```vb
Dim wind As Long
wind = RiffPlayNoiseLoopOnBus(RiffBusMusic, RiffWavePinkNoise, 0.035!)
```

### Compatibility Playback Helpers

These remain available for older code. New code should prefer `RiffPlay`, `RiffPlayOnce`, `RiffPlayOscillator`, `RiffPlayNoise`, `RiffPlayNoiseOnBus`, and `RiffPlayNoiseLoop`.

```vb
Public Function RiffPlayBus(ByVal bufferHandle As Long, ByVal busID As RiffBusId, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!) As Long
Public Function RiffPlayBusOnce(ByVal bufferHandle As Long, ByVal busID As RiffBusId, Optional ByVal looped As Boolean = False, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!) As Long
Public Function RiffPlayOscillatorBus(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ByVal busID As RiffBusId, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!, Optional ByVal durationSec As Single = 0!) As Long
Public Function RiffPlayNoiseBus(Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, Optional ByVal busID As RiffBusId = RiffBusMain, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!, Optional ByVal durationSec As Single = RIFF_DEFAULT_NOISE_DURATION_SEC) As Long
```

## Voice State and Transport

### RiffPause

```vb
Public Sub RiffPause(ByVal voiceHandle As Long)
```

Pauses a voice while keeping its slot, source position, and DSP state.

```vb
RiffPause musicVoice
```

### RiffResume

```vb
Public Sub RiffResume(ByVal voiceHandle As Long)
```

Resumes a paused voice.

```vb
RiffResume musicVoice
```

### RiffStop

```vb
Public Sub RiffStop(ByVal voiceHandle As Long)
```

Immediately stops and frees a voice.

```vb
RiffStop musicVoice
musicVoice = -1
```

### RiffVoiceStop

```vb
Public Sub RiffVoiceStop(ByVal voiceHandle As Long)
```

Alias for `RiffStop` with a more consistent voice-oriented name.

```vb
RiffVoiceStop v
```

### RiffStopAll

```vb
Public Sub RiffStopAll()
```

Stops all active voices.

```vb
RiffStopAll
```

### RiffFadeIn

```vb
Public Sub RiffFadeIn(ByVal voiceHandle As Long, ByVal durationSec As Single)
```

Applies a linear fade-in multiplier. It does not change `RiffVoiceVolume`; it applies an additional fade envelope.

```vb
ambienceVoice = RiffPlay(sndForest, RiffBusMusic, True, 0.35!)
RiffFadeIn ambienceVoice, 3!
```

### RiffFadeOut

```vb
Public Sub RiffFadeOut(ByVal voiceHandle As Long, ByVal durationSec As Single)
```

Fades a voice out and then stops it automatically.

```vb
RiffFadeOut musicVoice, 2!
musicVoice = -1
```

### RiffSetLoopRegionSec

```vb
Public Sub RiffSetLoopRegionSec(ByVal voiceHandle As Long, ByVal startSec As Single, ByVal endSec As Single)
```

Sets a loop region in seconds for a buffer-backed voice.

```vb
Dim v As Long
v = RiffPlay(sndAmbience, RiffBusMusic, True, 0.4!)
RiffSetLoopRegionSec v, 1.2!, 12.8!
```

### RiffVoiceIsPlaying

```vb
Public Property Get RiffVoiceIsPlaying(ByVal voiceHandle As Long) As Boolean
```

Returns `True` when a voice is active and currently playing.

```vb
If RiffVoiceIsPlaying(musicVoice) Then Exit Sub
```

### RiffVoicePlaying

```vb
Public Function RiffVoicePlaying(ByVal voiceHandle As Long) As Boolean
```

Function-style equivalent for checking whether a voice is playing. Useful when you prefer function naming over property naming.

### RiffVoiceActive

```vb
Public Function RiffVoiceActive(ByVal voiceHandle As Long) As Boolean
```

Returns `True` if the voice slot is active, even if the voice is paused.

```vb
If RiffVoiceActive(musicVoice) Then
    Debug.Print "Music slot still allocated"
End If
```

### RiffVoiceIsPaused

```vb
Public Property Get RiffVoiceIsPaused(ByVal voiceHandle As Long) As Boolean
```

Returns `True` if the voice is active and paused.

### RiffFindPlayingVoice

```vb
Public Function RiffFindPlayingVoice(ByVal bufferHandle As Long, Optional ByVal busID As Long = -1) As Long
```

Finds the first currently playing voice that uses `bufferHandle`. If `busID` is `-1`, all buses are searched. If a valid bus id is provided, only that bus is searched.

```vb
Dim existing As Long
existing = RiffFindPlayingVoice(sndMusic, RiffBusMusic)

If existing <> -1 Then
    musicVoice = existing
End If
```

### RiffBufferIsPlaying

```vb
Public Function RiffBufferIsPlaying(ByVal bufferHandle As Long, Optional ByVal busID As Long = -1) As Boolean
```

Returns `True` if a buffer is currently playing. If `busID` is provided, restricts the search to that bus.

```vb
If Not RiffBufferIsPlaying(sndMusic, RiffBusMusic) Then
    musicVoice = RiffPlay(sndMusic, RiffBusMusic, True, 0.45!)
End If
```

## Voice Properties

### RiffVoiceBus

```vb
Public Property Get RiffVoiceBus(ByVal voiceHandle As Long) As RiffBusId
Public Property Let RiffVoiceBus(ByVal voiceHandle As Long, ByVal value As RiffBusId)
```

Gets or sets the bus route for a voice. New code should pass the bus directly to `RiffPlay`, but this property is useful for dynamic rerouting.

```vb
RiffVoiceBus(v) = RiffBusVoice
```

### RiffVoiceGetPeak

```vb
Public Sub RiffVoiceGetPeak(ByVal voiceHandle As Long, ByRef peakLeft As Single, ByRef peakRight As Single)
```

Gets decaying per-voice peak values.

```vb
Dim l As Single, r As Single
RiffVoiceGetPeak v, l, r
```

### RiffVoiceLoop

```vb
Public Property Get RiffVoiceLoop(ByVal voiceHandle As Long) As Boolean
Public Property Let RiffVoiceLoop(ByVal voiceHandle As Long, ByVal value As Boolean)
```

Gets or sets loop mode. Prefer passing `looped:=True` directly to `RiffPlay` for new playback.

```vb
RiffVoiceLoop(musicVoice) = True
```

### RiffVoicePositionSec

```vb
Public Property Get RiffVoicePositionSec(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoicePositionSec(ByVal voiceHandle As Long, ByVal value As Single)
```

Gets or seeks a buffer-backed voice position in seconds.

```vb
RiffVoicePositionSec(musicVoice) = 30!
Debug.Print RiffVoicePositionSec(musicVoice)
```

Has no meaningful effect for oscillator/noise voices.

### RiffVoiceVolume

```vb
Public Property Get RiffVoiceVolume(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceVolume(ByVal voiceHandle As Long, ByVal value As Single)
```

Voice gain multiplier. Use volume at play time when possible.

```vb
v = RiffPlay(sndClick, RiffBusUi, False, 0.7!)
RiffVoiceVolume(v) = 0.5!
```

### RiffVoicePitch

```vb
Public Property Get RiffVoicePitch(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoicePitch(ByVal voiceHandle As Long, ByVal value As Single)
```

Playback pitch/speed multiplier. `1.0` is normal. `2.0` is one octave up and twice as fast. `0.5` is one octave down and half speed.

```vb
RiffVoicePitch(v) = 1.25!
```

### RiffVoicePan

```vb
Public Property Get RiffVoicePan(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoicePan(ByVal voiceHandle As Long, ByVal value As Single)
```

Stereo pan. `-1` is left, `0` is center, `1` is right.

```vb
RiffVoicePan(v) = -0.35!
```

## Effect Helpers and Presets

### RiffVoiceClearEffects

```vb
Public Sub RiffVoiceClearEffects(ByVal voiceHandle As Long)
```

Resets the voice DSP parameters to neutral/dry state.

```vb
RiffVoiceClearEffects v
```

### RiffVoiceSetReverb

```vb
Public Sub RiffVoiceSetReverb(ByVal voiceHandle As Long, ByVal mix As Single, ByVal roomTime As Single)
```

Sets reverb mix and decay together.

```vb
RiffVoiceSetReverb v, 0.35!, 0.7!
```

### RiffVoiceSetDelay

```vb
Public Sub RiffVoiceSetDelay(ByVal voiceHandle As Long, ByVal delayTime As Single, ByVal feedback As Single, ByVal mix As Single)
```

Sets delay time, feedback, and wet mix together.

```vb
RiffVoiceSetDelay v, 0.28!, 0.45!, 0.5!
```

### RiffVoiceSetChorus

```vb
Public Sub RiffVoiceSetChorus(ByVal voiceHandle As Long, ByVal depth As Single, ByVal rate As Single)
```

Sets chorus depth and rate.

```vb
RiffVoiceSetChorus v, 0.4!, 1.2!
```

### RiffVoiceSetFlanger

```vb
Public Sub RiffVoiceSetFlanger(ByVal voiceHandle As Long, ByVal depth As Single, ByVal rate As Single, ByVal feedback As Single)
```

Sets flanger depth, rate, and feedback.

```vb
RiffVoiceSetFlanger v, 0.55!, 0.35!, 0.45!
```

### RiffVoiceSetFilter

```vb
Public Sub RiffVoiceSetFilter(ByVal voiceHandle As Long, ByVal lowPass As Single, ByVal highPass As Single)
```

Sets low-pass and high-pass controls together.

```vb
RiffVoiceSetFilter v, 0.35!, 0.05!
```

### RiffVoiceSetFilterHz

```vb
Public Sub RiffVoiceSetFilterHz(ByVal voiceHandle As Long, ByVal lowPassHz As Single, Optional ByVal highPassHz As Single = 0!)
```

Sets low-pass and high-pass filters using real frequency values in Hz instead of normalized filter controls.

```vb
' Muffled wall/underwater style.
RiffVoiceSetFilterHz v, 900!, 0!

' Radio/telephone band.
RiffVoiceSetFilterHz v, 3000!, 300!

' Remove sub-rumble but keep most of the sound open.
RiffVoiceSetFilterHz v, 0!, 80!
```

Parameters:

| Parameter | Meaning |
|:---|:---|
| `lowPassHz` | Frequencies above this point are reduced. Use `0` to disable low-pass filtering. |
| `highPassHz` | Frequencies below this point are reduced. Use `0` to disable high-pass filtering. |

Common ranges:

| Use | Typical Hz |
|:---|:---|
| Sub/rumble cleanup | high-pass `40` to `100` Hz |
| Telephone/radio | high-pass `300` Hz, low-pass `3000` Hz |
| Muffled wall/underwater | low-pass `600` to `1200` Hz |
| Soft rain/wind noise | low-pass `1800` to `3500` Hz |
| UI click/beep polish | low-pass `4000` to `8000` Hz |

### RiffVoiceApplyPreset

```vb
Public Sub RiffVoiceApplyPreset(ByVal voiceHandle As Long, ByVal preset As RiffEffectPreset, Optional ByVal amount As Single = 1!)
```

Applies a named effect preset. `amount` controls intensity and is typically clamped between `0` and `1`.

```vb
RiffVoiceApplyPreset v, RiffFxSmallRoom, 0.7!
RiffVoiceApplyPreset v, RiffFxLoFi, 0.45!
RiffVoiceApplyPreset v, RiffFxDry
```

Preset examples:

```vb
Public Sub PlayRadioVoice(ByVal buf As Long)
    Dim v As Long
    v = RiffPlay(buf, RiffBusVoice, False, 0.85!)
    If v <> -1 Then RiffVoiceApplyPreset v, RiffFxRadio, 0.7!
End Sub
```

```vb
Public Sub PlayUnderwaterImpact(ByVal buf As Long)
    Dim v As Long
    v = RiffPlay(buf, RiffBusSfx, False, 0.9!)
    If v <> -1 Then RiffVoiceApplyPreset v, RiffFxUnderwater, 1!
End Sub
```

## Musical Preset Packs

v1.0.9 expands Riff's preset set beyond utility effects into more musical sound-design packs. These presets are still implemented as combinations of existing per-voice DSP controls, so they remain lightweight and compatible with the normal voice pipeline.

### Warm and Tape-Style Presets

```vb
RiffVoiceApplyPreset v, RiffFxWarmTape, 0.6!
RiffVoiceApplyPreset v, RiffFxVHS, 0.55!
RiffVoiceApplyPreset v, RiffFxLoFi, 0.45!
```

Suggested uses:

| Preset | Use |
|:---|:---|
| `RiffFxWarmTape` | Music, ambience, narration, soft retro coloration. |
| `RiffFxVHS` | Old video, damaged memory, analog scenes. |
| `RiffFxLoFi` | Old sampler, dusty UI, retro ambience. |

### Space and Ambience Presets

```vb
RiffVoiceApplyPreset padVoice, RiffFxDreamPad, 0.75!
RiffVoiceApplyPreset caveVoice, RiffFxDarkCave, 0.8!
RiffVoiceApplyPreset ambienceVoice, RiffFxSoftFocus, 0.5!
```

Suggested uses:

| Preset | Use |
|:---|:---|
| `RiffFxDreamPad` | Soft pads, dream scenes, emotional ambience. |
| `RiffFxDarkCave` | Caves, horror scenes, distant indoor ambience. |
| `RiffFxSoftFocus` | Gentle smoothing for music or dialogue. |
| `RiffFxAmbient` | General spacious ambience. |

### Speaker and Device Presets

```vb
RiffVoiceApplyPreset voice, RiffFxTinySpeaker, 0.65!
RiffVoiceApplyPreset voice, RiffFxMegaphone, 0.7!
RiffVoiceApplyPreset ui, RiffFxGameBoy, 0.75!
```

Suggested uses:

| Preset | Use |
|:---|:---|
| `RiffFxTinySpeaker` | Phone/laptop/small radio style. |
| `RiffFxMegaphone` | Announcements, projected voices, PA systems. |
| `RiffFxGameBoy` | Crunchy handheld game tone. |
| `RiffFxRadio` | Radio/communication voice. |

### Environmental and Cinematic Presets

```vb
RiffVoiceApplyPreset windVoice, RiffFxWind, 0.7!
RiffVoiceApplyPreset rainVoice, RiffFxRain, 0.6!
RiffVoiceApplyPreset boomVoice, RiffFxCinematicBoom, 0.85!
RiffVoiceApplyPreset droneVoice, RiffFxHorrorDrone, 0.75!
```

Suggested uses:

| Preset | Use |
|:---|:---|
| `RiffFxWind` | Wind layers, filtered noise, exterior ambience. |
| `RiffFxRain` | Rain layers and natural noise ambience. |
| `RiffFxCinematicBoom` | Low impacts, hits, transitions, cinematic UI. |
| `RiffFxHorrorDrone` | Dark drones and unsettling textures. |

### Bus-Level Preset Pack Usage

Presets can also be applied to entire buses.

```vb
Public Sub EnterCave()
    RiffBusApplyPreset RiffBusMusic, RiffFxDarkCave, 0.55!
    RiffBusApplyPreset RiffBusSfx, RiffFxDarkCave, 0.35!
    RiffMasterApplyPreset RiffMasterFxDark, 0.5!
End Sub

Public Sub ExitCave()
    RiffBusClearEffects RiffBusMusic
    RiffBusClearEffects RiffBusSfx
    RiffMasterApplyPreset RiffMasterFxClean
End Sub
```

## DSP Parameters

### Bitcrusher

#### RiffVoiceBitDepth

```vb
Public Property Get RiffVoiceBitDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceBitDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

Controls quantization depth. `32` is effectively full quality. Lower values sound more digital/stepped.

```vb
RiffVoiceBitDepth(v) = 8!
```

#### RiffVoiceSampleRateReduction

```vb
Public Property Get RiffVoiceSampleRateReduction(ByVal voiceHandle As Long) As Long
Public Property Let RiffVoiceSampleRateReduction(ByVal voiceHandle As Long, ByVal value As Long)
```

Holds samples for multiple frames to simulate a lower sample rate. `1` disables the effect.

```vb
RiffVoiceSampleRateReduction(v) = 4
```

### Ring Modulator

#### RiffVoiceRingModFreq

```vb
Public Property Get RiffVoiceRingModFreq(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceRingModFreq(ByVal voiceHandle As Long, ByVal value As Single)
```

Carrier frequency for ring modulation. `0` disables.

```vb
RiffVoiceRingModFreq(v) = 120!
```

#### RiffVoiceRingModMix

```vb
Public Property Get RiffVoiceRingModMix(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceRingModMix(ByVal voiceHandle As Long, ByVal value As Single)
```

Wet/dry blend for ring modulation.

```vb
RiffVoiceRingModMix(v) = 0.5!
```

### Auto-Pan

#### RiffVoiceAutoPanRate

```vb
Public Property Get RiffVoiceAutoPanRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceAutoPanRate(ByVal voiceHandle As Long, ByVal value As Single)
```

Auto-pan LFO speed in Hz.

#### RiffVoiceAutoPanDepth

```vb
Public Property Get RiffVoiceAutoPanDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceAutoPanDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

Auto-pan intensity.

```vb
RiffVoiceAutoPanRate(v) = 0.5!
RiffVoiceAutoPanDepth(v) = 0.9!
```

### 3-Band EQ

#### RiffVoiceEqBass

```vb
Public Property Get RiffVoiceEqBass(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceEqBass(ByVal voiceHandle As Long, ByVal value As Single)
```

Low band gain. `1.0` is neutral.

#### RiffVoiceEqMid

```vb
Public Property Get RiffVoiceEqMid(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceEqMid(ByVal voiceHandle As Long, ByVal value As Single)
```

Mid band gain. `1.0` is neutral.

#### RiffVoiceEqTreble

```vb
Public Property Get RiffVoiceEqTreble(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceEqTreble(ByVal voiceHandle As Long, ByVal value As Single)
```

High band gain. `1.0` is neutral.

```vb
' Warm/dark voice.
RiffVoiceEqBass(v) = 1.2!
RiffVoiceEqMid(v) = 0.9!
RiffVoiceEqTreble(v) = 0.65!
```

### Compressor

#### RiffVoiceCompressorThreshold

```vb
Public Property Get RiffVoiceCompressorThreshold(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceCompressorThreshold(ByVal voiceHandle As Long, ByVal value As Single)
```

Amplitude threshold above which compression starts.

#### RiffVoiceCompressorRatio

```vb
Public Property Get RiffVoiceCompressorRatio(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceCompressorRatio(ByVal voiceHandle As Long, ByVal value As Single)
```

Compression ratio.

```vb
RiffVoiceCompressorThreshold(v) = 0.65!
RiffVoiceCompressorRatio(v) = 3!
```

### Flanger

#### RiffVoiceFlangerDepth

```vb
Public Property Get RiffVoiceFlangerDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceFlangerDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

Flanger wet depth.

#### RiffVoiceFlangerRate

```vb
Public Property Get RiffVoiceFlangerRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceFlangerRate(ByVal voiceHandle As Long, ByVal value As Single)
```

Flanger LFO rate.

#### RiffVoiceFlangerFeedback

```vb
Public Property Get RiffVoiceFlangerFeedback(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceFlangerFeedback(ByVal voiceHandle As Long, ByVal value As Single)
```

Flanger feedback amount.

```vb
RiffVoiceFlangerDepth(v) = 0.65!
RiffVoiceFlangerRate(v) = 0.25!
RiffVoiceFlangerFeedback(v) = 0.5!
```

### Distortion

#### RiffVoiceDistortion

```vb
Public Property Get RiffVoiceDistortion(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDistortion(ByVal voiceHandle As Long, ByVal value As Single)
```

Input gain before clipping/saturation. `1.0` is neutral.

```vb
RiffVoiceDistortion(v) = 2.5!
```

### Low-Pass and High-Pass Filters

#### RiffVoiceLowPass

```vb
Public Property Get RiffVoiceLowPass(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceLowPass(ByVal voiceHandle As Long, ByVal value As Single)
```

Normalized low-pass control. `1.0` is open/neutral. Lower values are more muffled.

```vb
RiffVoiceLowPass(v) = 0.35!
```

For exact cutoff-style control, prefer `RiffVoiceSetFilterHz` in new code:

```vb
RiffVoiceSetFilterHz v, 1200!, 0!
```

#### RiffVoiceHighPass

```vb
Public Property Get RiffVoiceHighPass(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceHighPass(ByVal voiceHandle As Long, ByVal value As Single)
```

Normalized high-pass control. `0.0` is neutral. Higher values remove more low end.

```vb
RiffVoiceHighPass(v) = 0.45!
```

### Stereo Width

#### RiffVoiceStereoWidth

```vb
Public Property Get RiffVoiceStereoWidth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceStereoWidth(ByVal voiceHandle As Long, ByVal value As Single)
```

Mid/side stereo width. `1.0` is neutral, `0.0` is mono, values above `1.0` widen.

```vb
RiffVoiceStereoWidth(v) = 1.6!
```

### Tremolo

#### RiffVoiceTremoloRate

```vb
Public Property Get RiffVoiceTremoloRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceTremoloRate(ByVal voiceHandle As Long, ByVal value As Single)
```

Tremolo LFO rate in Hz.

#### RiffVoiceTremoloDepth

```vb
Public Property Get RiffVoiceTremoloDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceTremoloDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

Tremolo depth.

```vb
RiffVoiceTremoloRate(v) = 5!
RiffVoiceTremoloDepth(v) = 0.5!
```

### Chorus

#### RiffVoiceChorusDepth

```vb
Public Property Get RiffVoiceChorusDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceChorusDepth(ByVal voiceHandle As Long, ByVal value As Single)
```

Chorus wet depth.

#### RiffVoiceChorusRate

```vb
Public Property Get RiffVoiceChorusRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceChorusRate(ByVal voiceHandle As Long, ByVal value As Single)
```

Chorus modulation rate.

```vb
RiffVoiceChorusDepth(v) = 0.35!
RiffVoiceChorusRate(v) = 1.2!
```

### Reverb

#### RiffVoiceReverbMix

```vb
Public Property Get RiffVoiceReverbMix(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceReverbMix(ByVal voiceHandle As Long, ByVal value As Single)
```

Reverb wet mix.

#### RiffVoiceReverbTime

```vb
Public Property Get RiffVoiceReverbTime(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceReverbTime(ByVal voiceHandle As Long, ByVal value As Single)
```

Reverb decay/time control.

```vb
RiffVoiceReverbMix(v) = 0.35!
RiffVoiceReverbTime(v) = 0.75!
```

### Delay

#### RiffVoiceDelayTime

```vb
Public Property Get RiffVoiceDelayTime(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDelayTime(ByVal voiceHandle As Long, ByVal value As Single)
```

Delay time in seconds.

#### RiffVoiceDelayFeedback

```vb
Public Property Get RiffVoiceDelayFeedback(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDelayFeedback(ByVal voiceHandle As Long, ByVal value As Single)
```

Delay feedback amount.

#### RiffVoiceDelayMix

```vb
Public Property Get RiffVoiceDelayMix(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDelayMix(ByVal voiceHandle As Long, ByVal value As Single)
```

Delay wet mix.

```vb
RiffVoiceDelayTime(v) = 0.25!
RiffVoiceDelayFeedback(v) = 0.35!
RiffVoiceDelayMix(v) = 0.45!
```

## Practical Recipes

### Full Startup and Shutdown Pattern

```vb
Private sndClick As Long
Private sndExplosion As Long
Private sndMusic As Long
Private musicVoice As Long

Public Sub GameAudioStart()
    If Not RiffOpen() Then
        MsgBox "Failed to start Riff. Error: " & CStr(RiffLastError), vbCritical
        Exit Sub
    End If

    RiffAutoSuspendTimer = True
    RiffMasterVolume = 1!
    RiffBusVolume(RiffBusMusic) = 0.4!
    RiffBusVolume(RiffBusSfx) = 0.9!
    RiffBusVolume(RiffBusUi) = 1!

    sndClick = RiffLoad(ActivePresentation.Path & "\audio\click.wav")
    sndExplosion = RiffLoad(ActivePresentation.Path & "\audio\explosion.wav")
    sndMusic = RiffLoad(ActivePresentation.Path & "\audio\music.mp3")

    musicVoice = -1
End Sub

Public Sub GameAudioStop()
    RiffClose
End Sub
```

### Button Click Sound

```vb
Public Sub PlayClick()
    RiffPlay sndClick, RiffBusUi, False, 0.65!
End Sub
```

### Explosion With Random Pan and Rumble

```vb
Public Sub PlayExplosion()
    Dim v As Long
    Dim rumble As Long
    Dim pan As Single

    pan = (Rnd() * 2!) - 1!

    v = RiffPlay(sndExplosion, RiffBusSfx, False, 0.95!, pan)
    If v <> -1 Then
        RiffVoiceApplyPreset v, RiffFxSmallRoom, 0.35!
        RiffVoiceDistortion(v) = 1.4!
    End If

    rumble = RiffPlayNoise(RiffWaveBrownNoise, RiffBusSfx, 0.08!, pan)
    If rumble <> -1 Then
        RiffVoiceLowPass(rumble) = 0.2!
        RiffFadeOut rumble, 0.45!
    End If
End Sub
```

### Music That Never Duplicates

```vb
Public Sub EnsureMusicPlaying()
    musicVoice = RiffPlayOnce(sndMusic, RiffBusMusic, True, 0.42!)
End Sub
```

### Music Fade Out

```vb
Public Sub StopMusicSmooth()
    If musicVoice <> -1 Then
        RiffFadeOut musicVoice, 1.2!
        musicVoice = -1
    End If
End Sub
```

### Pause Menu Ducking

```vb
Public Sub EnterPauseMenu()
    RiffBusFadeTo RiffBusMusic, 0.18!, 300
    RiffBusFadeTo RiffBusSfx, 0.55!, 200
End Sub

Public Sub LeavePauseMenu()
    RiffBusFadeTo RiffBusMusic, 0.42!, 500
    RiffBusFadeTo RiffBusSfx, 0.9!, 200
End Sub
```

### Dialogue Ducking

```vb
Public Sub StartDialogue()
    RiffBusFadeTo RiffBusMusic, 0.2!, 450
    RiffBusFadeTo RiffBusSfx, 0.55!, 250
    RiffBusVolume(RiffBusVoice) = 1!
End Sub

Public Sub EndDialogue()
    RiffBusFadeTo RiffBusMusic, 0.45!, 700
    RiffBusFadeTo RiffBusSfx, 0.9!, 300
End Sub
```

### Rain Ambience With Pink Noise

```vb
Private rainVoice As Long

Public Sub StartRain()
    If rainVoice <> -1 Then
        If RiffVoiceActive(rainVoice) Then Exit Sub
    End If

    rainVoice = RiffPlayNoise(RiffWavePinkNoise, RiffBusMusic, 0.055!)
    If rainVoice <> -1 Then
        RiffVoiceLoop(rainVoice) = True
        RiffVoiceLowPass(rainVoice) = 0.55!
        RiffVoiceHighPass(rainVoice) = 0.05!
        RiffVoiceStereoWidth(rainVoice) = 1.4!
        RiffVoiceApplyPreset rainVoice, RiffFxAmbient, 0.35!
        RiffFadeIn rainVoice, 2!
    End If
End Sub

Public Sub StopRain()
    If rainVoice <> -1 Then
        RiffFadeOut rainVoice, 2!
        rainVoice = -1
    End If
End Sub
```

### Retro UI Beep

```vb
Public Sub PlayRetroBeep()
    Dim v As Long
    v = RiffPlayOscillator(RiffWaveSquare, 880!, RiffBusUi, 0.18!)

    If v <> -1 Then
        RiffVoiceBitDepth(v) = 8!
        RiffVoiceSampleRateReduction(v) = 2
        RiffFadeOut v, 0.09!
    End If
End Sub
```

### Radio Voice

```vb
Public Sub PlayRadioLine(ByVal voiceBuffer As Long)
    Dim v As Long
    v = RiffPlay(voiceBuffer, RiffBusVoice, False, 0.85!)

    If v <> -1 Then
        RiffVoiceApplyPreset v, RiffFxRadio, 0.65!
        RiffVoiceCompressorThreshold(v) = 0.55!
        RiffVoiceCompressorRatio(v) = 4!
    End If
End Sub
```

### Underwater Scene

```vb
Public Sub EnterUnderwater()
    RiffBusFadeTo RiffBusMusic, 0.3!, 600
    RiffBusFadeTo RiffBusSfx, 0.6!, 300
End Sub

Public Sub PlayUnderwaterSfx(ByVal buf As Long)
    Dim v As Long
    v = RiffPlay(buf, RiffBusSfx, False, 0.8!)

    If v <> -1 Then
        RiffVoiceApplyPreset v, RiffFxUnderwater, 0.8!
    End If
End Sub
```

### Simple VU Meter

```vb
Public Sub UpdateMeters()
    Dim mL As Single, mR As Single
    Dim sL As Single, sR As Single

    RiffMasterGetPeak mL, mR
    RiffBusGetPeak RiffBusSfx, sL, sR

    Debug.Print "Master:", Format(mL, "0.00"), Format(mR, "0.00")
    Debug.Print "SFX:", Format(sL, "0.00"), Format(sR, "0.00")
End Sub
```

### Stability Test

```vb
Public Sub AudioStressDiagnostics()
    RiffResetAdaptiveStats
    RiffPlayOnce sndMusic, RiffBusMusic, True, 0.4!

    Debug.Print "Switch tabs or stress the host for a few seconds, then run PrintRiffDiagnostics."
End Sub

Public Sub PrintRiffDiagnostics()
    Debug.Print "Adaptive queue:", RiffAdaptiveQueueMs
    Debug.Print "Underruns:", RiffUnderrunCount
    Debug.Print "Padding:", RiffLastPaddingFrames
    Debug.Print "Available:", RiffLastFramesAvailable
    Debug.Print "Written:", RiffLastFramesWritten
End Sub
```

## Best Practices

### Load During Initialization

Avoid calling `RiffLoad` during gameplay, animation ticks, or button spam.

Good:

```vb
Public Sub Init()
    sndClick = RiffLoad(pathClick)
End Sub

Public Sub ButtonClick()
    RiffPlay sndClick, RiffBusUi
End Sub
```

Bad:

```vb
Public Sub ButtonClick()
    Dim snd As Long
    snd = RiffLoad(pathClick)
    RiffPlay snd, RiffBusUi
End Sub
```

### Prefer Unified Playback

Use:

```vb
RiffPlay snd, RiffBusSfx, False, 0.8!, 0!
```

Instead of:

```vb
v = RiffPlay(snd)
RiffVoiceBus(v) = RiffBusSfx
RiffVoiceVolume(v) = 0.8!
```

The unified call applies routing and volume before the voice becomes active.

### Use RiffPlayOnce for Music

Do not start background music repeatedly from slide events or button callbacks.

```vb
musicVoice = RiffPlayOnce(sndMusic, RiffBusMusic, True, 0.45!)
```

### Use Buses for User Settings

Do not manually loop through all voices to change all music/SFX volumes. Use buses.

```vb
RiffBusVolume(RiffBusMusic) = musicVolume
RiffBusVolume(RiffBusSfx) = sfxVolume
```

### Avoid Busy-Wait Loops

Avoid loops like this while audio is playing:

```vb
Do While Timer < t + seconds
    DoEvents
Loop
```

They can starve the Office host and worsen audio timing. If you need waits in demos, use a lightweight wait that yields more gently, or design demos around scheduled callbacks.

### Always Close on Exit

Riff is more stable when it is explicitly closed before the Office host closes or the VBA project resets.

```vb
RiffClose
```

### Keep DSP Reasonable

Every active voice has its own DSP chain. Thirty-two voices with reverb, chorus, flanger, delay, compressor, ring mod, and filters all active can become expensive in VBA. Use presets and effects selectively.

### Prefer Bus Presets for Scene-Wide Effects

For global scene states such as underwater, cave, dream, radio, or horror, use bus-level presets instead of manually applying a preset to every individual voice.

```vb
RiffBusApplyPreset RiffBusMusic, RiffFxUnderwater, 0.55!
RiffBusApplyPreset RiffBusSfx, RiffFxUnderwater, 0.75!
```

### Use Master Processors for Final Polish, Not Per-Sound Design

Use voice presets for individual sound design and master presets for broad final mix shaping.

```vb
RiffVoiceApplyPreset voice, RiffFxRadio, 0.65!
RiffMasterApplyPreset RiffMasterFxGlue, 0.5!
```

Avoid stacking heavy master coloration with many already-heavy voice presets unless that is the intended effect.

### Keep Procedural Sources Finite by Default

Use `RiffPlayNoise` for one-shot procedural effects and `RiffPlayNoiseLoop` only for intentional continuous ambience. This prevents accidental accumulation in gameplay loops.

```vb
' Good for SFX:
RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.12!, 0!, 0.05!

' Good for ambience:
ambience = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.04!)
```

### Tune Burst Caps for Game Feel

In NEXT 1.2.1 the explicit burst caps default to `0`, meaning disabled. This keeps large projects from feeling artificially limited. For dense action scenes, you can opt into caps carefully while watching diagnostics.

```vb
RiffMaxVoicesPerBuffer = 4
RiffMaxVoicesPerBus = 18
Debug.Print RiffActiveVoiceCount()
```

If the same SFX feels too aggressively limited, slightly increase `RiffMaxVoicesPerBuffer`. If gameplay stutters under heavy SFX, lower it.

### Prefer Hz Filters for Sound Design

Normalized low/high-pass properties are lightweight and compatible with old code, but Hz helpers are easier to reason about when designing game audio.

```vb
RiffVoiceSetFilterHz v, 3000!, 300!  ' radio/phone
RiffVoiceSetFilterHz v, 900!, 0!     ' muffled
RiffVoiceSetFilterHz v, 0!, 80!      ' remove sub-rumble
```

### Keep the VBE Responsive During Development

If you are testing inside the editor, prefer explicit cleanup in stop/shutdown paths. The VBE-safe timer reduces the chance of IntelliSense getting stuck, but explicit cleanup is still the cleanest workflow when editing a lot. If you are going to change code while audio may still be active, pause the render timer first with `RiffPrepareForVbeEdit` and resume it with `RiffResumeAfterVbeEdit` after editing.

```vb
Public Sub AudioDevReset()
    RiffStopAll
    RiffSuspend
End Sub

Public Sub AudioEditorEmergencyStop()
    RiffEditorEmergencyStop
End Sub

Public Sub AudioPrepareCodeEdit()
    RiffPrepareForVbeEdit
End Sub

Public Sub AudioResumeAfterCodeEdit()
    If Not RiffResumeAfterVbeEdit() Then Debug.Print RiffLastError
End Sub

Public Sub AudioFullShutdown()
    RiffClose
End Sub
```

Use `RiffPrepareForVbeEdit` before live code edits, `RiffResumeAfterVbeEdit` after the edit, and `RiffEditorEmergencyStop` only when the VBE/editor is stuck after a manual Stop/Reset or an interrupted test run. Use `RiffClose` for real shutdown and release of audio resources.

For gameplay tests, `RiffAutoSuspendTimer = False` is acceptable because the current build keeps audio warm only briefly while idle. For macro tools, examples, and code snippets that users may run from the VBE, prefer `True`.

### Avoid Applying Heavy Presets Per Frame

`RiffVoiceApplyPreset` is much faster in the lazy temporal-buffer build, but it is still setup work. Apply presets when a voice starts, when a bus/scene changes, or when a sound-design state changes. Avoid reapplying the same preset every frame.

Good:

```vb
v = RiffPlay(sndRadioVoice, RiffBusVoice, False, 0.8!)
If v <> -1 Then RiffVoiceApplyPreset v, RiffFxRadio, 0.7!
```

Bad:

```vb
' Do not do this every frame for the same active voice.
RiffVoiceApplyPreset v, RiffFxRadio, 0.7!
```

## Performance Benchmarks

The benchmark modules used during the performance pass measured burst playback, generated sources, preset setup, memory stability, active-voice cleanup, and game-loop behavior. The exact values depend on host, CPU, Office version, and audio device, but the tested performance build reached roughly:

```txt
Short dry RiffPlay:       ~11 us/call
Long dry RiffPlay:        ~13 us/call
Noise one-shot:           ~14-15 us/call
Oscillator one-shot:      ~13-15 us/call
Preset/DSP setup:         ~20 us/call
Failed handles:           0
Active voices after wait: 0
Game-loop underruns:      0 in the final pumped-wait benchmark
Final active voices:      0
Memory delta:             stable / no observed leak in the benchmark runs
```

The most important benchmark signs are not just low `us/call`. For real gameplay stability, also watch:

```txt
Failed handles            should remain 0 or intentionally capped/limited
Active voices after wait  should return to 0 for finite sounds
Underrun delta            should stay low, ideally 0 in normal test loops
Private MB delta          should stay stable across repeated runs
```

If a synthetic benchmark uses `Sleep` without pumping messages, the VBE/Office timer may not process enough callbacks during the wait. For audio cleanup tests, a pumped wait that periodically calls `DoEvents` gives a more realistic result inside Office.

## Troubleshooting

### `RiffOpen` returns `False`

Check:

- Audio device connected and enabled.
- Windows audio service is running.
- Host is Windows Office, not Mac Office.
- `RiffLastError` for specific failure.

```vb
If Not RiffOpen() Then
    Debug.Print "RiffOpen failed:", RiffLastError
End If
```

### Sound plays twice or music overlaps

Use `RiffPlayOnce` for long-running music/ambience:

```vb
musicVoice = RiffPlayOnce(sndMusic, RiffBusMusic, True, 0.45!)
```

Or check manually:

```vb
If RiffBufferIsPlaying(sndMusic, RiffBusMusic) Then Exit Sub
```

### Audio clicks when stopping

Use fade-out for musical stops:

```vb
RiffFadeOut musicVoice, 0.5!
```

Use immediate `RiffStop` only for hard cuts.

### Audio stutters when switching tabs

This can happen because Office/VBA is not a real-time host. The adaptive buffer should reduce this. Inspect:

```vb
Debug.Print RiffAdaptiveQueueMs
Debug.Print RiffUnderrunCount
```

If underruns increase, reduce expensive UI work, avoid busy-wait loops, preload assets, and keep DSP counts reasonable.

### VBE autocomplete stops or VBE looks like it is still running

This usually means a timer/callback is still active and the VBA project is being kept in a running state. In affected builds, this could make IntelliSense stop working or make the VBE title bar flicker between normal and running states.

There are two related cases:

```txt
Idle timer case       -> audio finished, but timer stayed alive too long
Manual Stop/Reset case -> VBE reset variables while an old timer callback was pending
```

The VBE-safe performance build reduces the first case by stopping the render timer automatically after a short idle period, even when `RiffAutoSuspendTimer = False` was used for game-style warm playback. The stop-safe build also handles the second case by having stale timer callbacks kill their received timer id instead of continuing to wake the editor.

For immediate recovery during development, call:

```vb
RiffStopAll
RiffSuspend
```

If the editor is still visibly stuck, or if this happened after clicking the VBE Stop/Reset button, use the editor emergency helper:

```vb
RiffEditorEmergencyStop
```

For full cleanup and resource release:

```vb
RiffClose
```

Recommended settings:

```vb
' Normal Office/editor workflows.
RiffAutoSuspendTimer = True

' Gameplay loop / rapid SFX testing. VBE-safe idle guard still stops later.
RiffAutoSuspendTimer = False
```

If IntelliSense is still stuck after a manual break or host-level interruption, run `RiffEditorEmergencyStop` first. If you want a full engine restart, then run `RiffClose` and start again with `RiffOpen`.


### Project crashes or corrupts after editing code while music is playing

This is the live-edit scenario that NEXT 1.2.1 specifically tries to make safer. The risky pattern is:

```txt
Riff timer/callback is active
VBE receives code edits
VBA recompiles or changes project state
callback touches VBA/native state at the wrong moment
```

Use the manual edit workflow before changing code while Riff is initialized:

```vb
RiffPrepareForVbeEdit
' Edit the VBA project.
RiffResumeAfterVbeEdit
```

For a team project, it is a good idea to expose these as simple dev buttons/macros so testers do not need to remember the exact function names.

```vb
Public Sub SafeEditOn()
    RiffPrepareForVbeEdit
End Sub

Public Sub SafeEditOff()
    If Not RiffResumeAfterVbeEdit() Then Debug.Print RiffLastError
End Sub
```

Avoid relying on the automatic mode until the project has been tested in your workflow:

```vb
RiffEditorSafeMode = True
```

The automatic guard can be useful, but it is intentionally disabled by default because it may pause the render timer when you are only trying to test audio from the VBE.

### Fast repeated SFX eventually stutter or feel like they accumulate

Use the runtime counters to confirm whether voices are actually accumulating:

```vb
Debug.Print "Active:", RiffActiveVoiceCount()
Debug.Print "SFX:", RiffBusVoiceCount(RiffBusSfx)
Debug.Print "This buffer:", RiffBufferVoiceCount(sndHit, RiffBusSfx)
```

Expected behavior after a short burst:

```txt
Active voices should fall back toward 0 after finite sounds finish.
Private memory should not climb continuously between test runs.
Peak voices should not stay stuck at `RiffMaxVoices`, and `RiffMaxVoices` itself may be higher than 32 after dynamic growth or manual reservation.
```

### `RiffPlay` is fast but preset setup is still heavier

This is expected. A dry `RiffPlay` only needs to allocate/reset a voice and route it. A preset may configure filters, EQ, compressor, delay, reverb, chorus, flanger, stereo width, and other DSP parameters.

The current build uses lazy temporal-buffer preparation, so presets are much cheaper than earlier builds. Still, prefer applying presets once at voice creation or at bus/scene transitions, not every frame.

```vb
Dim v As Long
v = RiffPlay(sndHit, RiffBusSfx, False, 0.8!)
If v <> -1 Then RiffVoiceApplyPreset v, RiffFxSmallRoom, 0.35!
```

### Underruns appear during synthetic tests but not during gameplay

Check whether the test is blocking the Office message loop. A raw `Sleep` can prevent the VBA timer callback from running during the wait. Use a pumped wait with periodic `DoEvents` when testing cleanup, voice lifetime, and timer behavior inside Office.

Also verify that assets are already loaded and that the test is not decoding files during the benchmark.


Recommended fixes:

```vb
RiffVoiceStealingEnabled = True
RiffMaxVoicesPerBuffer = 4
RiffMaxVoicesPerBus = 18
```

For procedural effects, avoid accidentally creating continuous sources:

```vb
' Short one-shot noise; safe for repeated gameplay triggers.
RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.12!, 0!, 0.04!

' Continuous noise; only for ambience, stop/fade manually.
wind = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.04!)
```

### A sound starts with a small click/pop or a weird tiny artifact

The stable build applies tiny safety ramps to new and released voices, but clicks can still happen if volume, filters, delay feedback, or preset changes are extreme. Prefer short finite durations and avoid forcing abrupt stop/start loops every frame.

```vb
v = RiffPlayOscillator(RiffWaveSquare, 880!, RiffBusUi, 0.18!, 0!, 0.05!)
```

For manually stopped continuous sounds, fade instead of stopping instantly:

```vb
RiffFadeOut windVoice, 0.25!
```

### A preset sounds wrong or too intense

Preset amount is clamped, and the stable build sanitizes DSP values after preset application. Still, presets stack with existing voice/bus/master settings. Clear effects first when testing a preset in isolation:

```vb
RiffVoiceClearEffects v
RiffVoiceApplyPreset v, RiffFxRadio, 0.7!
```

For bus-wide processing, remember that persistent presets affect future voices too:

```vb
RiffBusClearEffects RiffBusSfx
RiffBusPresetEnabled(RiffBusSfx) = False
```

### No sound from a bus

Check mute/solo state:

```vb
Debug.Print RiffBusVolume(RiffBusMusic)
Debug.Print RiffBusMuted(RiffBusMusic)
Debug.Print RiffBusSolo(RiffBusMusic)
```

If any bus is soloed, non-solo buses may be silent.

### `RiffPlay` returns `-1`

Check:

- Buffer handle is valid.
- Engine is initialized.
- Voice capacity is available or can grow. If not, check memory pressure, `RiffMaxVoices`, `RiffActiveVoiceCount`, and any optional burst caps.
- `RiffLastError`.

```vb
Dim v As Long
v = RiffPlay(sndClick, RiffBusUi)
If v = -1 Then Debug.Print RiffLastError
```

### `RiffLoad` returns `-1`

Check:

- File path exists.
- Buffer capacity can grow or was reserved successfully.
- Office has enough memory for the decoded audio.
- Media Foundation supports the file.
- Buffer capacity can grow and memory allocation is succeeding.
- There is enough memory.

```vb
If Dir$(path) = vbNullString Then Debug.Print "Missing file"
```

### A whole bus sounds filtered or processed unexpectedly

Check whether a persistent bus preset is enabled.

```vb
Debug.Print RiffBusPresetEnabled(RiffBusMusic)
Debug.Print RiffBusPreset(RiffBusMusic)
Debug.Print RiffBusPresetAmount(RiffBusMusic)
```

Clear it if needed:

```vb
RiffBusClearEffects RiffBusMusic
```

### The entire mix sounds too dark, compressed, or saturated

Check the master processor state.

```vb
Debug.Print RiffMasterProcessorEnabled
Debug.Print RiffMasterLowPass
Debug.Print RiffMasterDrive
Debug.Print RiffMasterCompressorThreshold
Debug.Print RiffMasterCompressorRatio
```

Reset the master stage:

```vb
RiffMasterClearProcessors
RiffMasterApplyPreset RiffMasterFxClean
```

## Stable Gameplay Build Notes

These notes describe behavior that is important when using the current stable implementation behind the v1.0.9 documentation target.

### Noise Semantics

`RiffPlayNoise` is now intended for short one-shot noise SFX by default. This is different from older examples where noise was commonly started and then faded manually. The old continuous behavior is still available through `durationSec:=0`, but `RiffPlayNoiseLoop` is clearer.

```vb
RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.1!, 0!, 0.05!
ambience = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.04!)
```

### Voice Reuse Hygiene

When a voice is released, stopped, or stolen, the stable build clears lifetime counters, fade state, loop state, source metadata, and DSP state needed to prevent stale behavior from leaking into the next reused voice slot. This is especially important under heavy repeated SFX loads.

### Preset Sanitation

Voice and bus presets are treated as recipes, then clamped/sanitized. This prevents feedback, filter, delay, stereo width, drive, compressor, and modulation values from landing outside practical ranges after aggressive preset combinations.

### Benchmark Expectations

A healthy repeated-SFX benchmark should show:

```txt
Active voices after wait: 0
Peak active voices: below the hard pool limit
Private memory final-start: near 0 MB
Underrun delta: low and not continuously rising
```

## Complete Public API Index

### Enums

```vb
Public Enum RiffWaveType
Public Enum RiffBusId
Public Enum RiffEffectPreset
Public Enum RiffMasterPreset
Public Enum RiffErrorCode
```

### Engine

```vb
Public Function RiffOpen() As Boolean
Public Sub RiffClose()
Public Sub RiffEditorEmergencyStop()
Public Property Get RiffIsInitialized() As Boolean
Public Property Get RiffAutoSuspendTimer() As Boolean
Public Property Let RiffAutoSuspendTimer(ByVal value As Boolean)
Public Sub RiffSuspend()
Public Sub RiffPrepareForVbeEdit()
Public Function RiffResumeAfterVbeEdit() As Boolean
Public Function RiffWake() As Boolean
Public Property Get RiffEditorSafeMode() As Boolean
Public Property Let RiffEditorSafeMode(ByVal value As Boolean)
Public Property Get RiffEditorTimerSuspended() As Boolean
```

### Diagnostics

```vb
Public Property Get RiffLastError() As RiffErrorCode
Public Property Get RiffMaxVoices() As Long
Public Property Get RiffMaxBuffers() As Long
Public Property Get RiffMaxBuses() As Long
Public Function RiffReserveBuffers(ByVal capacity As Long) As Boolean
Public Function RiffReserveVoices(ByVal capacity As Long) As Boolean
Public Function RiffReserveBuses(ByVal capacity As Long) As Boolean
Public Property Get RiffAdaptiveQueueMs() As Long
Public Property Get RiffUnderrunCount() As Long
Public Property Get RiffLastPaddingFrames() As Long
Public Property Get RiffLastFramesAvailable() As Long
Public Property Get RiffLastFramesWritten() As Long
Public Sub RiffResetAdaptiveStats()
```

### Master Processors and Buses

```vb
Public Property Get RiffMasterVolume() As Single
Public Property Let RiffMasterVolume(ByVal value As Single)
Public Property Get RiffSoftClipEnabled() As Boolean
Public Property Let RiffSoftClipEnabled(ByVal value As Boolean)
Public Property Get RiffMasterProcessorEnabled() As Boolean
Public Property Let RiffMasterProcessorEnabled(ByVal value As Boolean)
Public Property Get RiffMasterLowPass() As Single
Public Property Let RiffMasterLowPass(ByVal value As Single)
Public Property Get RiffMasterHighPass() As Single
Public Property Let RiffMasterHighPass(ByVal value As Single)
Public Property Get RiffMasterEqBass() As Single
Public Property Let RiffMasterEqBass(ByVal value As Single)
Public Property Get RiffMasterEqMid() As Single
Public Property Let RiffMasterEqMid(ByVal value As Single)
Public Property Get RiffMasterEqTreble() As Single
Public Property Let RiffMasterEqTreble(ByVal value As Single)
Public Property Get RiffMasterCompressorThreshold() As Single
Public Property Let RiffMasterCompressorThreshold(ByVal value As Single)
Public Property Get RiffMasterCompressorRatio() As Single
Public Property Let RiffMasterCompressorRatio(ByVal value As Single)
Public Property Get RiffMasterDrive() As Single
Public Property Let RiffMasterDrive(ByVal value As Single)
Public Property Get RiffMasterStereoWidth() As Single
Public Property Let RiffMasterStereoWidth(ByVal value As Single)
Public Property Get RiffMasterOutputGain() As Single
Public Property Let RiffMasterOutputGain(ByVal value As Single)
Public Sub RiffMasterClearProcessors()
Public Sub RiffMasterApplyPreset(ByVal preset As RiffMasterPreset, Optional ByVal amount As Single = 1!)
Public Sub RiffMasterGetPeak(ByRef peakLeft As Single, ByRef peakRight As Single)

Public Property Get RiffBusVolume(ByVal busID As RiffBusId) As Single
Public Property Let RiffBusVolume(ByVal busID As RiffBusId, ByVal value As Single)
Public Property Get RiffBusMuted(ByVal busID As RiffBusId) As Boolean
Public Property Let RiffBusMuted(ByVal busID As RiffBusId, ByVal value As Boolean)
Public Property Get RiffBusSolo(ByVal busID As RiffBusId) As Boolean
Public Property Let RiffBusSolo(ByVal busID As RiffBusId, ByVal value As Boolean)
Public Sub RiffBusFadeTo(ByVal busID As RiffBusId, ByVal targetVolume As Single, Optional ByVal durationMs As Long = 250)
Public Sub RiffBusGetPeak(ByVal busID As RiffBusId, ByRef peakLeft As Single, ByRef peakRight As Single)
Public Sub RiffBusReset(ByVal busID As RiffBusId)
Public Sub RiffBusApplyPreset(ByVal busID As RiffBusId, ByVal preset As RiffEffectPreset, Optional ByVal amount As Single = 1!, Optional ByVal persistent As Boolean = True, Optional ByVal applyActive As Boolean = True)
Public Sub RiffBusClearEffects(ByVal busID As RiffBusId, Optional ByVal clearPersistent As Boolean = True)
Public Property Get RiffBusPresetEnabled(ByVal busID As RiffBusId) As Boolean
Public Property Let RiffBusPresetEnabled(ByVal busID As RiffBusId, ByVal value As Boolean)
Public Property Get RiffBusPreset(ByVal busID As RiffBusId) As RiffEffectPreset
Public Property Get RiffBusPresetAmount(ByVal busID As RiffBusId) As Single
```

### Assets and Export

```vb
Public Function RiffLoad(ByVal filePath As String) As Long
Public Function RiffLoadFromMemory(ByRef audioData() As Byte) As Long
Public Sub RiffUnload(ByVal bufferHandle As Long)
Public Property Get RiffBufferDurationSec(ByVal bufferHandle As Long) As Single
Public Function RiffExportBufferWav(ByVal bufferHandle As Long, ByVal filePath As String) As Boolean
Public Function RiffRenderOscillatorWav(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ByVal durationSec As Single, ByVal filePath As String) As Boolean
```

### Playback

```vb
Public Function RiffPlay(ByVal bufferHandle As Long, Optional ByVal busID As RiffBusId = RiffBusMain, Optional ByVal looped As Boolean = False, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!) As Long
Public Function RiffPlayOnce(ByVal bufferHandle As Long, Optional ByVal busID As RiffBusId = RiffBusMain, Optional ByVal looped As Boolean = False, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!) As Long
Public Function RiffPlayOscillator(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, Optional ByVal busID As RiffBusId = RiffBusMain, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!, Optional ByVal durationSec As Single = 0!) As Long
Public Function RiffPlayNoise(Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, Optional ByVal busID As RiffBusId = RiffBusMain, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!, Optional ByVal durationSec As Single = RIFF_DEFAULT_NOISE_DURATION_SEC) As Long
Public Function RiffPlayNoiseOnBus(ByVal busID As RiffBusId, Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!, Optional ByVal durationSec As Single = RIFF_DEFAULT_NOISE_DURATION_SEC) As Long
Public Function RiffPlayNoiseLoop(Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, Optional ByVal busID As RiffBusId = RiffBusMain, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!) As Long
Public Function RiffPlayNoiseLoopOnBus(ByVal busID As RiffBusId, Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!) As Long
```

### Compatibility Playback Wrappers

```vb
Public Function RiffPlayBus(ByVal bufferHandle As Long, ByVal busID As RiffBusId, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!) As Long
Public Function RiffPlayBusOnce(ByVal bufferHandle As Long, ByVal busID As RiffBusId, Optional ByVal looped As Boolean = False, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!) As Long
Public Function RiffPlayOscillatorBus(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ByVal busID As RiffBusId, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!, Optional ByVal durationSec As Single = 0!) As Long
Public Function RiffPlayNoiseBus(Optional ByVal noiseType As RiffWaveType = RiffWaveWhiteNoise, Optional ByVal busID As RiffBusId = RiffBusMain, Optional ByVal volume As Single = RIFF_DEFAULT_VOICE_VOLUME, Optional ByVal pan As Single = 0!, Optional ByVal durationSec As Single = RIFF_DEFAULT_NOISE_DURATION_SEC) As Long
```

### Voice Transport and State

```vb
Public Sub RiffPause(ByVal voiceHandle As Long)
Public Sub RiffResume(ByVal voiceHandle As Long)
Public Sub RiffStop(ByVal voiceHandle As Long)
Public Sub RiffVoiceStop(ByVal voiceHandle As Long)
Public Sub RiffStopAll()
Public Sub RiffFadeIn(ByVal voiceHandle As Long, ByVal durationSec As Single)
Public Sub RiffFadeOut(ByVal voiceHandle As Long, ByVal durationSec As Single)
Public Sub RiffSetLoopRegionSec(ByVal voiceHandle As Long, ByVal startSec As Single, ByVal endSec As Single)
Public Property Get RiffVoiceIsPlaying(ByVal voiceHandle As Long) As Boolean
Public Function RiffVoicePlaying(ByVal voiceHandle As Long) As Boolean
Public Function RiffVoiceActive(ByVal voiceHandle As Long) As Boolean
Public Property Get RiffVoiceIsPaused(ByVal voiceHandle As Long) As Boolean
Public Function RiffFindPlayingVoice(ByVal bufferHandle As Long, Optional ByVal busID As Long = -1) As Long
Public Function RiffBufferIsPlaying(ByVal bufferHandle As Long, Optional ByVal busID As Long = -1) As Boolean
```

### Voice Routing, Metering, and Basic Properties

```vb
Public Property Get RiffVoiceBus(ByVal voiceHandle As Long) As RiffBusId
Public Property Let RiffVoiceBus(ByVal voiceHandle As Long, ByVal value As RiffBusId)
Public Sub RiffVoiceGetPeak(ByVal voiceHandle As Long, ByRef peakLeft As Single, ByRef peakRight As Single)
Public Property Get RiffVoiceLoop(ByVal voiceHandle As Long) As Boolean
Public Property Let RiffVoiceLoop(ByVal voiceHandle As Long, ByVal value As Boolean)
Public Property Get RiffVoicePositionSec(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoicePositionSec(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceVolume(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceVolume(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoicePitch(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoicePitch(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoicePan(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoicePan(ByVal voiceHandle As Long, ByVal value As Single)
```

### Effect Helpers and Presets

```vb
Public Sub RiffVoiceClearEffects(ByVal voiceHandle As Long)
Public Sub RiffVoiceSetReverb(ByVal voiceHandle As Long, ByVal mix As Single, ByVal roomTime As Single)
Public Sub RiffVoiceSetDelay(ByVal voiceHandle As Long, ByVal delayTime As Single, ByVal feedback As Single, ByVal mix As Single)
Public Sub RiffVoiceSetChorus(ByVal voiceHandle As Long, ByVal depth As Single, ByVal rate As Single)
Public Sub RiffVoiceSetFlanger(ByVal voiceHandle As Long, ByVal depth As Single, ByVal rate As Single, ByVal feedback As Single)
Public Sub RiffVoiceSetFilter(ByVal voiceHandle As Long, ByVal lowPass As Single, ByVal highPass As Single)
Public Sub RiffVoiceSetFilterHz(ByVal voiceHandle As Long, ByVal lowPassHz As Single, Optional ByVal highPassHz As Single = 0!)
Public Sub RiffVoiceApplyPreset(ByVal voiceHandle As Long, ByVal preset As RiffEffectPreset, Optional ByVal amount As Single = 1!)
```

### DSP Properties

```vb
Public Property Get RiffVoiceBitDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceBitDepth(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceSampleRateReduction(ByVal voiceHandle As Long) As Long
Public Property Let RiffVoiceSampleRateReduction(ByVal voiceHandle As Long, ByVal value As Long)
Public Property Get RiffVoiceRingModFreq(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceRingModFreq(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceRingModMix(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceRingModMix(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceAutoPanRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceAutoPanRate(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceAutoPanDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceAutoPanDepth(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceEqBass(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceEqBass(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceEqMid(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceEqMid(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceEqTreble(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceEqTreble(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceCompressorThreshold(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceCompressorThreshold(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceCompressorRatio(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceCompressorRatio(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceFlangerDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceFlangerDepth(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceFlangerRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceFlangerRate(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceFlangerFeedback(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceFlangerFeedback(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceDistortion(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDistortion(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceLowPass(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceLowPass(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceHighPass(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceHighPass(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceStereoWidth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceStereoWidth(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceTremoloRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceTremoloRate(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceTremoloDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceTremoloDepth(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceChorusDepth(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceChorusDepth(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceChorusRate(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceChorusRate(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceReverbMix(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceReverbMix(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceReverbTime(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceReverbTime(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceDelayTime(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDelayTime(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceDelayFeedback(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDelayFeedback(ByVal voiceHandle As Long, ByVal value As Single)
Public Property Get RiffVoiceDelayMix(ByVal voiceHandle As Long) As Single
Public Property Let RiffVoiceDelayMix(ByVal voiceHandle As Long, ByVal value As Single)
```
