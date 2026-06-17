<div align="center">
  <img src="resources/logo.png" width="160" alt="Riff logo" />
</div>

<h1 align="center">Riff</h1>

<p align="center">
  <b>A high-performance, single-file audio engine for Microsoft Office.</b><br/>
  Real-time WASAPI playback, Media Foundation decoding, Studio DSP, asset-key playback, cooperative preload, scene templates, adaptive buffering, dynamic SafeArray pools, automatic overload protection, burst-safe playback, musical presets, limiter-safe master processing, editor-safe development guards, and VBE-safe cleanup.
</p>

<p align="center">
  <img src="https://github.com/vbacollective/riff/actions/workflows/ci.yml/badge.svg" alt="CI" />
  <img src="https://github.com/vbacollective/riff/actions/workflows/release-assets.yml/badge.svg" alt="Release" />
  <img src="https://img.shields.io/github/v/release/vbacollective/riff" alt="Latest Version" />
  <img src="https://img.shields.io/badge/language-VBA-867DB1.svg" alt="Language" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D6.svg" alt="Platform" />
  <img src="https://img.shields.io/badge/arch-32%20%26%2064--bit-green.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/WASAPI-Shared%20Mode-blue.svg" alt="WASAPI" />
  <img src="https://img.shields.io/badge/Media%20Foundation-Decoding-orange.svg" alt="Media Foundation" />
  <img src="https://img.shields.io/badge/DSP-Studio%20Pipeline-critical.svg" alt="DSP" />
  <img src="https://img.shields.io/badge/dependencies-none-success.svg" alt="Dependencies" />
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" />
</p>

**Riff** is a complete, production-grade audio engine contained within a single `.bas` module. It allows VBA developers to integrate professional-quality audio playback, synthesis, routing, DSP, effect presets, and final mix processing into Microsoft Office applications without external DLLs, ActiveX controls, installers, or additional references.

Whether you are building interactive dashboards in Excel, immersive presentations in PowerPoint, educational tools in Word, or automation systems in Access, Riff provides a practical real-time audio layer powered by the Windows Audio Session API (WASAPI), Media Foundation, and a native timer-driven DSP loop.

Riff is designed for developers who want game-like audio behavior inside Office: responsive UI sounds, background music, routed buses, named asset playback, cooperative loading screens, one-call scene templates, persistent scene effects, master bus processing, limiter/headroom safety, fades, procedural oscillators, one-shot and looped white/pink/brown noise, dynamic buffer/voice capacity, automatic burst coalescing, fast preset setup, and safer cleanup when the host application or the VBA editor resets.


## Key Capabilities

- **Zero Dependencies:** No DLLs to ship or register. Import `Riff.bas` and run.
- **Single-File Distribution:** The entire engine lives in one VBA module.
- **WASAPI Playback:** Shared-mode Windows audio output with low-latency rendering.
- **Media Foundation Decoding:** Loads common formats supported by the system, including WAV, MP3, AAC, FLAC, WMA, and more.
- **WAV Fast Path:** Compatible WAV files bypass Media Foundation and load through a direct RIFF parser for much faster startup.
- **Unified Playback API:** `RiffPlay` accepts bus, loop, volume, and pan parameters directly.
- **Asset Registry:** `RiffLoadAs`, `RiffPlayKey`, `RiffPlayKeyOnce`, `RiffPlayKeyLoop`, `RiffStopKey`, and `RiffFadeOutKey` let projects play audio by stable string keys instead of manually passing handles everywhere.
- **Cooperative Preload:** `RiffPreloadAdd`, `RiffPreloadStart`, and `RiffPreloadUpdate` support loading-screen style asset preparation without forcing a giant one-frame load spike.
- **Duplicate Prevention:** `RiffPlayOnce` and `RiffPlayKeyOnce` prevent music or repeated ambience from stacking accidentally.
- **Scene Templates:** `RiffApplyScene` provides ready-made mixer states for battle, night, cave, underwater, radio call, dream, horror, retro, cinematic, pause menu, and normal playback.
- **Dynamic SafeArray Pools:** Buffers, voices, and bus state start with practical defaults and grow automatically instead of being locked to a hardcoded ceiling.
- **Capacity Reservation:** `RiffReserveBuffers`, `RiffReserveVoices`, and `RiffReserveBuses` let large projects preallocate room before loading or playing many sounds.
- **Adaptive Buffering:** Dynamically increases render queue safety during host stalls and returns to low latency when stable.
- **Automatic Overload Protection:** The engine detects extreme play-call spam internally and coalesces/reuses voices without requiring the user to call manual begin/end burst functions.
- **Burst Protection:** Voice stealing, active-voice budgets, per-buffer caps, per-bus caps, and loop coalescing reduce stutter when many SFX are triggered rapidly.
- **Stress-Safe Defaults:** `RiffApplyStressSafeDefaults` applies practical voice budgets, overload protection, limiter/headroom safety, and clipping protection for SFX-heavy projects.
- **Anti-Accumulation Playback:** Repeated short sounds, procedural noise, and oscillator beeps are capped, stolen, coalesced, or cleaned so game loops do not silently pile up voices.
- **One-Shot Procedural Noise:** `RiffPlayNoise` is finite by default; use `RiffPlayNoiseLoop` for continuous ambience.
- **Finite Oscillator Beeps:** `RiffPlayOscillator` accepts `durationSec` for procedural UI/game SFX that clean themselves automatically.
- **Hz-Based Filtering:** `RiffVoiceSetFilterHz` exposes low-pass and high-pass filters using real frequency values.
- **Lazy DSP Buffer Preparation:** Delay, reverb, chorus, and flanger ring buffers are prepared only when they are actually needed, reducing `Play` and preset setup cost.
- **VBE-Safe Timer Cleanup:** Idle and Stop/Reset-safe cleanup reduce the chance of the VBA Editor staying stuck in `Running` mode or breaking IntelliSense.
- **Editor-Safe Development Guards:** `RiffPrepareForVbeEdit` and `RiffResumeAfterVbeEdit` provide a controlled way to pause the risky timer/callback path before editing VBA code while audio is active.
- **Studio DSP Pipeline:** Independent per-voice effects including Reverb, Delay, Chorus, Flanger, Compressor, EQ, Filters, Distortion, Bitcrusher, Ring Modulation, Tremolo, Auto-Pan, and Stereo Width.
- **Musical Preset Packs:** Expanded voice presets for tape, cassette, VHS, old computer, arcade cabinets, dream menus, caves, dungeons, boss rooms, space stations, low health, memory flashbacks, cutscenes, combat impacts, tiny speakers, megaphones, retro game effects, wind, rain, horror drones, cinematic booms, and soft-focus scenes.
- **Persistent Bus Effects:** Apply a preset to a whole bus so current and future voices inherit the scene style automatically.
- **Master Bus Processors:** Final mix chain with low-pass, high-pass, 3-band EQ, compressor, drive, stereo width, output gain, soft clipping, limiter controls, balance, tilt EQ, master presets, and master meters.
- **Real-Time Synthesis:** BLEP-corrected oscillators for sine, square, and saw waveforms.
- **Noise Generation:** White, pink, and brown noise for procedural ambience, wind, rain, static, rumble, and retro effects.
- **Audio Routing:** Named buses for Music, SFX, UI, Voice, and auxiliary groups, backed by dynamically sized bus state.
- **Mixer Controls:** Bus volume, mute, solo, fades, peak meters, RMS meters, clip counters, and master peak/RMS monitoring.
- **Smoothing:** Volume, pan, and pitch smoothing reduce clicks during parameter changes.
- **Soft Clipping and Limiting:** Master soft clipper and limiter help prevent harsh digital clipping when many voices overlap.
- **Diagnostics:** Render counters, underrun counters, overload counters, clipping counters, buffer status, active voice counts, RMS/peak meters, and adaptive queue information.
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
        Debug.Print "Playback failed. Error:", RiffLastError
    End If
End Sub
```


### Modern Asset-Key Workflow

For bigger projects, prefer the v1.1.3 asset registry. It lets you load sounds with names and play them later without keeping a global variable for every buffer handle.

```vb
Public Sub AudioLoadKeys()
    If Not RiffOpen() Then Exit Sub

    RiffLoadAs "ui.click", ActivePresentation.Path & "\audio\click.wav"
    RiffLoadAs "sfx.hit", ActivePresentation.Path & "\audio\hit.wav"
    RiffLoadAs "music.menu", ActivePresentation.Path & "\audio\menu.mp3"
End Sub

Public Sub PlayMenuClick()
    RiffPlayKey "ui.click", RiffBusUi, False, 0.75!, 0!
End Sub

Public Sub PlayMenuMusic()
    RiffPlayKeyOnce "music.menu", RiffBusMusic, True, 0.45!, 0!
End Sub
```

This is especially useful in PowerPoint games and large Excel tools where asset names are easier to manage than dozens of `Long` variables.

### Stress-Safe Game Setup

For SFX-heavy projects, v1.1.3 provides one setup call that enables practical overload protection, limiter/headroom safety, and voice budgets.

```vb
Public Sub AudioInitForHeavyGameplay()
    If Not RiffOpen() Then Exit Sub

    RiffReserveBuffers 500
    RiffReserveVoices 160
    RiffReserveBuses 32

    ' max active voices, max voices per buffer, max voices per bus
    RiffApplyStressSafeDefaults 64, 6, 32
End Sub
```

You still call `RiffPlay`, `RiffPlayKey`, `RiffPlayNoise`, and `RiffPlayOscillator` normally. The overload/coalescing behavior is internal.

### Optional Capacity Reservation

Riff starts with practical default capacity for compatibility, but the internal buffer, voice, and bus pools are no longer fixed to the old hardcoded limits. When more room is needed, the engine can grow its internal `SafeArray` pools automatically.

For small projects, you do not need to do anything special:

```vb
click = RiffLoad("C:\Audio\click.wav")
voice = RiffPlay(click, RiffBusUi)
```

For larger games, presentations, or tools with many assets, you can reserve capacity during startup. This avoids repeated pool growth while the project is already busy loading or playing sounds.

```vb
Public Sub AudioReserveForLargeProject()
    If Not RiffOpen() Then Exit Sub

    RiffReserveBuffers 500
    RiffReserveVoices 128
    RiffReserveBuses 32
End Sub
```

The diagnostic capacity properties now report the current pool size instead of a permanent hard limit:

```vb
Debug.Print "Buffer capacity:", RiffMaxBuffers
Debug.Print "Voice capacity:", RiffMaxVoices
Debug.Print "Bus capacity:", RiffMaxBuses
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

If you are actively developing in the VBA Editor and intentionally hit the **Stop/Reset** button during a test, the current stop-safe build is designed to kill orphaned timer callbacks automatically. If an old test build ever leaves the editor in a stuck state, call:

```vb
RiffEditorEmergencyStop
```

This is an editor recovery helper, not normal application shutdown. Production code should still use `RiffClose`.

When you need to edit VBA code while Riff audio is active, use the manual editor-safe workflow. This pauses the dangerous timer/callback path before the VBE recompiles or rearranges project memory, then resumes the engine after editing.

```vb
Public Sub EnterSafeEditMode()
    RiffPrepareForVbeEdit
End Sub

Public Sub LeaveSafeEditMode()
    RiffResumeAfterVbeEdit
End Sub
```

`RiffEditorSafeMode` also exists for automatic experiments, but the recommended development workflow is still manual because automatic editor detection can pause audio when you are simply testing from the VBE.

## Basic Usage

### Load Assets Once

Audio loading is synchronous. Load your assets during startup, not during gameplay, animation ticks, or button-spam interactions.

```vb
Private sndClick As Long
Private sndExplosion As Long
Private sndMusic As Long

Public Sub AudioLoad()
    If Not RiffOpen() Then Exit Sub

    ' Optional for larger projects: reserve before loading many assets.
    RiffReserveBuffers 256
    RiffReserveVoices 96

    sndClick = RiffLoad(ActivePresentation.Path & "\audio\click.wav")
    sndExplosion = RiffLoad(ActivePresentation.Path & "\audio\explosion.wav")
    sndMusic = RiffLoad(ActivePresentation.Path & "\audio\music.wav")

    RiffBusVolume(RiffBusUi) = 0.9
    RiffBusVolume(RiffBusSfx) = 0.85
    RiffBusVolume(RiffBusMusic) = 0.45
End Sub
```


### Load Assets by Key

The asset registry keeps the code cleaner when a project has many sounds.

```vb
Public Sub AudioLoadRegistry()
    If Not RiffOpen() Then Exit Sub

    RiffLoadAs "ui.click", ActivePresentation.Path & "\audio\ui_click.wav"
    RiffLoadAs "ui.cancel", ActivePresentation.Path & "\audio\ui_cancel.wav"
    RiffLoadAs "sfx.explosion", ActivePresentation.Path & "\audio\explosion.wav"
    RiffLoadAs "music.menu", ActivePresentation.Path & "\audio\music.wav"
End Sub

Public Sub PlayRegistryClick()
    RiffPlayKey "ui.click", RiffBusUi, False, 0.7!, 0!
End Sub

Public Sub EnsureRegistryMusic()
    RiffPlayKeyOnce "music.menu", RiffBusMusic, True, 0.5!, 0!
End Sub
```

Useful helpers:

```vb
Debug.Print RiffAssetExists("ui.click")
Debug.Print RiffAssetHandle("ui.click")
Debug.Print RiffAssetPath("ui.click")

RiffStopKey "music.menu"
RiffFadeOutKey "music.menu", 0.75
```

### Cooperative Preload

For a loading screen, queue assets and advance loading over multiple calls.

```vb
Public Sub BuildLoadingQueue()
    RiffPreloadClear
    RiffPreloadAdd "ui.click", ActivePresentation.Path & "\audio\ui_click.wav"
    RiffPreloadAdd "sfx.explosion", ActivePresentation.Path & "\audio\explosion.wav"
    RiffPreloadAdd "music.menu", ActivePresentation.Path & "\audio\music.wav"
    RiffPreloadStart
End Sub

Public Sub LoadingTick()
    RiffPreloadUpdate

    Debug.Print "Loading:", Format$(RiffPreloadProgress * 100!, "0") & "%"

    If RiffPreloadFinished Then
        Debug.Print "Loaded:", RiffPreloadLoadedCount
        Debug.Print "Failed:", RiffPreloadFailedCount
    End If
End Sub
```

Use `RiffPreloadUpdateAll` only when you intentionally want to finish the whole queue immediately.

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
        RiffFadeOut musicVoice, 0.5
    End If
End Sub
```

### Generate a Procedural Sound

Use `durationSec` for short procedural sounds. This avoids leaving a continuous oscillator running by accident.

```vb
Public Sub PlayBeep()
    RiffPlayOscillator RiffWaveSine, 880, RiffBusUi, 0.25, 0, 0.08
End Sub
```

For continuous synthesis, pass `durationSec:=0` or omit it, then stop or fade the returned voice manually.

```vb
Public Sub StartLowHum()
    Dim hum As Long
    hum = RiffPlayOscillator(RiffWaveSine, 55, RiffBusSfx, 0.12, 0)
    RiffFadeOut hum, 1.5
End Sub
```

### Generate Noise

`RiffPlayNoise` is a short one-shot by default. This is safer for game SFX because it prevents accidental infinite procedural voices.

```vb
Public Sub PlayDustHit()
    Dim v As Long
    v = RiffPlayNoise(RiffWaveWhiteNoise, RiffBusSfx, 0.08, 0, 0.05)

    If v >= 0 Then
        RiffVoiceSetFilterHz v, 2200, 180
    End If
End Sub
```

For continuous rain, wind, ambience, or drones, use `RiffPlayNoiseLoop`.

```vb
Public Sub PlayRainLayer()
    Dim v As Long
    v = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.08, 0)

    If v >= 0 Then
        RiffVoiceApplyPreset v, RiffFxRain, 0.65
        RiffVoiceStereoWidth(v) = 1.35
    End If
End Sub
```

## Professional Audio Routing

Riff supports logical summing through named audio buses backed by dynamically sized bus state. Buses allow you to group related sounds, such as Music, SFX, UI, Voice, or Ambience, and control them as a single unit.

This is more efficient and cleaner than iterating through active voices manually. It also enables mixer-like behavior such as lowering music during dialogue, muting UI sounds, soloing a bus for debugging, fading whole categories, or applying scene-wide effects.

### Signal Hierarchy

The final volume of a sound is determined by:

```text
Master Volume × Bus Volume × Voice Volume × Fade/Smoothing × Master Processing
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


## Scene Templates

v1.1.3 includes built-in scene templates. These are higher-level mixer recipes that adjust buses, persistent effects, and master color in one call.

```vb
Public Sub EnterBattle()
    RiffApplyScene RiffSceneBattle
End Sub

Public Sub EnterCave()
    RiffApplyScene RiffSceneCave
End Sub

Public Sub EnterRadioCall()
    RiffApplyScene RiffSceneRadioCall
End Sub

Public Sub LeaveSpecialScene()
    RiffApplyScene RiffSceneNormal
End Sub
```

Available scenes:

```vb
RiffSceneNormal
RiffScenePauseMenu
RiffSceneBattle
RiffSceneNight
RiffSceneCave
RiffSceneUnderwater
RiffSceneRadioCall
RiffSceneDream
RiffSceneHorror
RiffSceneRetro
RiffSceneCinematic
```

Scenes work best when audio is routed to meaningful buses:

```vb
RiffPlayKey "music.battle", RiffBusMusic, True, 0.45!, 0!
RiffPlayKey "sfx.hit", RiffBusSfx, False, 0.85!, 0!
RiffPlayKey "ui.confirm", RiffBusUi, False, 0.65!, 0!
RiffPlayKey "voice.npc", RiffBusVoice, False, 1!, 0!
```

If everything is played on `RiffBusMain`, scene templates have less useful routing information to work with.


## Unified Playback API

The recommended public API is intentionally compact. Handle-based playback remains available, and v1.1.3 adds key-based playback for larger projects.

```vb
voice = RiffPlay(bufferHandle, busID, looped, volume, pan)
voice = RiffPlayOnce(bufferHandle, busID, looped, volume, pan)
voice = RiffPlayKey("sfx.hit", busID, looped, volume, pan)
voice = RiffPlayKeyOnce("music.menu", busID, looped, volume, pan)
voice = RiffPlayKeyLoop("amb.wind", busID, volume, pan)
voice = RiffPlayOscillator(waveType, frequencyHz, busID, volume, pan, durationSec)
voice = RiffPlayNoise(noiseType, busID, volume, pan, durationSec)
voice = RiffPlayNoiseLoop(noiseType, busID, volume, pan)
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

### `RiffPlayKey`

Use this when the sound was loaded with `RiffLoadAs` or registered through the asset registry.

```vb
RiffLoadAs "sfx.explosion", ActivePresentation.Path & "\audio\explosion.wav"
RiffPlayKey "sfx.explosion", RiffBusSfx, False, 0.9!, 0!
```

### `RiffPlayKeyOnce` and `RiffPlayKeyLoop`

Use these for music and ambience that should not stack.

```vb
RiffLoadAs "music.menu", ActivePresentation.Path & "\audio\menu.mp3"
RiffPlayKeyOnce "music.menu", RiffBusMusic, True, 0.45!, 0!

RiffLoadAs "amb.wind", ActivePresentation.Path & "\audio\wind.wav"
RiffPlayKeyLoop "amb.wind", RiffBusMusic, 0.25!, 0!
```

### `RiffStopKey` and `RiffFadeOutKey`

```vb
RiffFadeOutKey "music.menu", 0.8
RiffStopKey "amb.wind"
```

### `RiffPlayOscillator`

`frequencyHz` controls pitch. `durationSec` is optional: `0` means continuous, while a positive value creates a finite procedural one-shot.

```vb
' Short retro UI beep.
RiffPlayOscillator RiffWaveSquare, 880, RiffBusUi, 0.2, 0, 0.07

' Continuous oscillator layer.
Dim osc As Long
osc = RiffPlayOscillator(RiffWaveSine, 55, RiffBusSfx, 0.12, 0)
RiffFadeOut osc, 0.4
```

### `RiffPlayNoise`

`RiffPlayNoise` is finite by default. Use it for dust hits, static bursts, impacts, wind puffs, and other short procedural SFX.

```vb
Dim dust As Long
dust = RiffPlayNoise(RiffWaveWhiteNoise, RiffBusSfx, 0.12, 0, 0.04)
RiffVoiceSetFilterHz dust, 2600, 180
```

For ambience, use the explicit loop helper:

```vb
Dim wind As Long
wind = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.05, 0)
RiffVoiceSetFilterHz wind, 1800, 80
```

Bus-first helpers are also available when readability matters:

```vb
RiffPlayNoiseOnBus RiffBusSfx, RiffWaveWhiteNoise, 0.08, 0, 0.03
RiffPlayNoiseLoopOnBus RiffBusMusic, RiffWavePinkNoise, 0.04, 0
```

### Compatibility Wrappers

Older function names are kept for compatibility:

```vb
RiffPlayBus bufferHandle, busID
RiffPlayBusOnce bufferHandle, busID, looped
RiffPlayOscillatorBus waveType, frequencyHz, busID, volume, pan, durationSec
RiffPlayNoiseBus noiseType, busID, volume, pan, durationSec
RiffPlayNoiseOnBus busID, noiseType, volume, pan, durationSec
RiffPlayNoiseLoop noiseType, busID, volume, pan
RiffPlayNoiseLoopOnBus busID, noiseType, volume, pan
```

They forward internally to the unified playback path.


## Effect Presets

Presets provide fast, musical starting points for common sound design situations.

```vb
RiffVoiceApplyPreset voice, RiffFxLoFi, 0.55
RiffVoiceApplyPreset voice, RiffFxRadio, 0.6
RiffVoiceApplyPreset voice, RiffFxUnderwater, 0.7
RiffVoiceApplyPreset voice, RiffFxWarmTape, 0.5
```

### Core Presets

| Preset | Use Case |
|:---|:---|
| `RiffFxDry` | Clears effect-heavy coloration and returns toward a clean signal. |
| `RiffFxSmallRoom` | Subtle room reflections for close spaces. |
| `RiffFxHall` | Wider ambience for music or narration. |
| `RiffFxCathedral` | Large washed-out reverb tail. |
| `RiffFxEcho` | Musical delay/echo effect. |
| `RiffFxLoFi` | Tape/sampler-style degradation without destroying the sound. |
| `RiffFxRadio` | Band-limited radio or speaker effect. |
| `RiffFxUnderwater` | Muffled filtered sound with movement. |
| `RiffFxRobot` | Ring-modulated robotic coloration. |
| `RiffFxWide` | Enhanced stereo width. |
| `RiffFxAmbient` | Spacious ambience for beds and pads. |

### Musical Preset Packs

| Preset | Use Case |
|:---|:---|
| `RiffFxWarmTape` | Warm tape-like coloration for music, ambience, or narration. |
| `RiffFxVHS` | Warbly degraded old-media sound. |
| `RiffFxDreamPad` | Wide, soft chorus/reverb for dream scenes and pads. |
| `RiffFxDarkCave` | Dark, deep, cave-like space. |
| `RiffFxTinySpeaker` | Phone, toy speaker, laptop, or small radio tone. |
| `RiffFxMegaphone` | PA, announcement, or projected voice tone. |
| `RiffFxGameBoy` | Crunchy retro handheld game color. |
| `RiffFxHorrorDrone` | Dark modulated unsettling texture. |
| `RiffFxWind` | Airy filtered wind/noise treatment. |
| `RiffFxRain` | Soft natural noise ambience. |
| `RiffFxCinematicBoom` | Big low-heavy impact treatment. |
| `RiffFxSoftFocus` | Gentle smoothing and width. |

### v1.1.3 Expanded Presets

| Preset | Use Case |
|:---|:---|
| `RiffFxCassette` | Cassette-style warmth, dullness, and soft degradation. |
| `RiffFxOldComputer` | Crunchy old PC / multimedia speaker character. |
| `RiffFxArcadeCabinet` | Bright, narrow, arcade-machine style UI and SFX tone. |
| `RiffFxDreamMenu` | Soft menu ambience with a dreamlike polished tone. |
| `RiffFxSpaceStation` | Cold sci-fi room tone for machinery, ambience, and UI. |
| `RiffFxDungeon` | Darker reverb and low-mid weight for RPG rooms. |
| `RiffFxBossRoom` | Stronger, heavier scene color for combat areas. |
| `RiffFxLowHealth` | Tense filtered state for low HP, warning, or danger moments. |
| `RiffFxMemoryFlashback` | Soft degraded recollection / flashback coloration. |
| `RiffFxCutscene` | Polished cinematic treatment for narrative moments. |
| `RiffFxCombatImpact` | Punchier transient-heavy effect for hits and impacts. |

### Preset Amount

`amount` usually ranges from `0.0` to `1.0`.

```vb
RiffVoiceApplyPreset v, RiffFxWarmTape, 0.3  ' light coloration
RiffVoiceApplyPreset v, RiffFxWarmTape, 0.7  ' stronger effect
RiffVoiceApplyPreset v, RiffFxDry, 1.0       ' clear preset-style coloration
```

The current performance build sanitizes preset values and prepares temporal DSP buffers lazily. Presets that do not actually use delay, reverb, chorus, or flanger no longer pay the cost of clearing large ring buffers during the `RiffVoiceApplyPreset` call.

## Persistent Bus Effects

Riff can apply voice presets to a whole bus. This is useful for scene-wide states such as underwater, cave, radio, dream, horror, or retro menus.

By default, `RiffBusApplyPreset` affects currently active voices and stores the preset for future voices routed to that bus.

```vb
Public Sub EnterUnderwaterScene()
    RiffBusApplyPreset RiffBusMusic, RiffFxUnderwater, 0.55
    RiffBusApplyPreset RiffBusSfx, RiffFxUnderwater, 0.8
    RiffBusApplyPreset RiffBusVoice, RiffFxUnderwater, 0.45
End Sub

Public Sub LeaveUnderwaterScene()
    RiffBusClearEffects RiffBusMusic
    RiffBusClearEffects RiffBusSfx
    RiffBusClearEffects RiffBusVoice
End Sub
```

Apply only to future voices:

```vb
RiffBusApplyPreset RiffBusVoice, RiffFxRadio, 0.65, True, False
```

Apply only to currently active voices:

```vb
RiffBusApplyPreset RiffBusSfx, RiffFxSmallRoom, 0.4, False, True
```

Inspect bus preset state:

```vb
Debug.Print RiffBusPresetEnabled(RiffBusMusic)
Debug.Print RiffBusPreset(RiffBusMusic)
Debug.Print RiffBusPresetAmount(RiffBusMusic)
```


## Master Bus Processors

Master processors run after the full voice/bus mix. They are intended for final polish, safety limiting, broad scene coloration, diagnostics, and loud SFX protection.

```vb
RiffMasterApplyPreset RiffMasterFxGlue, 0.7
RiffMasterApplyPreset RiffMasterFxCinematic, 0.6
```

### Master Presets

| Preset | Use Case |
|:---|:---|
| `RiffMasterFxClean` | Neutral master stage. |
| `RiffMasterFxGlue` | Light compression and soft limiting for a cohesive mix. |
| `RiffMasterFxWarm` | Warmer tone and subtle saturation. |
| `RiffMasterFxBright` | Brighter overall mix. |
| `RiffMasterFxDark` | Darker and softer output. |
| `RiffMasterFxRadio` | Global radio/band-limited sound. |
| `RiffMasterFxCinematic` | Wider, fuller, slightly compressed cinematic shaping. |
| `RiffMasterFxNight` | Softer, lower-energy night mix. |
| `RiffMasterFxSoftLimiter` | Safety limiting for SFX-heavy scenes. |

### Limiter and Headroom Safety

For scenes where many one-shots can overlap, enable the master processor, limiter, output headroom, and soft clipper.

```vb
Public Sub ApplySafeMasterForGameplay()
    RiffMasterProcessorEnabled = True
    RiffMasterSetLimiter 0.92!, 0.9!
    RiffMasterOutputGain = 0.78!
    RiffSoftClipEnabled = True
End Sub
```

`RiffApplyStressSafeDefaults` applies this kind of setup automatically with practical voice budgets.

```vb
RiffApplyStressSafeDefaults 64, 6, 32
```

### Manual Master Chain

```vb
Public Sub ApplyManualMasterChain()
    RiffMasterProcessorEnabled = True

    RiffMasterLowPass = 0.92
    RiffMasterHighPass = 0.02

    RiffMasterEqBass = 1.08
    RiffMasterEqMid = 1
    RiffMasterEqTreble = 0.95

    RiffMasterCompressorThreshold = 0.72
    RiffMasterCompressorRatio = 2.5

    RiffMasterDrive = 1.06
    RiffMasterStereoWidth = 1.1
    RiffMasterOutputGain = 0.96
    RiffSoftClipEnabled = True
End Sub
```

### Tilt EQ and Balance

```vb
' Negative values lean darker; positive values lean brighter.
RiffMasterTiltEq -0.15!

' -1 = left, 0 = center, 1 = right.
RiffMasterBalance 0!
```

### Master Meters

```vb
Public Sub PrintMasterMeters()
    Dim l As Single
    Dim r As Single

    RiffMasterGetRms l, r

    Debug.Print "Master RMS:", l, r
    Debug.Print "Master RMS dB:", RiffMasterRmsDb
    Debug.Print "Master Peak dB:", RiffMasterPeakDb
    Debug.Print "Master Clips:", RiffMasterClipCount
End Sub
```

Bus and voice-level RMS helpers are also available for visualizers, debug overlays, and balancing tools.

```vb
Dim l As Single
Dim r As Single

RiffBusGetRms RiffBusMusic, l, r
RiffVoiceGetRms musicVoice, l, r
```

Clear master processing:

```vb
RiffMasterClearProcessors
RiffMasterApplyPreset RiffMasterFxClean
```

## Manual DSP Helpers

You can use the individual properties directly, or use helper functions to set common groups of parameters.

```vb
RiffVoiceSetReverb v, 0.35, 0.7
RiffVoiceSetDelay v, 0.28, 0.45, 0.5
RiffVoiceSetChorus v, 0.5, 1.2
RiffVoiceSetFlanger v, 0.6, 0.35, 0.4
RiffVoiceSetFilter v, 0.45, 0.05
RiffVoiceSetFilterHz v, 3000, 300
RiffVoiceClearEffects v
```

### Smooth Voice Changes

Use smoothing helpers to avoid clicks and sudden jumps.

```vb
RiffVoiceVolumeTo musicVoice, 0.2, 500
RiffVoicePanTo voice, -0.5, 150
RiffVoicePitchTo voice, 1.2, 100
```

### Frequency-Based Filters

Use `RiffVoiceSetFilterHz` when you want sound-design values that map to real audio frequencies.

```vb
' Telephone/radio band.
RiffVoiceSetFilterHz voice, 3000, 300

' Muffled wall or underwater style.
RiffVoiceSetFilterHz voice, 900, 0

' Remove sub-rumble while keeping the top end open.
RiffVoiceSetFilterHz voice, 0, 80
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
' A4 sine beep for 80 ms.
RiffPlayOscillator RiffWaveSine, 440, RiffBusSfx, 0.2, 0, 0.08

' Short pink-noise hit for 60 ms.
RiffPlayNoise RiffWavePinkNoise, RiffBusSfx, 0.08, 0, 0.06

' Continuous ambience.
Dim rain As Long
rain = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.05, 0)
```

### One-Shot vs Loop Behavior

The current stable build treats procedural noise as a one-shot by default. This is intentionally different from early builds, where noise behaved like an infinite generator unless manually stopped.

| Task | Recommended Call |
|:---|:---|
| Dust burst, static tick, hit layer | `RiffPlayNoise(..., durationSec)` |
| Rain, wind, fire, ambience bed | `RiffPlayNoiseLoop(...)` |
| Short beep, coin, UI confirmation | `RiffPlayOscillator(..., durationSec)` |
| Continuous synth/drone | `RiffPlayOscillator(..., durationSec:=0)` |

## WAV Fast Path

`RiffLoad` attempts a direct WAV fast path before falling back to Media Foundation.

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

Riff includes adaptive buffering to reduce audible dropouts:

```text
Normal state:
    target queue stays low for responsive playback

Stall or underrun risk:
    target queue rises automatically for safety

Stable state:
    target queue gradually returns to low latency
```

The performance build can keep the endpoint warm with silence during rapid SFX bursts so the WASAPI buffer does not drain to zero between short sounds. The editor-safe build also auto-suspends the timer after idle time so the VBA Editor does not remain stuck in `Running` mode.

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

Triggering the same sound many times in a short time can overwhelm any mixer, even with dynamic voice capacity. Riff v1.1.3 protects this path internally.

The important part: the user does **not** need to call manual `Begin` / `End` burst functions. You keep using the normal API.

```vb
RiffPlay sndHit, RiffBusSfx, False, 0.8!, 0!
RiffPlayKey "sfx.hit", RiffBusSfx, False, 0.8!, 0!
RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.08!, 0!, 0.04!
RiffPlayOscillator RiffWaveSquare, 880!, RiffBusUi, 0.15!, 0!, 0.05!
```

Internally, the overload path can:

- coalesce repeated play requests;
- reuse compatible voices during extreme spam;
- steal old non-looping voices before touching important looped audio;
- respect active-voice budgets;
- respect per-buffer and per-bus budgets;
- keep limiter/headroom safety active when stress defaults are enabled.

Recommended heavy-game setup:

```vb
RiffVoiceStealingEnabled = True
RiffOverloadProtectionEnabled = True
RiffLoopCoalescingEnabled = True

RiffMaxActiveVoices = 64
RiffMaxVoicesPerBuffer = 6
RiffMaxVoicesPerBus = 32

RiffMasterProcessorEnabled = True
RiffMasterSetLimiter 0.92!, 0.9!
RiffMasterOutputGain = 0.78!
RiffSoftClipEnabled = True
```

Or use the helper:

```vb
RiffApplyStressSafeDefaults 64, 6, 32
```

### Diagnostics

```vb
Debug.Print "Active voices:", RiffActiveVoiceCount()
Debug.Print "Click instances:", RiffBufferVoiceCount(sndClick, RiffBusUi)
Debug.Print "SFX bus voices:", RiffBusVoiceCount(RiffBusSfx)
Debug.Print "Stolen:", RiffOverloadStolenCount
Debug.Print "Dropped:", RiffOverloadDroppedCount
Debug.Print "Coalesced:", RiffOverloadCoalescedCount
Debug.Print "Clips:", RiffMasterClipCount
```

Reset counters before stress tests:

```vb
RiffResetOverloadStats
RiffResetDiagnostics
RiffResetAdaptiveStats
```

Dynamic voice capacity gives the engine more room, while overload protection keeps gameplay mistakes from turning into hundreds or thousands of redundant real voices.


## Performance and Stability Notes

The v1.1.3 performance pass focuses on making the common gameplay path extremely cheap while keeping the earlier anti-accumulation and editor-safety fixes intact.

Key internal improvements include:

- `RiffPlay` no longer clears large temporal ring buffers for dry one-shot voices.
- `RiffVoiceApplyPreset` no longer eagerly clears delay/reverb/chorus/flanger buffers unless those stages are actually needed.
- Temporal DSP buffer preparation is lazy and tied to the first render tick that needs it.
- Voice allocation combines free-slot search, dynamic pool growth, per-buffer caps, per-bus caps, active voice budgets, and voice-steal candidates in a cheaper path.
- Buffers, voices, and bus state use dynamic `SafeArray` pools instead of being locked to the old fixed capacities.
- Generated noise and oscillator one-shots use finite lifetimes and short release ramps.
- Idle warm-buffer behavior reduces underruns during rapid SFX bursts.
- Stop/Reset-safe editor cleanup kills stale timer callbacks after a VBE reset.
- Automatic overload detection coalesces or reuses repeated play requests during extreme same-tick spam.
- Stress-safe master settings keep limiter/headroom protection active during loud SFX-heavy bursts.

Observed benchmark from the v1.1.3 local performance test:

```text
Short dry RiffPlay burst:       5000 calls, ~22.921 ms total, ~4.584 us/call
Long dry RiffPlay burst:        5000 calls, ~22.234 ms total, ~4.447 us/call
Noise one-shot burst:           5000 calls, ~24.315 ms total, ~4.863 us/call
Oscillator one-shot burst:      5000 calls, ~22.946 ms total, ~4.589 us/call
Preset/DSP setup burst:         2500 calls, ~32.318 ms total, ~12.927 us/call
Gameplay loop:                  3600 frames, 1827 play calls, 0 failed handles
Missed frame budget:            0
Peak active voices:             7
Final active voices:            0
Gameplay underrun delta:        0
Gameplay memory peak-start:     0.000 MB
```

These numbers are not guaranteed across machines, Office versions, or host load, but they describe the current target behavior: very fast one-shot playback, no voice accumulation, stable memory, low underrun counts, and safe handling of extreme SFX spam.

When a benchmark reports thousands of successful handles during a burst, it means the calls were accepted successfully. It does not necessarily mean Riff created thousands of unique real voices. Under overload, v1.1.3 may coalesce or reuse voices internally to protect the mixer.

## VBA Editor Safety

Riff uses a native timer callback to drive the render loop. That gives Office projects real-time audio behavior, but it also means the VBA Editor must be treated carefully during development.

The most dangerous situation is editing code while audio is active. The VBE may internally recompile modules, move project state, invalidate procedure addresses, or reset runtime structures while the render callback is still scheduled. In normal Office automation this is unusual, but an audio engine with native callbacks can hit that edge case if the developer changes code while the sound is still playing.

The current stop-safe build includes cleanup for idle and editor reset cases:

- when idle, the timer can auto-suspend to release the VBE;
- when the VBE Stop/Reset button is clicked, stale timer callbacks attempt to kill themselves;
- `RiffEditorEmergencyStop` is available as a manual recovery helper for development sessions;
- `RiffPrepareForVbeEdit` and `RiffResumeAfterVbeEdit` provide a safer manual workflow before and after editing code.

Recommended workflow before changing code while the engine is active:

```vb
Public Sub EnterSafeEditMode()
    RiffPrepareForVbeEdit
End Sub

Public Sub LeaveSafeEditMode()
    RiffResumeAfterVbeEdit
End Sub
```

Emergency cleanup during development:

```vb
Public Sub DevAudioReset()
    RiffEditorEmergencyStop
    RiffClose
End Sub
```

`RiffEditorSafeMode` can be enabled for automatic editor protection experiments, and `RiffEditorTimerSuspended` can be inspected for diagnostics. It is disabled by default so normal playback tests from the VBE do not get paused unexpectedly.

Use these helpers only while developing. Normal applications should call `RiffClose` from their shutdown path.


## Feature Summary

| Category | Features |
|:---|:---|
| **Core** | Single `.bas`, dynamic buffers, dynamic voices, dynamic bus state, x86/x64 support |
| **I/O** | Media Foundation decoding, in-memory loading, WAV fast path, WAV export |
| **Asset Management** | Asset registry, key-based playback, key-based stop/fade, cooperative preload queue |
| **Playback** | Unified `RiffPlay`, `RiffPlayOnce`, `RiffPlayKey`, `RiffPlayKeyOnce`, finite oscillator/noise one-shots, seeking, looping, fades, stop/reset behavior |
| **Routing** | Bus volume, mute, solo, fade, peak/RMS meters, persistent bus presets, master volume |
| **Scenes** | Built-in templates for normal, pause, battle, night, cave, underwater, radio, dream, horror, retro, and cinematic states |
| **Stability** | Adaptive buffering, automatic overload protection, burst coalescing, voice stealing, active voice budgets, dynamic pool growth, idle timer cleanup, Stop/Reset-safe editor cleanup, manual VBE edit guards |
| **Presets** | Core FX presets, musical preset packs, v1.1.3 expanded presets, persistent bus effects, master presets |
| **Master Processing** | Soft clip, limiter, low-pass, high-pass, 3-band EQ, tilt EQ, compressor, drive, stereo width, balance, output gain |
| **Synthesis** | BLEP sine/square/saw, finite oscillator beeps, one-shot and looped white/pink/brown noise |
| **Dynamics** | Compressor, soft clipping, limiting, distortion, bitcrusher |
| **Spatial** | Reverb, delay, stereo width, pan, auto-pan |
| **Modulation** | Chorus, flanger, tremolo, ring modulation |
| **Filters** | Biquad low-pass, high-pass, Hz-based filter helper, 3-band EQ, master tilt EQ |
| **Diagnostics** | Underruns, render errors, clipped samples, overload stolen/dropped/coalesced counters, buffer state, active voice, bus voice, buffer voice, RMS, peak, and capacity counters |
| **Export** | Loaded buffer export and oscillator render to PCM WAV |


## Riff v1.1.3 Highlights

- Added the asset registry workflow with `RiffLoadAs`, `RiffRegisterAsset`, `RiffAssetExists`, `RiffAssetHandle`, `RiffAssetPath`, `RiffReloadAsset`, `RiffAssetCount`, `RiffAssetKey`, `RiffUnloadKey`, and `RiffClearAssets`.
- Added key-based playback with `RiffPlayKey`, `RiffPlayKeyOnce`, `RiffPlayKeyLoop`, `RiffStopKey`, and `RiffFadeOutKey`.
- Added cooperative preload helpers: `RiffPreloadAdd`, `RiffPreloadStart`, `RiffPreloadUpdate`, `RiffPreloadUpdateAll`, `RiffPreloadFinished`, `RiffPreloadProgress`, `RiffPreloadLoadedCount`, `RiffPreloadFailedCount`, `RiffPreloadTotalCount`, `RiffPreloadCurrentKey`, `RiffPreloadItemStatus`, `RiffPreloadItemKey`, `RiffPreloadItemHandle`, and `RiffPreloadClear`.
- Added scene templates through `RiffApplyScene` and `RiffScenePreset` for normal, pause menu, battle, night, cave, underwater, radio call, dream, horror, retro, and cinematic states.
- Added expanded v1.1.3 presets: `RiffFxCassette`, `RiffFxOldComputer`, `RiffFxArcadeCabinet`, `RiffFxDreamMenu`, `RiffFxSpaceStation`, `RiffFxDungeon`, `RiffFxBossRoom`, `RiffFxLowHealth`, `RiffFxMemoryFlashback`, `RiffFxCutscene`, and `RiffFxCombatImpact`.
- Added automatic overload protection that detects extreme play-call spam internally without requiring public `Begin`/`End` burst calls.
- Added internal auto-burst/coalescing behavior so repeated same-tick sounds can be accepted without forcing thousands of redundant real voices into the mixer.
- Added `RiffOverloadProtectionEnabled`, `RiffLoopCoalescingEnabled`, `RiffMaxActiveVoices`, `RiffOverloadDroppedCount`, `RiffOverloadStolenCount`, `RiffOverloadCoalescedCount`, `RiffResetOverloadStats`, and `RiffApplyStressSafeDefaults`.
- Added master limiter and safety tools: `RiffMasterSetLimiter`, `RiffMasterLimiterEnabled`, `RiffMasterLimiterCeiling`, `RiffMasterOutputGain`, and `RiffSoftClipEnabled` stress-safe setup.
- Added master mix helpers and diagnostics: `RiffMasterBalance`, `RiffMasterTiltEq`, `RiffMasterGetRms`, `RiffMasterGetRmsDb`, `RiffMasterRmsDb`, `RiffMasterPeakDb`, `RiffMasterClipCount`, and `RiffResetDiagnostics`.
- Added/expanded bus and voice RMS/clip diagnostics for balancing, visualizers, and stress tests.
- Improved burst benchmarks from the older `~11–13 us/call` dry path to around `~4–5 us/call` in the latest local stress benchmark.
- Preserved the normal user-facing workflow: existing code can keep calling `RiffPlay`, `RiffPlayKey`, `RiffPlayNoise`, and `RiffPlayOscillator` normally.

## v1.1.2 Highlights


- Reworked buffer and voice storage from fixed-size pools into dynamic `SafeArray` pools.
- Removed the old hardcoded capacity ceiling for loaded buffers and active voices.
- Added dynamic bus-state capacity while preserving the named bus workflow.
- Added `RiffReserveBuffers`, `RiffReserveVoices`, and `RiffReserveBuses` for projects that want to preallocate capacity during startup.
- Updated `RiffMaxBuffers`, `RiffMaxVoices`, and `RiffMaxBuses` so they describe current capacity instead of an old fixed limit.
- Kept handle-based compatibility so existing code can continue using numeric buffer and voice handles.
- Added manual editor-safe development helpers: `RiffPrepareForVbeEdit` and `RiffResumeAfterVbeEdit`.
- Added `RiffEditorSafeMode` and `RiffEditorTimerSuspended` for optional automatic protection and diagnostics.
- Improved safety around timer/callback state when the developer edits code in the VBE while audio is active.
- Changed default burst-cap behavior so `0` means no cap when a project intentionally wants unrestricted layering.
- Preserved the v1.0.9 playback, DSP, presets, adaptive buffering, WAV fast path, and editor emergency cleanup features.

## v1.0.9 Highlights

- Added expanded musical voice preset packs.
- Added persistent bus presets for current and future voices.
- Added master bus processors.
- Added master presets for glue, warm, bright, dark, radio, cinematic, night, and soft-limiter mixes.
- Added manual master controls for filters, EQ, compression, drive, stereo width, output gain, and soft clipping.
- Improved scene-level workflows for underwater, cave, dream, horror, retro, radio, and cinematic states.
- Added gameplay-stability refinements for repeated SFX, procedural one-shots, and voice reuse.
- Added finite `durationSec` support for generated oscillator and noise playback.
- Added explicit `RiffPlayNoiseLoop` helpers for continuous procedural ambience.
- Added `RiffVoiceSetFilterHz` for frequency-based filter setup.
- Added editor-safe timer cleanup and `RiffEditorEmergencyStop` for development recovery.
- Improved `RiffPlay` and preset setup performance through lazy temporal-buffer preparation.
- Preserved the unified playback API and compatibility wrappers from v1.0.8.
- Preserved adaptive buffering, burst safety, noise generation, WAV fast path, and diagnostics.

## Documentation

- [**API Reference**](docs/API_REFERENCE.md) – Detailed guide to every function, property, enum, and practical pattern.
- [**Architecture**](docs/ARCHITECTURE.md) – Deep dive into WASAPI, Media Foundation, thunks, callback flow, buffers, and DSP.
- [**Effect Cookbook**](docs/EFFECT_COOKBOOK.md) – Ready-to-use recipes for voice presets, bus scenes, master processing, ambience, retro effects, and more.
- [**Troubleshooting**](docs/TROUBLESHOOTING.md) – Fixes for initialization, device, stutter, cleanup, and host-specific issues.
- [**Examples**](examples/README.md) – Practical demos and integration patterns.
- [**Benchmarks**](benchmarks/) – Optional local stress tests for burst playback, game-loop playback, memory, underruns, and VBE timer behavior.


## Roadmap

- [x] WASAPI Shared Mode playback.
- [x] Media Foundation decoding.
- [x] WAV fast path loader.
- [x] Unified playback API.
- [x] Dynamic SafeArray buffer pool.
- [x] Dynamic polyphonic mixer with expandable voice capacity.
- [x] Manual capacity reservation helpers.
- [x] Named bus routing system backed by dynamically sized bus state.
- [x] Bus mute, solo, fade, peak meters, RMS meters, and clip counters.
- [x] Persistent bus effect presets.
- [x] Full per-voice Studio DSP pipeline.
- [x] Core effect presets.
- [x] Musical preset packs.
- [x] Expanded v1.1.3 effect presets.
- [x] Master bus processors.
- [x] Master processor presets.
- [x] Master limiter, output headroom, balance, tilt EQ, RMS, peak, and clip diagnostics.
- [x] White, pink, and brown noise.
- [x] One-shot and looped procedural noise helpers.
- [x] BLEP oscillators.
- [x] Finite oscillator duration support.
- [x] Hz-based voice filter helper.
- [x] Adaptive buffering.
- [x] Burst-safe voice management.
- [x] Automatic overload protection and internal auto-burst coalescing.
- [x] Stress-safe defaults for SFX-heavy projects.
- [x] Lazy temporal-buffer preparation for faster `Play` and preset setup.
- [x] Asset registry and key-based playback.
- [x] Cooperative preload helpers.
- [x] Built-in scene templates.
- [x] VBE-safe idle, Stop/Reset timer cleanup, and manual safe-edit guards.
- [x] x86/x64 native thunk driver.
- [x] Offline WAV export.
- [x] Diagnostics, render counters, overload counters, meters, and current capacity counters.

### Planned

- [ ] Optional native decode backend for faster MP3/OGG loading.
- [ ] Optional priority groups for more explicit music/voice/SFX stealing rules.
- [ ] More musical preset packs and scene templates.
- [ ] Additional visualizer helpers for PowerPoint/Excel UI meters.
- [ ] macOS support through CoreAudio/AudioToolbox if the project expands beyond Windows.


## Performance Notes

For best performance in Office:

- Prefer WAV for short SFX and UI sounds.
- Preload assets at startup with `RiffLoad` or `RiffLoadAs`.
- Use the asset registry for large projects so gameplay code can call `RiffPlayKey` instead of carrying many handle variables.
- Use cooperative preload when a project needs a visible loading screen or a staged asset-loading flow.
- For large projects, reserve capacity at startup with `RiffReserveBuffers`, `RiffReserveVoices`, and `RiffReserveBuses`.
- For SFX-heavy games, call `RiffApplyStressSafeDefaults 64, 6, 32` as a strong baseline.
- Avoid decoding MP3/OGG during interaction-heavy moments.
- Use `RiffPlayOnce` or `RiffPlayKeyOnce` for music and ambience.
- Use `RiffPlayNoiseLoop` for continuous procedural ambience instead of looping a default `RiffPlayNoise` one-shot.
- Use `durationSec` for short oscillator beeps and procedural noise hits.
- Route audio to meaningful buses such as `RiffBusMusic`, `RiffBusSfx`, `RiffBusUi`, and `RiffBusVoice`; this makes scene templates and bus effects much more useful.
- Use `RiffApplyScene` for broad scene states instead of manually rebuilding the same bus/master settings everywhere.
- Use bus volume/fades instead of changing many voices one by one.
- Use persistent bus presets for scene-wide effects.
- Use master processors lightly for final polish rather than heavy per-sample coloration on every voice.
- Keep effect-heavy processing for important voices only.
- Use overload counters to inspect heavy scenes: `RiffOverloadStolenCount`, `RiffOverloadDroppedCount`, and `RiffOverloadCoalescedCount`.
- Use `RiffMaxVoicesPerBuffer` and `RiffMaxVoicesPerBus` to prevent repeated button clicks or collision sounds from stacking too many copies, or set them to `0` when you intentionally want no explicit cap.
- Use the master limiter/headroom path for loud scenes: `RiffMasterSetLimiter`, `RiffMasterOutputGain`, and `RiffSoftClipEnabled`.
- Use `RiffVoiceSetFilterHz` when frequency-based filtering is clearer than normalized filter values.
- Keep `RiffAutoSuspendTimer` enabled for editor stability unless you intentionally need a warm timer during heavy gameplay bursts.
- Use `RiffPrepareForVbeEdit` before editing code while audio is active, then `RiffResumeAfterVbeEdit` when finished.
- Use `RiffEditorEmergencyStop` only as a development recovery helper if an old session leaves the VBE stuck.
- Always call `RiffClose` on exit.

> [!NOTE]
> Riff is a pure VBA engine with a native callback bridge. It is highly capable for Office, but it is still hosted inside Excel, PowerPoint, Word, or Access. If the host application or the entire system stalls hard enough, audio scheduling can be affected. Adaptive buffering and overload protection reduce this, but a fully independent audio thread would require a native backend.

## License

MIT. Designed for freedom, integration, and serious audio experimentation inside Microsoft Office.
