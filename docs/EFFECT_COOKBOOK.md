# Effect Cookbook

This cookbook contains practical Riff effect recipes you can paste into VBA projects. It covers the v1.0.9 preset system, musical preset packs, persistent bus effects, master bus processors, procedural noise, manual DSP recipes, gameplay-safe one-shots, Hz-based filtering, performance-oriented preset usage, and VBE-safe cleanup patterns.

Riff effects can be applied at three levels:

1. **Voice effects**: apply to one active voice.
2. **Bus effects**: apply to all current and/or future voices routed to a bus.
3. **Master processors**: apply to the final mixed output.

```vb
' Voice-level effect
RiffVoiceApplyPreset voiceId, RiffFxRadio, 0.65

' Bus-level effect for current and future voices
RiffBusApplyPreset RiffBusMusic, RiffFxUnderwater, 0.55

' Master-level final mix processing
RiffMasterApplyPreset RiffMasterFxGlue, 0.7
```

## Quick Start

### Apply a Voice Preset

```vb
Public Sub MakeVoiceRadio(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceApplyPreset voiceId, RiffFxRadio, 0.65
End Sub
```

### Apply a Bus-Wide Scene Effect

```vb
Public Sub EnterUnderwaterAudio()
    RiffBusApplyPreset RiffBusMusic, RiffFxUnderwater, 0.55
    RiffBusApplyPreset RiffBusSfx, RiffFxUnderwater, 0.8
    RiffBusFadeTo RiffBusMusic, 0.3, 600
End Sub

Public Sub LeaveUnderwaterAudio()
    RiffBusClearEffects RiffBusMusic
    RiffBusClearEffects RiffBusSfx
    RiffBusFadeTo RiffBusMusic, 0.45, 600
End Sub
```

### Apply a Master Mix Preset

```vb
Public Sub EnableCinematicMix()
    RiffMasterApplyPreset RiffMasterFxCinematic, 0.65
End Sub

Public Sub ResetMasterMix()
    RiffMasterClearProcessors
    RiffMasterApplyPreset RiffMasterFxClean
End Sub
```

### Gameplay-Safe One-Shot SFX

For fast gameplay effects, prefer finite one-shots. The current Riff build is optimized for this path and protects against accidental accumulation.

```vb
Public Sub PlayUiTick()
    RiffPlayOscillator RiffWaveSquare, 1320!, RiffBusUi, 0.16!, 0!, 0.035!
End Sub

Public Sub PlayDustHit()
    Dim v As Long

    v = RiffPlayNoise(RiffWavePinkNoise, RiffBusSfx, 0.09!, 0!, 0.045!)
    If v >= 0 Then RiffVoiceSetFilterHz v, 2200!, 120!
End Sub
```

### Continuous Noise Ambience

Use the explicit loop helper for ambience. Do not use default `RiffPlayNoise` for long ambience, because it is now a finite one-shot by default.

```vb
Private windVoice As Long

Public Sub StartWindBed()
    windVoice = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.04!, 0!)
    If windVoice >= 0 Then
        RiffVoiceSetFilterHz windVoice, 1800!, 80!
        RiffVoiceApplyPreset windVoice, RiffFxWind, 0.6!
        RiffFadeIn windVoice, 1.5!
    End If
End Sub

Public Sub StopWindBed()
    If windVoice >= 0 Then
        RiffFadeOut windVoice, 1.2!
        windVoice = -1
    End If
End Sub
```

## The Preset System

The easiest way to apply complex effects is to use the built-in preset engine. A preset configures multiple DSP stages in a single call: filters, EQ, reverb, chorus, delay, compression, distortion, stereo width, and modulation.

```vb
' Apply a preset at full strength
RiffVoiceApplyPreset voiceId, RiffFxRadio

' Apply a subtle ambient wash
RiffVoiceApplyPreset voiceId, RiffFxAmbient, 0.5
```

`amount` controls intensity. Most presets sound best between `0.35` and `0.85`.

```vb
RiffVoiceApplyPreset v, RiffFxWarmTape, 0.35  ' subtle
RiffVoiceApplyPreset v, RiffFxWarmTape, 0.7   ' obvious
RiffVoiceApplyPreset v, RiffFxWarmTape, 1     ' full
```

### Important Behavior

`RiffVoiceApplyPreset` is intended to give a known, predictable result. Apply a preset first, then make manual adjustments.

```vb
RiffVoiceApplyPreset v, RiffFxRadio, 0.65

' Manual tweaks after preset
RiffVoiceVolume(v) = 0.8
RiffVoicePan(v) = -0.25
RiffVoiceCompressorRatio(v) = 4
```

To clear the DSP chain:

```vb
RiffVoiceClearEffects v
```

Or use the dry preset:

```vb
RiffVoiceApplyPreset v, RiffFxDry
```

### Performance Notes for Presets

The current performance build uses lazy temporal-buffer preparation. This means presets that use delay, reverb, chorus, or flanger no longer pay the full ring-buffer cleanup cost during the `RiffVoiceApplyPreset` call. The heavy temporal buffer is prepared only when the voice actually renders with a temporal effect enabled.

For very high-rate gameplay SFX, the recommended pattern is still:

```vb
Dim v As Long

v = RiffPlay(sndHit, RiffBusSfx, False, 0.55!, 0!)
If v >= 0 Then
    RiffVoiceApplyPreset v, RiffFxSmallRoom, 0.35!
End If
```

For repeated UI beeps and procedural SFX, prefer finite oscillator/noise calls instead of creating an infinite voice and fading it later:

```vb
RiffPlayOscillator RiffWaveSine, 880!, RiffBusUi, 0.18!, 0!, 0.045!
RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.08!, 0!, 0.035!
```

Presets are sanitized after application. Amounts and internal DSP values are clamped so exaggerated combinations are less likely to create runaway feedback, harsh pops, or invalid filter states.

## Voice Preset Reference

### Utility Presets

| Preset | Character | Best For |
|:---|:---|:---|
| `RiffFxDry` | Neutral/dry reset | Clearing effects |
| `RiffFxSmallRoom` | Short room reflections | UI, close dialogue, small spaces |
| `RiffFxHall` | Wider room ambience | Music, ambience, narration |
| `RiffFxCathedral` | Long, washed reverb | Dramatic scenes, magic, dreams |
| `RiffFxSlapback` | Short echo | Retro voice, rockabilly, UI feedback |
| `RiffFxEcho` | Longer repeated delay | Space, transitions, emphasis |
| `RiffFxChorus` | Thickened modulation | Pads, synths, soft music |
| `RiffFxFlanger` | Sweeping comb movement | Sci-fi, transitions, mechanical sounds |
| `RiffFxLoFi` | Musical degraded sampler/tape | Retro, old media, dusty UI |
| `RiffFxRadio` | Band-limited comms | Radios, intercoms, distant voice |
| `RiffFxUnderwater` | Muffled and unstable | Underwater, dream, memory effects |
| `RiffFxWide` | Enhanced stereo image | Music, ambience, pads |
| `RiffFxRobot` | Metallic ring modulation | Robots, alien voice, machines |
| `RiffFxAmbient` | Space and softness | Rain, pads, backgrounds |

### Musical Preset Packs

| Preset | Character | Best For |
|:---|:---|:---|
| `RiffFxWarmTape` | Warm, soft, subtly saturated | Music, ambience, narration |
| `RiffFxVHS` | Warbly analog degradation | Old video, flashbacks, broken memories |
| `RiffFxDreamPad` | Wide chorus and reverb | Dreams, emotional pads, menu ambience |
| `RiffFxDarkCave` | Dark, deep, reverberant | Caves, horror, underground areas |
| `RiffFxTinySpeaker` | Small bandwidth speaker | Phone, laptop, toy speaker |
| `RiffFxMegaphone` | Mid-forward projection | PA systems, announcements |
| `RiffFxGameBoy` | Crunchy retro chip color | Pixel UI, handheld game effects |
| `RiffFxHorrorDrone` | Dark modulated unease | Horror ambience, tension |
| `RiffFxWind` | Airy filtered motion | Wind layers, exterior ambience |
| `RiffFxRain` | Soft natural noise bed | Rain, ambience, weather |
| `RiffFxCinematicBoom` | Big low-heavy impact | Hits, booms, transitions |
| `RiffFxSoftFocus` | Gentle smoothing and width | Emotional scenes, soft ambience |

## Master Preset Reference

Master presets process the final mixed output. They are not replacements for voice effects. Use them for final polish, scene-level coloration, and mix safety.

```vb
RiffMasterApplyPreset RiffMasterFxGlue, 0.7
```

| Preset | Character | Best For |
|:---|:---|:---|
| `RiffMasterFxClean` | Neutral master stage | Resetting the mix |
| `RiffMasterFxGlue` | Light compression and control | General gameplay mix |
| `RiffMasterFxWarm` | Warmer tone and subtle saturation | Cozy scenes, music-heavy scenes |
| `RiffMasterFxBright` | Brighter overall mix | UI-heavy menus, clean presentations |
| `RiffMasterFxDark` | Darker and softer output | Night, caves, horror |
| `RiffMasterFxRadio` | Global band-limited sound | Full-scene radio/monitor effect |
| `RiffMasterFxCinematic` | Wide, compressed, fuller mix | Cutscenes, dramatic moments |
| `RiffMasterFxNight` | Softer, lower-energy mix | Night levels, quiet states |
| `RiffMasterFxSoftLimiter` | Safety limiting | Loud SFX-heavy scenes |

## Preset Recipes

### Radio Voice

```vb
Public Sub ApplyRadioVoice(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxRadio, 0.7
    RiffVoiceCompressorThreshold(v) = 0.55
    RiffVoiceCompressorRatio(v) = 4
End Sub
```

### Telephone Voice

Use `RiffFxTinySpeaker` for a more natural small-device sound, or `RiffFxRadio` for more aggressive communication filtering.

```vb
Public Sub ApplyTelephoneVoice(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxTinySpeaker, 0.75
    RiffVoiceSetFilterHz v, 3000!, 300!
    RiffVoiceCompressorRatio(v) = 3.5
End Sub
```

### VHS Memory

```vb
Public Sub ApplyVHSMemory(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxVHS, 0.65
    RiffVoiceStereoWidth(v) = 0.85
    RiffVoiceVolume(v) = 0.75
End Sub
```

### Warm Tape Music

```vb
Public Sub ApplyWarmTapeMusic(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxWarmTape, 0.55
    RiffVoiceCompressorThreshold(v) = 0.75
    RiffVoiceCompressorRatio(v) = 2.2
End Sub
```

### Dream Sequence

```vb
Public Sub ApplyDreamState(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxDreamPad, 0.75
    RiffVoiceAutoPanRate(v) = 0.12
    RiffVoiceAutoPanDepth(v) = 0.35
    RiffVoiceStereoWidth(v) = 1.6
End Sub
```

### Dark Cave

```vb
Public Sub ApplyDarkCave(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxDarkCave, 0.8
    RiffVoiceSetFilterHz v, 1400!, 45!
End Sub
```

### Game Boy UI

```vb
Public Sub ApplyGameBoyUi(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxGameBoy, 0.75
    RiffVoiceVolume(v) = 0.45
End Sub
```

### Cinematic Boom

```vb
Public Sub ApplyCinematicBoom(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxCinematicBoom, 0.85
    RiffVoiceEqBass(v) = 1.7
    RiffVoiceCompressorThreshold(v) = 0.45
    RiffVoiceCompressorRatio(v) = 8
End Sub
```

### Horror Drone

```vb
Public Sub ApplyHorrorDrone(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxHorrorDrone, 0.8
    RiffVoiceAutoPanRate(v) = 0.08
    RiffVoiceAutoPanDepth(v) = 0.45
End Sub
```

### Soft Focus Music

```vb
Public Sub ApplySoftFocus(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxSoftFocus, 0.55
    RiffVoiceStereoWidth(v) = 1.25
End Sub
```

## Bus-Wide Scene Recipes

Bus presets are the best way to apply a scene effect to all sounds of a category.

```vb
RiffBusApplyPreset RiffBusMusic, RiffFxUnderwater, 0.55
RiffBusApplyPreset RiffBusSfx, RiffFxUnderwater, 0.8
```

By default, `RiffBusApplyPreset` applies to active voices and stores a persistent preset for future voices.

### Underwater Scene

```vb
Public Sub EnterUnderwaterScene()
    RiffBusApplyPreset RiffBusMusic, RiffFxUnderwater, 0.55
    RiffBusApplyPreset RiffBusSfx, RiffFxUnderwater, 0.8
    RiffBusApplyPreset RiffBusVoice, RiffFxUnderwater, 0.45

    RiffBusFadeTo RiffBusMusic, 0.28, 600
    RiffBusFadeTo RiffBusSfx, 0.55, 350
    RiffMasterApplyPreset RiffMasterFxDark, 0.35
End Sub

Public Sub LeaveUnderwaterScene()
    RiffBusClearEffects RiffBusMusic
    RiffBusClearEffects RiffBusSfx
    RiffBusClearEffects RiffBusVoice

    RiffBusFadeTo RiffBusMusic, 0.45, 600
    RiffBusFadeTo RiffBusSfx, 0.9, 350
    RiffMasterApplyPreset RiffMasterFxClean
End Sub
```

### Cave Scene

```vb
Public Sub EnterCaveScene()
    RiffBusApplyPreset RiffBusMusic, RiffFxDarkCave, 0.65
    RiffBusApplyPreset RiffBusSfx, RiffFxSmallRoom, 0.45
    RiffBusApplyPreset RiffBusVoice, RiffFxDarkCave, 0.35

    RiffMasterApplyPreset RiffMasterFxDark, 0.55
End Sub

Public Sub LeaveCaveScene()
    RiffBusClearEffects RiffBusMusic
    RiffBusClearEffects RiffBusSfx
    RiffBusClearEffects RiffBusVoice

    RiffMasterClearProcessors
End Sub
```

### Dream Scene

```vb
Public Sub EnterDreamScene()
    RiffBusApplyPreset RiffBusMusic, RiffFxDreamPad, 0.75
    RiffBusApplyPreset RiffBusSfx, RiffFxSoftFocus, 0.55
    RiffBusApplyPreset RiffBusVoice, RiffFxSoftFocus, 0.35

    RiffMasterApplyPreset RiffMasterFxWarm, 0.35
    RiffBusFadeTo RiffBusMusic, 0.5, 800
End Sub

Public Sub LeaveDreamScene()
    RiffBusClearEffects RiffBusMusic
    RiffBusClearEffects RiffBusSfx
    RiffBusClearEffects RiffBusVoice

    RiffMasterApplyPreset RiffMasterFxClean
End Sub
```

### Radio Transmission Scene

```vb
Public Sub EnterRadioScene()
    RiffBusApplyPreset RiffBusVoice, RiffFxRadio, 0.75
    RiffBusApplyPreset RiffBusMusic, RiffFxTinySpeaker, 0.35

    RiffMasterApplyPreset RiffMasterFxRadio, 0.25
End Sub

Public Sub LeaveRadioScene()
    RiffBusClearEffects RiffBusVoice
    RiffBusClearEffects RiffBusMusic

    RiffMasterApplyPreset RiffMasterFxClean
End Sub
```

### Persistent Bus Preset Only for Future Voices

```vb
Public Sub SetFutureVoiceRadio()
    RiffBusApplyPreset RiffBusVoice, RiffFxRadio, 0.65, True, False
End Sub
```

### Apply Preset Only to Current Voices

```vb
Public Sub WashCurrentSfxOnly()
    RiffBusApplyPreset RiffBusSfx, RiffFxHall, 0.5, False, True
End Sub
```

### Clear Active Effects But Keep Persistent Bus Preset

```vb
Public Sub ClearCurrentVoiceColorButKeepFutureRule()
    RiffBusClearEffects RiffBusVoice, False
End Sub
```

## Master Processor Recipes

Master processors affect the final output. They are ideal for scene-level polish and safety processing.

### General Mix Glue

```vb
Public Sub ApplyMixGlue()
    RiffMasterApplyPreset RiffMasterFxGlue, 0.65
End Sub
```

### Cinematic Mix

```vb
Public Sub ApplyCinematicMaster()
    RiffMasterApplyPreset RiffMasterFxCinematic, 0.65
    RiffMasterOutputGain = 0.95
End Sub
```

### Warm Master

```vb
Public Sub ApplyWarmMaster()
    RiffMasterApplyPreset RiffMasterFxWarm, 0.55
End Sub
```

### Night Master

```vb
Public Sub ApplyNightMaster()
    RiffMasterApplyPreset RiffMasterFxNight, 0.7
    RiffBusFadeTo RiffBusMusic, 0.28, 800
    RiffBusFadeTo RiffBusSfx, 0.55, 300
End Sub
```

### Radio Master

```vb
Public Sub ApplyRadioMaster()
    RiffMasterApplyPreset RiffMasterFxRadio, 0.6
End Sub
```

### Soft Limiter Safety

Use this for SFX-heavy scenes where many one-shots may overlap.

```vb
Public Sub ApplyLimiterSafety()
    RiffMasterApplyPreset RiffMasterFxSoftLimiter, 0.85
    RiffSoftClipEnabled = True
End Sub
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

### Clear Master Processing

```vb
Public Sub ClearMasterProcessing()
    RiffMasterClearProcessors
    RiffSoftClipEnabled = True
    RiffMasterVolume = 1
End Sub
```

## Manual Voice Recipes

If you need fine-grained control beyond presets, use these manual recipes as starting points.

### Hz-Based Filter Cheat Sheet

Use `RiffVoiceSetFilterHz` when you want predictable, audio-style filter points instead of normalized filter values.

```vb
' Radio/telephone band.
RiffVoiceSetFilterHz v, 3000!, 300!

' Muffled wall/underwater tone.
RiffVoiceSetFilterHz v, 900!, 0!

' Remove low rumble but keep the source mostly open.
RiffVoiceSetFilterHz v, 0!, 80!
```

Typical starting points:

| Effect | Low-pass Hz | High-pass Hz |
|:---|---:|---:|
| Telephone / radio | `3000` | `300` |
| Muffled wall | `700` to `1200` | `0` |
| Underwater | `500` to `1000` | `20` to `80` |
| Soft rain / wind | `1800` to `3500` | `80` to `150` |
| Remove sub-rumble | `0` | `40` to `100` |
| Bright UI click polish | `5000` to `8000` | `120` to `300` |


### Clean UI Click

Perfect for worksheet buttons, PowerPoint buttons, and form interactions.

```vb
Public Sub ApplyCleanClick(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceSetFilter v, 0.92, 0.18
    RiffVoiceVolume(v) = 0.6
End Sub
```

### Clean Confirm Sound

```vb
Public Sub ApplyConfirmSound(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceVolume(v) = 0.65
    RiffVoiceEqBass(v) = 0.85
    RiffVoiceEqMid(v) = 1.05
    RiffVoiceEqTreble(v) = 1.15
    RiffVoiceSetReverb v, 0.08, 0.2
End Sub
```

### Error / Deny Sound

```vb
Public Sub ApplyDenySound(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoicePitch(v) = 0.82
    RiffVoiceVolume(v) = 0.7
    RiffVoiceLowPass(v) = 0.65
    RiffVoiceDistortion(v) = 1.15
End Sub
```

### Cinematic Impact

```vb
Public Sub ApplyCinematicImpact(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceDistortion(v) = 1.4
    RiffVoiceEqBass(v) = 1.8
    RiffVoiceSetReverb v, 0.45, 0.8
    RiffVoiceCompressorThreshold(v) = 0.4
    RiffVoiceCompressorRatio(v) = 12
End Sub
```

### Retro Chip

```vb
Public Sub ApplyRetroChip(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceBitDepth(v) = 6
    RiffVoiceSampleRateReduction(v) = 3
    RiffVoiceEqTreble(v) = 0.5
End Sub
```

### Soft Pad

```vb
Public Sub ApplySoftPad(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceSetChorus v, 0.45, 1.1
    RiffVoiceSetReverb v, 0.55, 0.8
    RiffVoiceStereoWidth(v) = 1.45
    RiffVoiceLowPass(v) = 0.75
End Sub
```

### Sci-Fi Sweep

```vb
Public Sub ApplySciFiSweep(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoiceSetFlanger v, 0.75, 0.35, 0.55
    RiffVoiceRingModFreq(v) = 90
    RiffVoiceRingModMix(v) = 0.25
    RiffVoiceStereoWidth(v) = 1.4
End Sub
```

### Monster Voice

```vb
Public Sub ApplyMonsterVoice(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoicePitch(v) = 0.68
    RiffVoiceDistortion(v) = 1.35
    RiffVoiceLowPass(v) = 0.62
    RiffVoiceEqBass(v) = 1.5
    RiffVoiceSetReverb v, 0.25, 0.55
End Sub
```

### Ghost Voice

```vb
Public Sub ApplyGhostVoice(ByVal v As Long)
    If v < 0 Then Exit Sub

    RiffVoicePitch(v) = 0.92
    RiffVoiceSetChorus v, 0.5, 0.8
    RiffVoiceSetReverb v, 0.6, 0.9
    RiffVoiceAutoPanRate(v) = 0.18
    RiffVoiceAutoPanDepth(v) = 0.45
    RiffVoiceHighPass(v) = 0.08
End Sub
```

## Synthesis Patterns

Use oscillators and noise generators for procedural effects without audio files.

### Clean UI Beep

```vb
Public Sub PlayCleanBeep()
    Dim v As Long

    v = RiffPlayOscillator(RiffWaveSine, 880!, RiffBusUi, 0.22!, 0!, 0.08!)
    If v < 0 Then Exit Sub
End Sub
```

### Retro Square Beep

```vb
Public Sub PlayRetroBeep()
    Dim v As Long

    v = RiffPlayOscillator(RiffWaveSquare, 660!, RiffBusUi, 0.18!, 0!, 0.07!)
    If v < 0 Then Exit Sub

    RiffVoiceBitDepth(v) = 8
    RiffVoiceSampleRateReduction(v) = 2
End Sub
```

### Power-Up Sweep

```vb
Public Sub PlayPowerUp()
    Dim v As Long

    v = RiffPlayOscillator(RiffWaveSawtooth, 220, RiffBusSfx, 0.18, 0)
    If v < 0 Then Exit Sub

    RiffVoicePitchTo v, 2.2, 450
    RiffVoiceVolumeTo v, 0, 500
    RiffVoiceSetChorus v, 0.25, 2.2
End Sub
```

### Alarm Pulse

```vb
Public Sub PlayAlarmPulse()
    Dim v As Long

    v = RiffPlayOscillator(RiffWaveSquare, 440, RiffBusSfx, 0.22, 0)
    If v < 0 Then Exit Sub

    RiffVoiceTremoloRate(v) = 7
    RiffVoiceTremoloDepth(v) = 0.85
    RiffVoiceDistortion(v) = 1.2
End Sub
```

### Windy Ambience

```vb
Public Sub PlayWind()
    Dim v As Long

    v = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusSfx, 0.12!, 0!)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxWind, 0.65!
    RiffVoiceSetFilterHz v, 1800!, 80!
    RiffVoiceAutoPanRate(v) = 0.05!
    RiffVoiceAutoPanDepth(v) = 0.6!
End Sub
```

### Rain Bed

```vb
Public Sub PlayRain()
    Dim v As Long

    v = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.08!, 0!)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxRain, 0.7!
    RiffVoiceSetFilterHz v, 2600!, 120!
    RiffVoiceStereoWidth(v) = 1.35!
End Sub
```

### Static / Vinyl Crackle

```vb
Public Sub PlayVinylCrackle()
    Dim v As Long

    v = RiffPlayNoise(RiffWaveWhiteNoise, RiffBusSfx, 0.05!, 0!, 0.035!)
    If v < 0 Then Exit Sub

    RiffVoiceBitDepth(v) = 3
    RiffVoiceSetFilterHz v, 7000!, 1200!
    RiffVoiceSampleRateReduction(v) = 2
End Sub
```

### Earthquake Rumble

```vb
Public Sub PlayEarthquakeRumble()
    Dim v As Long

    v = RiffPlayNoise(RiffWaveBrownNoise, RiffBusSfx, 0.16!, 0!, 2.5!)
    If v < 0 Then Exit Sub

    RiffVoiceSetFilterHz v, 180!, 25!
    RiffVoiceEqBass(v) = 1.6!
    RiffVoiceTremoloRate(v) = 4!
    RiffVoiceTremoloDepth(v) = 0.25!
End Sub
```

## Complete Scene Examples

### Underwater Level

```vb
Private underwaterAmbience As Long

Public Sub StartUnderwaterLevel()
    RiffMasterApplyPreset RiffMasterFxDark, 0.35

    RiffBusApplyPreset RiffBusMusic, RiffFxUnderwater, 0.55
    RiffBusApplyPreset RiffBusSfx, RiffFxUnderwater, 0.75
    RiffBusApplyPreset RiffBusVoice, RiffFxUnderwater, 0.45

    RiffBusFadeTo RiffBusMusic, 0.28, 800
    RiffBusFadeTo RiffBusSfx, 0.6, 400

    underwaterAmbience = RiffPlayNoiseLoop(RiffWaveBrownNoise, RiffBusMusic, 0.06!, 0!)
    If underwaterAmbience >= 0 Then
        RiffVoiceSetFilterHz underwaterAmbience, 650!, 35!
        RiffVoiceStereoWidth(underwaterAmbience) = 1.3!
        RiffFadeIn underwaterAmbience, 2!
    End If
End Sub

Public Sub StopUnderwaterLevel()
    If underwaterAmbience >= 0 Then
        RiffFadeOut underwaterAmbience, 1.5
        underwaterAmbience = -1
    End If

    RiffBusClearEffects RiffBusMusic
    RiffBusClearEffects RiffBusSfx
    RiffBusClearEffects RiffBusVoice

    RiffBusFadeTo RiffBusMusic, 0.45, 800
    RiffBusFadeTo RiffBusSfx, 0.9, 400

    RiffMasterApplyPreset RiffMasterFxClean
End Sub
```

### Horror Scene

```vb
Private horrorDrone As Long

Public Sub StartHorrorScene()
    RiffMasterApplyPreset RiffMasterFxNight, 0.75

    RiffBusApplyPreset RiffBusMusic, RiffFxDarkCave, 0.55
    RiffBusApplyPreset RiffBusSfx, RiffFxHorrorDrone, 0.35

    horrorDrone = RiffPlayNoiseLoop(RiffWaveBrownNoise, RiffBusMusic, 0.08!, 0!)
    If horrorDrone >= 0 Then
        RiffVoiceApplyPreset horrorDrone, RiffFxHorrorDrone, 0.8!
        RiffVoiceSetFilterHz horrorDrone, 320!, 25!
        RiffFadeIn horrorDrone, 3!
    End If
End Sub

Public Sub StopHorrorScene()
    If horrorDrone >= 0 Then
        RiffFadeOut horrorDrone, 2
        horrorDrone = -1
    End If

    RiffBusClearEffects RiffBusMusic
    RiffBusClearEffects RiffBusSfx
    RiffMasterApplyPreset RiffMasterFxClean
End Sub
```

### Retro Menu

```vb
Public Sub EnterRetroMenu()
    RiffMasterApplyPreset RiffMasterFxBright, 0.25
    RiffBusApplyPreset RiffBusUi, RiffFxGameBoy, 0.75
End Sub

Public Sub LeaveRetroMenu()
    RiffBusClearEffects RiffBusUi
    RiffMasterApplyPreset RiffMasterFxClean
End Sub

Public Sub RetroMenuMove()
    Dim v As Long

    v = RiffPlayOscillator(RiffWaveSquare, 660!, RiffBusUi, 0.16!, 0!, 0.06!)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxGameBoy, 0.8!
End Sub

Public Sub RetroMenuConfirm()
    Dim v As Long

    v = RiffPlayOscillator(RiffWaveSquare, 990!, RiffBusUi, 0.18!, 0!, 0.12!)
    If v < 0 Then Exit Sub

    RiffVoiceApplyPreset v, RiffFxGameBoy, 0.7!
End Sub
```

### Cinematic Transition

```vb
Public Sub PlayCinematicTransition(ByVal boomBuffer As Long)
    Dim boom As Long
    Dim rumble As Long

    RiffMasterApplyPreset RiffMasterFxCinematic, 0.55

    boom = RiffPlay(boomBuffer, RiffBusSfx, False, 0.95, 0)
    If boom >= 0 Then
        RiffVoiceApplyPreset boom, RiffFxCinematicBoom, 0.85
    End If

    rumble = RiffPlayNoise(RiffWaveBrownNoise, RiffBusSfx, 0.1!, 0!, 1.2!)
    If rumble >= 0 Then
        RiffVoiceSetFilterHz rumble, 160!, 25!
    End If
End Sub
```

## Gameplay SFX Recipes

These recipes are designed for high-frequency game loops where the same sound can be triggered many times per second.

### Footstep Tick With Burst Safety

```vb
Public Sub PlayFootstep(ByVal sndStep As Long, Optional ByVal pan As Single = 0!)
    Dim v As Long

    v = RiffPlay(sndStep, RiffBusSfx, False, 0.38!, pan)
    If v < 0 Then Exit Sub

    RiffVoiceSetFilterHz v, 4200!, 90!
End Sub
```

Recommended global safety for footsteps, bullets, UI ticks, and collision sounds:

```vb
RiffVoiceStealingEnabled = True
RiffMaxVoicesPerBuffer = 4
RiffMaxVoicesPerBus = 18
```

### Dust Puff / Landing Noise

```vb
Public Sub PlayLandingDust()
    Dim v As Long

    v = RiffPlayNoise(RiffWavePinkNoise, RiffBusSfx, 0.12!, 0!, 0.055!)
    If v < 0 Then Exit Sub

    RiffVoiceSetFilterHz v, 1800!, 120!
    RiffVoiceApplyPreset v, RiffFxSmallRoom, 0.25!
End Sub
```

### Coin Pickup Without Audio File

```vb
Public Sub PlayCoinPickup()
    RiffPlayOscillator RiffWaveSquare, 880!, RiffBusUi, 0.18!, 0!, 0.045!
    RiffPlayOscillator RiffWaveSquare, 1320!, RiffBusUi, 0.14!, 0!, 0.06!
End Sub
```

### Low Hit Layer

```vb
Public Sub PlayLowHitLayer()
    Dim v As Long

    v = RiffPlayNoise(RiffWaveBrownNoise, RiffBusSfx, 0.13!, 0!, 0.18!)
    If v < 0 Then Exit Sub

    RiffVoiceSetFilterHz v, 190!, 25!
    RiffVoiceCompressorThreshold(v) = 0.55!
    RiffVoiceCompressorRatio(v) = 5!
End Sub
```

### Debug Voice Accumulation

```vb
Public Sub PrintAudioBurstState(ByVal snd As Long)
    Debug.Print "Active voices:", RiffActiveVoiceCount()
    Debug.Print "SFX bus voices:", RiffBusVoiceCount(RiffBusSfx)
    Debug.Print "This buffer voices:", RiffBufferVoiceCount(snd, RiffBusSfx)
    Debug.Print "Underruns:", RiffUnderrunCount
End Sub
```

## Mixing Tips

### Voice vs Bus vs Master

Use the lowest level that matches your intent:

| Goal | Best Level |
|:---|:---|
| Make one sound radio-like | Voice preset |
| Make all dialogue radio-like | Bus preset on `RiffBusVoice` |
| Make the entire scene sound like a radio broadcast | Master preset |
| Lower all music volume | Bus volume/fade |
| Prevent harsh clipping globally | Master soft clip / limiter |
| Add cave reverb to SFX and music | Bus presets |
| Add warmth to the whole mix | Master preset |

### Avoid Overprocessing

Do not stack a heavy voice preset, heavy bus preset, and heavy master preset unless the effect is intentionally extreme.

Better:

```vb
RiffVoiceApplyPreset v, RiffFxRadio, 0.65
RiffMasterApplyPreset RiffMasterFxGlue, 0.35
```

Risky:

```vb
RiffVoiceApplyPreset v, RiffFxRadio, 1
RiffBusApplyPreset RiffBusVoice, RiffFxTinySpeaker, 1
RiffMasterApplyPreset RiffMasterFxRadio, 1
```

### Keep UI Sounds Clean

UI feedback should usually be short, bright, and not too reverberant.

```vb
RiffPlay sndClick, RiffBusUi, False, 0.55, 0
```

For themed UI, use subtle bus processing:

```vb
RiffBusApplyPreset RiffBusUi, RiffFxGameBoy, 0.45
```

### Use Master Safety for SFX-Heavy Scenes

```vb
RiffSoftClipEnabled = True
RiffMasterApplyPreset RiffMasterFxSoftLimiter, 0.7
RiffMaxVoicesPerBuffer = 4
RiffMaxVoicesPerBus = 16
```

## Pro Tips

### Apply Presets Before Manual Tweaks

```vb
RiffVoiceApplyPreset v, RiffFxDreamPad, 0.6

RiffVoiceVolume(v) = 0.4
RiffVoicePan(v) = -0.2
RiffVoiceLowPass(v) = 0.7
```

### Clear Scene Effects Explicitly

```vb
RiffBusClearEffects RiffBusMusic
RiffBusClearEffects RiffBusSfx
RiffBusClearEffects RiffBusVoice
RiffMasterApplyPreset RiffMasterFxClean
```

### Use Persistent Bus Presets for Slide-Based Games

If each slide creates new voices, persistent bus presets ensure new sounds inherit the scene style automatically.

```vb
Public Sub EnterDreamSlide()
    RiffBusApplyPreset RiffBusMusic, RiffFxDreamPad, 0.7, True, True
    RiffBusApplyPreset RiffBusSfx, RiffFxSoftFocus, 0.45, True, True
End Sub
```

### Inspect Persistent Bus State

```vb
Public Sub PrintBusFxState()
    Debug.Print "Music bus preset enabled:", RiffBusPresetEnabled(RiffBusMusic)
    Debug.Print "Music bus preset:", RiffBusPreset(RiffBusMusic)
    Debug.Print "Music bus amount:", RiffBusPresetAmount(RiffBusMusic)
End Sub
```

### Reset Everything Audio-Color Related

```vb
Public Sub ResetAudioColor()
    RiffBusClearEffects RiffBusMusic
    RiffBusClearEffects RiffBusSfx
    RiffBusClearEffects RiffBusVoice
    RiffBusClearEffects RiffBusUi

    RiffMasterClearProcessors
    RiffSoftClipEnabled = True
End Sub
```


### VBE-Safe Development Cleanup

When testing audio directly from the VBE, the editor can be interrupted with the Stop/Reset button. The current stop-safe build kills orphaned timer callbacks automatically, but this helper is available as an emergency cleanup tool during development:

```vb
Public Sub EmergencyAudioEditorReset()
    RiffEditorEmergencyStop
End Sub
```

Use normal cleanup in application code:

```vb
Public Sub ShutdownAudioNormally()
    RiffClose
End Sub
```

Use `RiffEditorEmergencyStop` only for editor recovery after interrupted tests, not as a normal game shutdown path.

### Gameplay Timer Warmth

For active gameplay, you can keep the audio path warm so short SFX do not hit an empty WASAPI buffer between bursts:

```vb
Public Sub StartGameplayAudioMode()
    RiffOpen
    RiffAutoSuspendTimer = False
End Sub
```

The stop-safe build still releases the VBE when idle, so IntelliSense should not remain stuck in a permanent Running state after one-shot playback finishes.

For menus, editors, and mostly idle screens, prefer:

```vb
Public Sub StartMenuAudioMode()
    RiffOpen
    RiffAutoSuspendTimer = True
End Sub
```
