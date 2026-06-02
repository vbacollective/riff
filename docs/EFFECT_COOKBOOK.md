# Effect Cookbook

This cookbook contains practical Riff effect recipes you can paste into a VBA project and adapt. Each recipe expects a valid voice handle returned by `RiffPlay` or `RiffPlayOscillator`.

## Before You Start

Always check the voice handle before applying settings:

```vb
Dim voiceId As Long
voiceId = RiffPlay(bufferId)

If voiceId < 0 Then
    Debug.Print "No voice slot is available."
    Exit Sub
End If
```

Riff settings are per voice. If you play the same buffer three times, each returned voice handle can have different effects.

## Clean UI Click

Use this for worksheet buttons, form controls, menu actions, and small feedback sounds.

```vb
Public Sub ApplyCleanClick(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceVolume(voiceId) = 0.55
    RiffVoiceHighPass(voiceId) = 0.18
    RiffVoiceLowPass(voiceId) = 0.92
    RiffVoicePan(voiceId) = 0
End Sub
```

Why it works: a mild high-pass removes low-frequency weight that can make UI sounds feel heavy, while the light low-pass keeps sharp clicks from sounding harsh.

## Radio Voice

Use this for narration, simulated speakers, dispatch audio, or compact dialog effects.

```vb
Public Sub ApplyRadioVoice(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceHighPass(voiceId) = 0.35
    RiffVoiceLowPass(voiceId) = 0.42
    RiffVoiceDistortion(voiceId) = 1.6
    RiffVoiceEqBass(voiceId) = 0.55
    RiffVoiceEqMid(voiceId) = 1.45
    RiffVoiceEqTreble(voiceId) = 0.75
End Sub
```

Why it works: radio and intercom sounds are narrow-band. The filters remove deep bass and upper air, while mid boost and light distortion add presence.

## Underwater or Dream State

Use this for muffled scenes, transitions, paused states, or surreal audio moments.

```vb
Public Sub ApplyUnderwater(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceLowPass(voiceId) = 0.08
    RiffVoiceReverbMix(voiceId) = 0.55
    RiffVoiceReverbTime(voiceId) = 0.85
    RiffVoiceChorusDepth(voiceId) = 0.35
    RiffVoiceChorusRate(voiceId) = 0.35
End Sub
```

Why it works: the very low cutoff removes clarity, while reverb and slow chorus make the signal feel distant and unstable.

## Cathedral Pad

Use this with oscillators or sustained music beds.

```vb
Public Sub ApplyCathedralPad(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceVolume(voiceId) = 0.16
    RiffVoiceReverbMix(voiceId) = 0.85
    RiffVoiceReverbTime(voiceId) = 0.95
    RiffVoiceStereoWidth(voiceId) = 1.25
    RiffVoiceLowPass(voiceId) = 0.78
End Sub
```

Why it works: high reverb mix and long reverb time create the space. Lower volume is important because dense reverb can build up quickly across multiple voices.

## Lo-Fi Bitcrush

Use this for retro game tones, degraded transitions, glitch effects, or stylized alerts.

```vb
Public Sub ApplyLoFiCrush(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceBitDepth(voiceId) = 5
    RiffVoiceSampleRateReduction(voiceId) = 8
    RiffVoiceDistortion(voiceId) = 1.8
    RiffVoiceLowPass(voiceId) = 0.7
End Sub
```

Why it works: bit-depth reduction creates stepped amplitude, sample-rate reduction creates aliasing, and low-pass keeps the result from becoming painfully bright.

## Sci-Fi Ring Mod

Use this for robotic voices, alien tones, alarms, or metallic effects.

```vb
Public Sub ApplySciFiRingMod(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceRingModFreq(voiceId) = 180
    RiffVoiceRingModMix(voiceId) = 0.7
    RiffVoiceDelayTime(voiceId) = 0.22
    RiffVoiceDelayFeedback(voiceId) = 0.35
    RiffVoiceDelayMix(voiceId) = 0.25
End Sub
```

Why it works: ring modulation adds non-harmonic sidebands. A short delay gives the result extra depth without hiding the metallic character.

## Wide Chorus Music Bed

Use this for background music or ambient loops that need width without dominating the mix.

```vb
Public Sub ApplyWideChorus(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceVolume(voiceId) = 0.35
    RiffVoiceChorusDepth(voiceId) = 0.55
    RiffVoiceChorusRate(voiceId) = 0.8
    RiffVoiceStereoWidth(voiceId) = 1.35
    RiffVoiceReverbMix(voiceId) = 0.18
End Sub
```

Why it works: chorus adds small modulation and width. Keeping reverb low prevents the loop from washing out repeated playback.

## Slow Tremolo Pulse

Use this for tension beds, warning states, and pulsing synthetic sounds.

```vb
Public Sub ApplySlowPulse(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceTremoloRate(voiceId) = 2.2
    RiffVoiceTremoloDepth(voiceId) = 0.65
    RiffVoiceAutoPanRate(voiceId) = 0.18
    RiffVoiceAutoPanDepth(voiceId) = 0.35
End Sub
```

Why it works: tremolo creates rhythmic volume movement, while shallow auto-pan adds motion without making the sound distracting on speakers.

## Jet Flanger

Use this for dramatic build-ups, transition sweeps, or to add movement to a static noise or pad sound.

```vb
Public Sub ApplyJetFlanger(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceFlangerDepth(voiceId) = 0.85
    RiffVoiceFlangerRate(voiceId) = 0.2
    RiffVoiceFlangerFeedback(voiceId) = 0.8
    RiffVoiceStereoWidth(voiceId) = 1.2
End Sub
```

Why it works: high feedback on the flanger creates a sharp, resonant comb filter. A slow rate sweeps this resonance up and down the frequency spectrum, creating the classic "jet plane" sound.

## Punchy Drum Compression

Use this to make percussive sounds or heavy sound effects cut through the mix without clipping the master bus.

```vb
Public Sub ApplyPunchyCompression(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoiceVolume(voiceId) = 1.0
    RiffVoiceCompressorThreshold(voiceId) = 0.4
    RiffVoiceCompressorRatio(voiceId) = 8.0
    RiffVoiceEqBass(voiceId) = 1.5
End Sub
```

Why it works: a low threshold and high ratio aggressively squash the audio, while the bass boost adds weight. The compressor keeps the boosted bass from overwhelming the overall output.

## Tape Slowdown

Use this to simulate a turntable losing power or a dramatic "time-stop" effect.

```vb
Public Sub ApplyTapeSlowdown(ByVal voiceId As Long)
    If voiceId < 0 Then Exit Sub

    RiffVoicePitch(voiceId) = 0.5
    RiffVoiceLowPass(voiceId) = 0.6
    RiffVoiceDistortion(voiceId) = 1.2
End Sub
```

Why it works: halving the pitch drastically drops the frequency and doubles the duration, while a mild low-pass and distortion simulate the physical artifacts of an analog tape machine running out of momentum.

## Soft Ducking Approximation

Riff does not include side-chain routing, but you can manually lower a music bus while a voice or dialog clip plays.

```vb
Public Sub PlayDialogWithMusicDuck(ByVal dialogBuffer As Long)
    Const BUS_MUSIC As Long = 0
    Const BUS_DIALOG As Long = 2

    Dim voiceId As Long
    voiceId = RiffPlay(dialogBuffer)
    If voiceId < 0 Then Exit Sub

    RiffVoiceBus(voiceId) = BUS_DIALOG
    RiffBusVolume(BUS_MUSIC) = 0.22
    RiffBusVolume(BUS_DIALOG) = 1!
End Sub

Public Sub RestoreMusicAfterDialog()
    Const BUS_MUSIC As Long = 0

    RiffBusVolume(BUS_MUSIC) = 0.5
End Sub
```

Why it works: bus volume gives you a simple global control layer. For timed dialog, restore the music bus after your form timer or application event determines the line is finished.

## Oscillator Alert Presets

These examples create useful synthesized tones without loading files.

```vb
Public Sub PlaySoftSineAlert()
    If Not RiffOpen() Then Exit Sub

    Dim voiceId As Long
    voiceId = RiffPlayOscillator(0, 880)
    If voiceId < 0 Then Exit Sub

    RiffVoiceVolume(voiceId) = 0.18
    RiffVoiceReverbMix(voiceId) = 0.12
    RiffFadeOut voiceId, 0.55
End Sub

Public Sub PlayRetroSquareAlert()
    If Not RiffOpen() Then Exit Sub

    Dim voiceId As Long
    voiceId = RiffPlayOscillator(1, 660)
    If voiceId < 0 Then Exit Sub

    RiffVoiceVolume(voiceId) = 0.16
    RiffVoiceBitDepth(voiceId) = 6
    RiffVoiceSampleRateReduction(voiceId) = 4
    RiffFadeOut voiceId, 0.35
End Sub
```

## Preset Application Pattern

For application code, keep effect recipes as small procedures and call them immediately after `RiffPlay`.

```vb
Public Sub PlayProcessedClip(ByVal bufferId As Long)
    Dim voiceId As Long
    voiceId = RiffPlay(bufferId)

    If voiceId >= 0 Then
        ApplyRadioVoice voiceId
        RiffVoiceVolume(voiceId) = 0.75
    End If
End Sub
```

This keeps playback code readable and makes presets easy to tune without searching through business logic.
