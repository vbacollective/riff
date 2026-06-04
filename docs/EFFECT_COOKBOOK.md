# Effect Cookbook

This cookbook contains practical Riff effect recipes you can paste into your VBA projects. With the addition of **Effect Presets**, applying professional audio shaping is now faster than ever.


## The Preset System

The easiest way to apply complex effects is to use the built-in preset engine. This handles the configuration of multiple DSP stages (Filters, EQ, Reverb, Chorus, etc.) in a single call.

```vb
' Apply a preset at full strength
RiffVoiceApplyPreset voiceId, RiffFxRadio

' Apply a subtle ambient wash (50% strength)
RiffVoiceApplyPreset voiceId, RiffFxAmbient, 0.5
```

### Available Presets
| Preset | Character |
|:---|:---|
| `RiffFxSmallRoom` / `RiffFxHall` / `RiffFxCathedral` | Spatial environments. |
| `RiffFxRadio` / `RiffFxTelephone` | Narrow-band communication. |
| `RiffFxUnderwater` | Muffled, resonant, and unstable. |
| `RiffFxLoFi` | Retro, bit-crushed, and degraded. |
| `RiffFxRobot` | Metallic ring-modulated textures. |
| `RiffFxWide` | Enhanced stereo field with subtle chorus. |
| `RiffFxSlapback` / `RiffFxEcho` | Time-based repetitions. |


## Custom Recipes

If you need fine-grained control beyond presets, use these manual recipes as a starting point.

### Clean UI Click
Perfect for worksheet buttons or form interactions.
```vb
Public Sub ApplyCleanClick(ByVal v As Long)
    RiffVoiceSetFilter v, 0.92, 0.18  ' Low-pass and High-pass
    RiffVoiceVolume(v) = 0.6
End Sub
```

### Cinematic Impact
For dramatic transitions or heavy feedback.
```vb
Public Sub ApplyCinematicImpact(ByVal v As Long)
    RiffVoiceDistortion(v) = 1.4
    RiffVoiceEqBass(v) = 1.8
    RiffVoiceSetReverb v, 0.45, 0.8
    RiffVoiceCompressorThreshold(v) = 0.4
    RiffVoiceCompressorRatio(v) = 12.0
End Sub
```

### Retro Game Engine
Simulate an 8-bit or 12-bit sound chip.
```vb
Public Sub ApplyRetroChip(ByVal v As Long)
    RiffVoiceBitDepth(v) = 6
    RiffVoiceSampleRateReduction(v) = 3
    RiffVoiceEqTreble(v) = 0.5  ' Tame the aliasing brightness
End Sub
```

### Dream Sequence
A wash of stereo movement and long tails.
```vb
Public Sub ApplyDreamState(ByVal v As Long)
    RiffVoiceSetChorus v, 0.6, 1.2
    RiffVoiceSetReverb v, 0.7, 0.9
    RiffVoiceAutoPanRate(v) = 0.15
    RiffVoiceAutoPanDepth(v) = 0.8
    RiffVoiceStereoWidth(v) = 1.5
End Sub
```


## Synthesis Patterns

Use the new Pink and Brown noise generators for organic textures.

### Windy Ambience
```vb
Public Sub PlayWind()
    Dim v As Long: v = RiffPlayNoise(RiffWavePinkNoise)
    If v < 0 Then Exit Sub
    
    RiffVoiceVolume(v) = 0.2
    RiffVoiceAutoPanRate(v) = 0.05
    RiffVoiceAutoPanDepth(v) = 0.6
    
    ' Modulate the filter to simulate gusting
    RiffVoiceLowPass(v) = 0.3
End Sub
```

### Static / Vinyl Crackle
```vb
Public Sub PlayVinylCrackle()
    Dim v As Long: v = RiffPlayNoise(RiffWaveWhiteNoise)
    If v < 0 Then Exit Sub
    
    RiffVoiceVolume(v) = 0.05
    RiffVoiceBitDepth(v) = 3
    RiffVoiceHighPass(v) = 0.6
End Sub
```


## Pro Tip: Preset Application
Always apply a preset **before** making manual adjustments to volume or pan, as `RiffVoiceApplyPreset` resets the DSP chain to a neutral state before applying its settings.
