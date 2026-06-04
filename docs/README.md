# Documentation

Welcome to the Riff technical documentation. This directory contains detailed guides on how to integrate, configure, and extend the engine.

## Resource Index

| File | Audience | Description |
|:---|:---|:---|
| [**API Reference**](API_REFERENCE.md) | Developers | Exhaustive guide to every public function, property, and enum. |
| [**Architecture**](ARCHITECTURE.md) | Advanced | Internal implementation details: WASAPI, Thunks, and DSP logic. |
| [**Effect Cookbook**](EFFECT_COOKBOOK.md) | Designers | Practical recipes for radio voices, spatial reverbs, and UI beeps. |
| [**Troubleshooting**](TROUBLESHOOTING.md) | All | Solutions for common setup, playback, and stability issues. |
| [**Release Checklist**](RELEASE_CHECKLIST.md) | Maintainers | Quality assurance steps required before every release. |

## Quick Integration Pattern

Riff is designed for simple, synchronous usage in standard VBA modules.

```vb
' Basic Playback Procedure
Public Sub RunSoundTest()
    If Not RiffOpen() Then Exit Sub

    Dim buf As Long: buf = RiffLoad("C:\Assets\intro.mp3")
    Dim v As Long:   v = RiffPlay(buf)

    If v >= 0 Then
        RiffVoiceVolume(v) = 0.75
        RiffVoiceReverbMix(v) = 0.2
    End If
End Sub
```

## Lifecycle Management

To ensure host stability (especially in Excel), the engine **must** be shut down before the workbook or host application closes. This releases native timers and COM interfaces that VBA cannot clean up automatically.

```vb
' Place in ThisWorkbook or equivalent host shutdown event
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RiffClose
End Sub
```
