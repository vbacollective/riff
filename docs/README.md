# Documentation

This directory contains the long-form technical documentation for Riff. Start here when you need exact API behavior, integration constraints, or implementation details beyond the quick examples in the root README.

## Files

| File | Audience | Contents |
|---|---|---|
| [API_REFERENCE.md](API_REFERENCE.md) | Users integrating Riff into a VBA project | Public functions, voice properties, DSP controls, export helpers, meters, buses, and practical patterns. |
| [ARCHITECTURE.md](ARCHITECTURE.md) | Contributors and advanced users | Engine state, WASAPI initialization, Media Foundation decoding, timer thunk behavior, memory layout, DSP order, and shutdown. |
| [EFFECT_COOKBOOK.md](EFFECT_COOKBOOK.md) | Users building polished audio behavior | Copy-ready effect recipes for UI sounds, radio voice, ambience, lo-fi, sci-fi, music beds, and alerts. |
| [TROUBLESHOOTING.md](TROUBLESHOOTING.md) | Users diagnosing integration problems | No-sound checks, failed initialization, codec issues, voice exhaustion, shutdown safety, and path guidance. |
| [RELEASE_CHECKLIST.md](RELEASE_CHECKLIST.md) | Maintainers | Package, runtime, host, DSP, documentation, Git, and release-note checks. |

## Recommended Reading Order

1. Read the root [README](../README.md) for installation and common usage examples.
2. Read [API_REFERENCE.md](API_REFERENCE.md) when writing application code against Riff.
3. Read [EFFECT_COOKBOOK.md](EFFECT_COOKBOOK.md) when tuning effects or building reusable presets.
4. Read [TROUBLESHOOTING.md](TROUBLESHOOTING.md) if playback, loading, or shutdown behavior is not working as expected.
5. Read [ARCHITECTURE.md](ARCHITECTURE.md) before changing internals or debugging low-level host behavior.

## Quick API Pattern

Most Riff code follows this lifecycle:

```vb
If Not RiffOpen() Then
    MsgBox "Audio initialization failed."
    Exit Sub
End If

Dim bufferId As Long
bufferId = RiffLoad("C:\Audio\clip.wav")

If bufferId >= 0 Then
    Dim voiceId As Long
    voiceId = RiffPlay(bufferId)

    If voiceId >= 0 Then
        RiffVoiceVolume(voiceId) = 0.7
        RiffVoiceReverbMix(voiceId) = 0.25
    End If
End If
```

Always call `RiffClose` during workbook, document, or host shutdown. The engine owns native timer and COM resources that should be released deliberately.
