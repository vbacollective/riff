# Troubleshooting Guide

This document provides solutions for common integration, playback, performance, editor-state, and stability issues encountered when using Riff.

This version also covers the newer gameplay-stability and performance builds, including burst-safe playback, one-shot procedural noise, lazy temporal-buffer preparation, VBE-safe timer behavior, and Stop/Reset-safe cleanup.

## Common Diagnostic Path

If you are not hearing audio or encountering errors, follow these steps:

1. **Verify Initialization:** Ensure `RiffOpen` is called once and returns `True`.
2. **Check Error Code:** Immediately after a failure, inspect `Debug.Print RiffLastError`.
3. **Verify Paths:** Ensure file paths passed to `RiffLoad` are absolute and accessible.
4. **Check Device:** Confirm the Windows default playback device is active and not muted by another application.
5. **Check Runtime Counters:** Inspect active voice counts to confirm that voices are being created and released correctly.
6. **Check Adaptive Stats:** Inspect underruns, padding, frames available, and frames written when debugging dropouts.
7. **Check Editor State:** If the VBE looks stuck in `Running`, use the editor cleanup section below.

Useful quick snapshot:

```vb
Public Sub RiffQuickSnapshot()
    Debug.Print "Initialized:        "; RiffIsInitialized
    Debug.Print "Last Error:         "; RiffLastError
    Debug.Print "Active Voices:      "; RiffActiveVoiceCount()
    Debug.Print "SFX Voices:         "; RiffBusVoiceCount(RiffBusSfx)
    Debug.Print "Music Voices:       "; RiffBusVoiceCount(RiffBusMusic)
    Debug.Print "Auto Suspend:       "; RiffAutoSuspendTimer
    Debug.Print "Adaptive Queue ms:  "; RiffAdaptiveQueueMs
    Debug.Print "Underruns:          "; RiffUnderrunCount
    Debug.Print "Padding Frames:     "; RiffLastPaddingFrames
    Debug.Print "Available Frames:   "; RiffLastFramesAvailable
    Debug.Print "Written Frames:     "; RiffLastFramesWritten
End Sub
```

## Specific Issues

### 1. No Sound During Playback

- **Engine not open:** Call `RiffOpen` once before loading or playing audio.
- **Auto-Suspend:** Riff may be in suspend mode to save CPU. Call `RiffWake` or ensure `RiffAutoSuspendTimer` is configured as desired.
- **Bus Volume:** Check if the voice is routed to a bus, such as `RiffBusMusic`, that is currently muted or has `RiffBusVolume = 0`.
- **Master Volume:** Ensure `RiffMasterVolume` is set to an audible level. The normal default is `1.0`.
- **Voice Exhaustion:** Riff supports 32 simultaneous voices. If you spawn many one-shots without stopping them, you may hit the limit. Use `RiffStopAll` to clear the pool for testing.
- **Voice cap stealing:** If burst caps are enabled, a very fast repeated sound may steal/reuse older voices instead of creating unlimited overlap. This is expected and prevents accumulation.
- **Invalid buffer handle:** `RiffPlay` returns `-1` if the buffer was not loaded or was unloaded.

Debug pattern:

```vb
Public Sub TestBasicTone()
    Dim v As Long

    If Not RiffOpen() Then
        Debug.Print "RiffOpen failed:", RiffLastError
        Exit Sub
    End If

    v = RiffPlayOscillator(RiffWaveSine, 440!, RiffBusUi, 0.2!, 0!, 0.25!)

    Debug.Print "Voice:", v
    Debug.Print "Active:", RiffActiveVoiceCount()
    Debug.Print "Error:", RiffLastError
End Sub
```

### 2. Host Crashes on Project Reset

- **Dangling Timer:** If you press the Reset/Stop button in the VBA Editor without calling `RiffClose`, a native timer callback may still be pending. Modern Riff builds include VBE and timer-id guards to kill stale callbacks safely, but the safest normal shutdown path is still `RiffClose`.
- **Stop/Reset-safe cleanup:** If an older timer/callback survives a VBA reset, call `RiffEditorEmergencyStop` after the editor becomes usable again.
- **Thunk Corruption:** Ensure `package/Riff.bas` is not modified manually, as the machine-code opcodes are sensitive to offset changes.
- **Manual editing risk:** Avoid changing private thunk/timer code unless you are intentionally rebuilding the low-level callback path.

Recommended emergency command:

```vb
Public Sub EmergencyAudioReset()
    RiffEditorEmergencyStop
End Sub
```

Recommended normal cleanup:

```vb
Public Sub ShutdownAudio()
    RiffClose
End Sub
```

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

### 3. Loading Fails (`RiffLoad` returns -1)

- **Format Support:** Riff uses Windows Media Foundation. If a specific file, such as `.flac` or `.aac`, will not load, ensure the necessary codecs are installed on the system. Some Windows editions require Media Feature Packs.
- **File Access:** VBA may struggle with networked paths or cloud-synced folders such as OneDrive if the file is not available locally. Use "Always keep on this device" for synced assets.
- **Buffer Pool Full:** Riff supports 64 loaded buffers. Use `RiffUnload` to free slots if you are dynamically loading many assets.
- **Relative Path Error:** Prefer absolute paths based on `ActivePresentation.Path`, `ThisWorkbook.Path`, or a known asset directory.
- **Locked File:** Some editors or sync tools may temporarily lock files during export/copy.

Debug pattern:

```vb
Public Function LoadChecked(ByVal path As String) As Long
    Dim h As Long

    h = RiffLoad(path)

    If h = -1 Then
        Debug.Print "RiffLoad failed"
        Debug.Print "Path: "; path
        Debug.Print "Error:"; RiffLastError
    Else
        Debug.Print "Loaded buffer:"; h
        Debug.Print "Duration sec:"; RiffBufferDurationSec(h)
    End If

    LoadChecked = h
End Function
```

### 4. Audio is Distorted or "Crackly"

- **Clipping:** If many loud voices are played simultaneously, the master bus may clip. Lower `RiffMasterVolume`, lower bus volume, or use master soft clipping/limiting.
- **CPU Saturation:** Extremely complex DSP chains on many simultaneous voices may stress older CPUs. Disable unused effects by setting depth/mix to `0`.
- **Overprocessed Presets:** Avoid stacking heavy voice preset + heavy bus preset + heavy master preset unless the effect is intentional.
- **Feedback too high:** Delay, flanger, or reverb feedback values that are too aggressive can create harsh build-up. Modern presets are sanitized, but manual values can still be extreme.
- **Low-quality source:** Distorted source audio remains distorted after decoding. Test with a known clean WAV.

Safer SFX-heavy mix:

```vb
Public Sub ApplySfxSafety()
    RiffMasterVolume = 0.85!
    RiffSoftClipEnabled = True
    RiffMasterApplyPreset RiffMasterFxSoftLimiter, 0.65!
    RiffMaxVoicesPerBuffer = 4
    RiffMaxVoicesPerBus = 18
End Sub
```

### 5. Fast Repeated Sounds Make the Game Slow

This was a major gameplay issue in older builds. The usual cause is not a literal VBA call stack leak. It is normally too many voices or too much DSP state being created, kept alive, or reset during rapid one-shot playback.

Modern Riff builds include:

- per-buffer voice caps;
- per-bus voice caps;
- voice stealing;
- finite noise one-shots;
- optional oscillator duration;
- micro attack/release ramps;
- safer voice cleanup before reuse;
- lazy temporal-buffer preparation for delay, reverb, chorus, and flanger;
- faster dry `RiffPlay`;
- faster preset setup.

Recommended gameplay defaults:

```vb
Public Sub AudioGameplayDefaults()
    RiffOpen

    RiffVoiceStealingEnabled = True
    RiffMaxVoicesPerBuffer = 4
    RiffMaxVoicesPerBus = 18

    ' Keeps audio responsive during active gameplay.
    ' Modern VBE-safe builds still auto-release the editor when idle.
    RiffAutoSuspendTimer = False
End Sub
```

Check for accumulation:

```vb
Public Sub PrintSfxPressure(ByVal stepBuffer As Long)
    Debug.Print "Active voices:", RiffActiveVoiceCount()
    Debug.Print "SFX voices:", RiffBusVoiceCount(RiffBusSfx)
    Debug.Print "Step voices:", RiffBufferVoiceCount(stepBuffer, RiffBusSfx)
End Sub
```

Expected behavior:

```txt
During burst: active voices may rise briefly.
After finite one-shots finish: active voices should return to 0.
Repeated SFX should not climb forever.
```

### 6. `RiffPlayNoise` Does Not Loop Anymore

In the gameplay-stable build, `RiffPlayNoise` is a finite one-shot by default. This prevents accidental infinite procedural noise voices when developers use noise as a short SFX.

Use this for short procedural hits:

```vb
RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.15!, 0!, 0.04!
```

Use this for continuous ambience:

```vb
Dim rain As Long

rain = RiffPlayNoiseLoop(RiffWavePinkNoise, RiffBusMusic, 0.05!, 0!)
RiffVoiceSetFilterHz rain, 2600!, 120!
RiffVoiceApplyPreset rain, RiffFxRain, 0.6!

' Later:
RiffFadeOut rain, 2!
```

You can still request continuous noise through `durationSec:=0`, but `RiffPlayNoiseLoop` is clearer:

```vb
Dim wind As Long
wind = RiffPlayNoise(RiffWavePinkNoise, RiffBusMusic, 0.04!, 0!, 0!)
```

### 7. Oscillator Beeps Keep Playing Forever

`RiffPlayOscillator` creates a continuous oscillator when `durationSec` is omitted or `0`.

Continuous oscillator:

```vb
Dim hum As Long
hum = RiffPlayOscillator(RiffWaveSine, 55!, RiffBusSfx, 0.12!)
```

Finite beep:

```vb
RiffPlayOscillator RiffWaveSquare, 880!, RiffBusUi, 0.2!, 0!, 0.08!
```

For UI and game SFX, prefer the finite form. This avoids accidental active voices during gameplay.

### 8. VBE IntelliSense Stops Working or Title Bar Flickers Between Running and Normal

This is usually caused by a timer/callback keeping VBA in a `Running`-like state. Earlier performance builds kept the WASAPI path warm with silence to reduce underruns between rapid SFX, but that could keep the VBE busy when the project was idle.

Modern VBE-safe builds solve this by:

- keeping the audio path warm during short active bursts;
- auto-suspending after idle;
- killing stale timer callbacks if the VBE reset variables;
- killing mismatched old timer ids;
- exposing `RiffEditorEmergencyStop` for manual cleanup.

Recommended pattern during development:

```vb
Public Sub DevAudioInit()
    RiffOpen
    RiffAutoSuspendTimer = True
End Sub
```

Recommended pattern during active gameplay:

```vb
Public Sub GameAudioInit()
    RiffOpen
    RiffAutoSuspendTimer = False
End Sub
```

If IntelliSense still appears stuck:

```vb
Public Sub FixVBEAudioState()
    RiffEditorEmergencyStop
End Sub
```

Then try the IntelliSense again.

### 9. Clicking the VBE Stop/Reset Button Leaves the Editor Weird

The VBE Stop/Reset button can reset VBA variables before the Riff shutdown path runs. In older builds, this could leave a timer callback orphaned. The Stop-safe build handles this inside the callback by killing stale timer ids, but `RiffEditorEmergencyStop` remains available as a manual escape hatch.

Symptoms:

- VBE title bar flickers;
- IntelliSense does not appear;
- project appears to alternate between idle and running;
- you need to press Pause and Stop repeatedly.

Fix:

```vb
RiffEditorEmergencyStop
```

Best practice:

```vb
Public Sub SafeStopAudioBeforeEditing()
    RiffStopAll
    RiffClose
End Sub
```

Use `RiffEditorEmergencyStop` for editor recovery. Use `RiffClose` for normal shutdown.

### 10. Underrun Count Keeps Rising

`RiffUnderrunCount` rises when Riff detects that the audio endpoint is at risk of running dry. A few underruns during editor work, breakpoint activity, window switching, or heavy host load can be normal. A rapidly rising value during normal gameplay indicates that the render loop is not keeping up or the endpoint is being drained between sounds.

Check:

```vb
Debug.Print "Underruns:", RiffUnderrunCount
Debug.Print "Queue:", RiffAdaptiveQueueMs
Debug.Print "Padding:", RiffLastPaddingFrames
Debug.Print "Available:", RiffLastFramesAvailable
Debug.Print "Written:", RiffLastFramesWritten
```

Recommended fixes:

- Keep `RiffAutoSuspendTimer = False` during active gameplay.
- Avoid breakpoints while testing real-time audio.
- Reduce simultaneous heavy DSP chains.
- Lower SFX overlap using `RiffMaxVoicesPerBuffer` and `RiffMaxVoicesPerBus`.
- Use `RiffMasterFxSoftLimiter` rather than extreme per-voice compression on many voices.
- Avoid playing dozens of long reverb/delay voices at once.

Gameplay pattern:

```vb
Public Sub EnterGameplay()
    RiffOpen
    RiffResetAdaptiveStats
    RiffAutoSuspendTimer = False
End Sub

Public Sub LeaveGameplay()
    RiffAutoSuspendTimer = True
End Sub
```

### 11. Presets Feel Slow When Applied to Many One-Shots

Older builds cleared large temporal buffers immediately when applying effects. This made preset-heavy one-shots slower than necessary.

Modern builds use lazy temporal-buffer preparation:

```txt
Dry preset / EQ / compressor / radio-style filter:
    no large temporal buffer clear during setup.

Delay / reverb / chorus / flanger:
    temporal buffer is prepared only when the effect is actually needed.
```

Recommended usage:

```vb
Dim v As Long
v = RiffPlay(sndClick, RiffBusUi, False, 0.55!, 0!)

If v >= 0 Then
    RiffVoiceApplyPreset v, RiffFxGameBoy, 0.55!
End If
```

For very frequent UI sounds, consider persistent bus presets:

```vb
Public Sub EnterRetroUi()
    RiffBusApplyPreset RiffBusUi, RiffFxGameBoy, 0.45!, True, False
End Sub
```

Then play normally:

```vb
RiffPlay sndClick, RiffBusUi, False, 0.55!, 0!
```

### 12. A Preset Sounds Wrong or Too Extreme

Presets are clamped/sanitized after application in modern builds, but preset intensity still matters. Some presets are intentionally dramatic.

Try:

```vb
RiffVoiceApplyPreset v, RiffFxRadio, 0.45!
RiffVoiceApplyPreset v, RiffFxUnderwater, 0.35!
RiffVoiceApplyPreset v, RiffFxCinematicBoom, 0.55!
```

Instead of:

```vb
RiffVoiceApplyPreset v, RiffFxRadio, 1!
RiffVoiceApplyPreset v, RiffFxUnderwater, 1!
RiffVoiceApplyPreset v, RiffFxCinematicBoom, 1!
```

If you use bus and master presets together, lower both amounts:

```vb
RiffBusApplyPreset RiffBusVoice, RiffFxRadio, 0.55!
RiffMasterApplyPreset RiffMasterFxRadio, 0.25!
```

### 13. Filter Values Are Confusing

Riff has two filter styles:

- normalized filters: `RiffVoiceSetFilter`, `RiffVoiceLowPass`, `RiffVoiceHighPass`;
- frequency filters in Hz: `RiffVoiceSetFilterHz`.

Use Hz filters when designing real-world sound bands:

```vb
' Telephone/radio band:
RiffVoiceSetFilterHz v, 3000!, 300!

' Muffled wall/underwater:
RiffVoiceSetFilterHz v, 900!, 0!

' Remove sub-rumble:
RiffVoiceSetFilterHz v, 0!, 80!
```

Use normalized filters when you want quick abstract tonal control:

```vb
RiffVoiceSetFilter v, 0.45!, 0.05!
```

### 14. Audio Delays Slightly After Being Idle

If Riff has auto-suspended its timer during idle, the next sound may need to wake the render path. This is good for editor stability but can make the first sound after a long idle feel slightly less immediate.

For active gameplay:

```vb
RiffAutoSuspendTimer = False
```

For editor/menu/idle-heavy usage:

```vb
RiffAutoSuspendTimer = True
```

Modern VBE-safe builds still self-release after idle, so `False` is safer than in older builds, but `True` is still the best choice when actively editing code.

### 15. Audio Keeps Playing After Slide Show Ends

Call `RiffClose` or `RiffStopAll` from the host shutdown path.

PowerPoint:

```vb
Public Sub OnSlideShowTerminate(ByVal Pres As Presentation)
    RiffClose
End Sub
```

If you cannot wire the host event cleanly, add a manual macro:

```vb
Public Sub StopAllAudioNow()
    RiffStopAll
    RiffClose
End Sub
```

### 16. Memory Appears to Increase Slightly During First Test

Small memory changes during first playback, first noise generation, or first preset use can be normal. Windows, VBA, Media Foundation, WASAPI, and internal scratch buffers may commit memory lazily.

A real leak usually looks like memory increasing every repeated test and never stabilizing.

Test pattern:

```vb
Public Sub MemoryLeakSanityCheck()
    Dim i As Long

    RiffOpen
    RiffResetAdaptiveStats

    For i = 1 To 1000
        RiffPlayNoise RiffWaveWhiteNoise, RiffBusSfx, 0.05!, 0!, 0.03!
        If (i Mod 25) = 0 Then DoEvents
    Next i

    DoEvents
    Debug.Print "Active voices:", RiffActiveVoiceCount()
    Debug.Print "Underruns:", RiffUnderrunCount
End Sub
```

After the one-shots finish, `RiffActiveVoiceCount()` should return to `0`.

## Diagnostic Snippet

Use this procedure to verify the engine's health in the Immediate Window (`Ctrl + G`):

```vb
Public Sub RiffDiagnostic()
    Dim v As Long

    Debug.Print "Initialized: "; RiffIsInitialized
    Debug.Print "Last Error:  "; RiffLastError
    Debug.Print "Active Voices: "; RiffActiveVoiceCount()

    If Not RiffIsInitialized Then
        If RiffOpen() Then
            Debug.Print "Engine started successfully."
        Else
            Debug.Print "Engine failed to start. Error: "; RiffLastError
            Exit Sub
        End If
    End If

    RiffResetAdaptiveStats

    v = RiffPlayOscillator(RiffWaveSine, 440!, RiffBusUi, 0.2!, 0!, 0.25!)

    Debug.Print "Test voice: "; v
    Debug.Print "Active after play: "; RiffActiveVoiceCount()

    DoEvents

    Debug.Print "Underruns: "; RiffUnderrunCount
    Debug.Print "Padding: "; RiffLastPaddingFrames
    Debug.Print "Available: "; RiffLastFramesAvailable
    Debug.Print "Written: "; RiffLastFramesWritten
End Sub
```

## Development Workflow Tips

### Use VBE-Safe Mode While Editing

When actively editing code, prefer:

```vb
RiffAutoSuspendTimer = True
```

This lets Riff release the editor more aggressively when idle.

### Use Warm Playback Mode During Gameplay

When running the actual game loop, prefer:

```vb
RiffAutoSuspendTimer = False
```

This keeps rapid SFX responsive and reduces underruns between short sounds. Modern builds still include idle release behavior to prevent VBE lockups.

### Always Keep a Manual Cleanup Macro

```vb
Public Sub AudioPanic()
    RiffEditorEmergencyStop
End Sub
```

Keep this macro available while developing experimental audio/game systems.

### Avoid Sleep-Only Waits in Audio Tests

If a benchmark or test uses `Sleep` without `DoEvents`, the VBA host may not pump timer messages normally. This can make voices appear active longer than expected.

Prefer a pumped wait:

```vb
Public Sub WaitPumped(ByVal seconds As Single)
    Dim t As Single
    t = Timer

    Do While Timer - t < seconds
        DoEvents
    Loop
End Sub
```

## Performance Expectations

These numbers are approximate and depend on the Office host, CPU, audio driver, Windows version, and whether the VBE is open.

Recent optimized builds measured approximately:

```txt
Dry RiffPlay:        ~11–13 us/call
Noise one-shot:      ~13–15 us/call
Oscillator one-shot: ~13–15 us/call
Preset/DSP setup:    ~20 us/call
```

A healthy gameplay-loop benchmark should generally show:

```txt
Failed handles:      0
Active after wait:   0
Game loop underruns: 0 or very low
Final active voices: 0
Memory delta:        near zero / stable
Missed frame budget: 0 or very low
```

## Final Recovery Checklist

If Riff behaves strangely while developing:

1. Run `RiffEditorEmergencyStop`.
2. Run `RiffClose`.
3. Press the VBE Stop button once.
4. Confirm `RiffActiveVoiceCount()` is `0`.
5. Run `RiffResetAdaptiveStats`.
6. Reopen with `RiffOpen`.
7. Retest with a simple finite oscillator.
