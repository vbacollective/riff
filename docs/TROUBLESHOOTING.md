# Troubleshooting

This guide covers common Riff integration issues in Office/VBA hosts.

## Quick Health Check

Run this first in a new standard module after importing `package/Riff.bas`:

```vb
Public Sub RiffHealthCheck()
    If Not RiffOpen() Then
        MsgBox "RiffOpen failed. Check audio device availability.", vbCritical
        Exit Sub
    End If

    Dim voiceId As Long
    voiceId = RiffPlayOscillator(0, 440)

    If voiceId < 0 Then
        MsgBox "Riff opened, but no voice slot was available.", vbExclamation
    Else
        RiffVoiceVolume(voiceId) = 0.2
        RiffFadeOut voiceId, 1!
    End If
End Sub
```

If this plays a short tone, the engine, timer, output device, and basic voice path are working.

## RiffOpen Returns False

Likely causes:

- No default Windows audio output device is available.
- The audio device is disabled, disconnected, or controlled by a failed driver session.
- Media Foundation initialization failed.
- The Office host is running in a restricted environment that blocks required native calls.

What to try:

- Confirm Windows can play audio in another application.
- Switch the default output device in Windows Sound settings.
- Fully close and reopen the Office host.
- Run the minimal oscillator health check instead of loading an audio file.
- Test in a new blank workbook or document to rule out project-specific state.

## No Sound, But No Error

Check these first:

- `RiffOpen` must return `True` before playback.
- `RiffLoad` must return a buffer handle greater than or equal to `0`.
- `RiffPlay` must return a voice handle greater than or equal to `0`.
- `RiffMasterVolume`, `RiffBusVolume`, and `RiffVoiceVolume` must be above `0`.
- Your Windows output device must not be muted.

Minimal trace:

```vb
Public Sub DebugPlayback()
    Debug.Print "Open:", RiffOpen()

    Dim bufferId As Long
    bufferId = RiffLoad("C:\Audio\test.wav")
    Debug.Print "Buffer:", bufferId

    Dim voiceId As Long
    voiceId = RiffPlay(bufferId)
    Debug.Print "Voice:", voiceId

    If voiceId >= 0 Then
        RiffMasterVolume = 1!
        RiffBusVolume(0) = 1!
        RiffVoiceVolume(voiceId) = 0.8
    End If
End Sub
```

## Audio File Does Not Load

`RiffLoad` uses Windows Media Foundation. Format support depends on the codecs available to the operating system.

What to check:

- The path is absolute or valid relative to the host process.
- The file exists and is not locked.
- The file can be played by a Windows media application.
- Try WAV first to separate file decoding issues from engine issues.
- Avoid paths that depend on the current workbook directory unless you explicitly build them.

Safer path pattern:

```vb
Dim audioPath As String
audioPath = ThisWorkbook.Path & "\audio\click.wav"

Dim bufferId As Long
bufferId = RiffLoad(audioPath)
```

## RiffPlay Returns -1

Riff has 32 voice slots. `RiffPlay` and `RiffPlayOscillator` return `-1` when every slot is active.

What to try:

- Lower the number of overlapping sounds.
- Stop finished or unneeded sounds with `RiffStop`.
- Use `RiffStopAll` before starting a new demo section.
- Avoid accidentally looping many short sounds.
- Fade out long ambience voices instead of starting new copies repeatedly.

## Office Crashes After Reset or Close

Riff uses a native Win32 timer thunk to drive audio. Always close the engine before resetting the VBA project or closing the host file.

Recommended Excel cleanup:

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RiffClose
End Sub
```

During development, also run this manually before pressing the VBA reset button:

```vb
Public Sub StopAudioEngine()
    RiffClose
End Sub
```

## MessageBox Makes a Ding During Demos

Windows may play a notification sound for some `MsgBox` icon styles. If you are demonstrating audio effects, avoid icon flags in normal prompts.

```vb
If MsgBox("Run this demo?", vbOKCancel, "Riff Demo") = vbOK Then
    Demo_ReverbRoomSize
End If
```

## Sound Is Too Loud or Distorted

Real-time effects can stack gain quickly, especially reverb, delay feedback, distortion, EQ boosts, and multiple simultaneous voices.

What to try:

- Lower `RiffVoiceVolume` before increasing effects.
- Keep `RiffVoiceReverbMix` and `RiffVoiceDelayFeedback` moderate.
- Lower the music bus instead of lowering every music voice.
- Use fewer overlapping voices for dense scenes.

Practical starting point:

```vb
RiffMasterVolume = 0.8
RiffBusVolume(0) = 0.45 ' Music
RiffBusVolume(1) = 0.85 ' SFX
```

## Effects Are Not Audible

Some effects need the right source material:

- Low-pass and high-pass are easiest to hear on bright sounds such as saw waves, cymbals, or noise.
- Chorus and flanger are easiest to hear on sustained pads.
- Delay is easiest to hear on short pulses.
- Ring modulation is clearer on simple tones or voice clips.
- Bitcrusher is most obvious on clean tones or short effects.

Use the showcase module in `examples/RiffShowcase.bas` when tuning effect presets.

## 32-bit and 64-bit Office

Riff supports both x86 and x64 Office through conditional compilation. If you see compile errors:

- Confirm you imported the complete `package/Riff.bas` file.
- Confirm the host supports VBA.
- Avoid editing pointer declarations unless you are intentionally changing platform support.
- Test with a blank workbook to rule out name conflicts from other modules.

## Relative Paths Behave Unexpectedly

Office hosts do not always use the workbook or document folder as the process current directory. Prefer explicit paths.

Excel example:

```vb
Dim audioPath As String
audioPath = ThisWorkbook.Path & "\audio\theme.mp3"
```

Access and Word have different host-specific path APIs. Build paths deliberately for the host you are using.

## Safe Development Checklist

- Import `package/Riff.bas` as a standard module.
- Call `RiffOpen` once before first playback.
- Check every buffer and voice handle for `>= 0`.
- Keep a manual `StopAudioEngine` macro available during development.
- Call `RiffClose` before VBA reset and host shutdown.
- Start effect-heavy scenes with conservative volume.
- Test a sine oscillator first when diagnosing output problems.
