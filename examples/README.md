# Examples

This directory contains ready-to-import VBA examples for Riff.

| File | Purpose |
|---|---|
| `RiffShowcase.bas` | Interactive demos that make the major DSP effects easy to hear. |

## Importing the Showcase

1. Import `package/Riff.bas` first.
2. Import `examples/RiffShowcase.bas`.
3. Run `Showcase` from the VBA editor or from a button in the host application.

The showcase module starts Riff if needed, runs short demonstrations for reverb, delay, chorus, flanger, filters, EQ, distortion, bitcrusher, tremolo, auto-pan, ring modulation, pitch, and fades, then shuts the engine down at the end.

## Example: Excel Button Playback

Use this when a worksheet button should play a short sound effect.

```vb
Option Explicit

Private clickBuffer As Long
Private clickAudioReady As Boolean

Public Sub LoadButtonAudio()
    If Not RiffOpen() Then
        MsgBox "Audio could not be initialized.", vbCritical
        Exit Sub
    End If

    clickAudioReady = False
    clickBuffer = RiffLoad("C:\Audio\button-click.wav")

    If clickBuffer < 0 Then
        MsgBox "Button audio could not be loaded.", vbExclamation
    Else
        clickAudioReady = True
    End If
End Sub

Public Sub Button_ClickSound()
    If Not clickAudioReady Then Exit Sub

    Dim voiceId As Long
    voiceId = RiffPlay(clickBuffer)

    If voiceId >= 0 Then
        RiffVoiceVolume(voiceId) = 0.65
        RiffVoiceHighPass(voiceId) = 0.2
    End If
End Sub
```

## Example: Ambient Loop With Fade

This pattern is useful for menus, background ambience, or presentation audio that should enter and leave smoothly.

```vb
Option Explicit

Private ambientBuffer As Long
Private ambientVoice As Long
Private ambientPlaying As Boolean

Public Sub StartAmbientLoop()
    If Not RiffOpen() Then Exit Sub

    ambientPlaying = False
    ambientBuffer = RiffLoad("C:\Audio\ambient-loop.mp3")
    If ambientBuffer < 0 Then Exit Sub

    ambientVoice = RiffPlay(ambientBuffer)

    If ambientVoice >= 0 Then
        RiffVoiceLoop(ambientVoice) = True
        RiffVoiceVolume(ambientVoice) = 0.4
        RiffVoiceReverbMix(ambientVoice) = 0.35
        RiffFadeIn ambientVoice, 1.5
        ambientPlaying = True
    End If
End Sub

Public Sub StopAmbientLoop()
    If ambientPlaying Then
        RiffFadeOut ambientVoice, 1!
        ambientPlaying = False
    End If
End Sub
```

## Example: Loop Region

Use loop regions when an audio file has a non-repeating intro followed by a seamless loop.

```vb
Public Sub PlayMusicWithIntro()
    Dim bufferId As Long
    bufferId = RiffLoad("C:\Audio\music-with-intro.wav")
    If bufferId < 0 Then Exit Sub

    Dim voiceId As Long
    voiceId = RiffPlay(bufferId)
    If voiceId < 0 Then Exit Sub

    RiffVoiceLoop(voiceId) = True
    RiffSetLoopRegionSec voiceId, 4.2, 38.7
End Sub
```

## Example: VU Meter Polling

Peak meters can be read from a form timer, worksheet timer, or host-specific update loop.

```vb
Private Sub Timer_Tick()
    Dim peakLeft As Single
    Dim peakRight As Single

    RiffMasterGetPeak peakLeft, peakRight

    Dim percent As Long
    percent = CLng(((peakLeft + peakRight) * 0.5) * 100)

    VUBar.Width = percent * 2

    If percent > 80 Then
        VUBar.BackColor = RGB(255, 50, 50)
    ElseIf percent > 50 Then
        VUBar.BackColor = RGB(255, 200, 0)
    Else
        VUBar.BackColor = RGB(50, 200, 50)
    End If
End Sub
```

## Example: Embedded Byte Array Loading

`RiffLoadFromMemory` lets an application load encoded audio from a byte array. This is useful when you store small assets in a worksheet, custom document part, resource table, or generated array.

```vb
Public Sub PlayEmbeddedSound(ByRef encodedAudio() As Byte)
    If Not RiffOpen() Then Exit Sub

    Dim bufferId As Long
    bufferId = RiffLoadFromMemory(encodedAudio)
    If bufferId < 0 Then Exit Sub

    RiffPlay bufferId
End Sub
```

## Cleanup

Always shut Riff down when the host file closes:

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RiffClose
End Sub
```

For non-Excel hosts, use the equivalent close or shutdown event.
