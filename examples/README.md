# Examples

This directory contains ready-to-import VBA examples and interactive demos for the Riff audio engine.

## Available Modules

| File | Purpose |
|:---|:---|
| [**RiffShowcase.bas**](RiffShowcase.bas) | An interactive "Studio" demo. Use this to hear all DSP effects, oscillators, and mixing capabilities in real-time. |

## How to Run Examples

1. **Import the Engine:** Ensure `package/Riff.bas` is already imported into your VBA project.
2. **Import the Example:** Press `Alt + F11`, then **File > Import File...** and select the `.bas` file from this directory.
3. **Run a Procedure:** Open the imported module and press `F5` on any `Public Sub`.

## Common Implementation Pattern

When building your own features, use this standard pattern for safe startup and clean audio logic.

```vb
Public Sub Example_LoopingAmbience()
    ' 1. Initialize hardware (safe to call multiple times)
    If Not RiffOpen() Then Exit Sub

    ' 2. Load the asset (returns a buffer handle)
    Dim buf As Long
    buf = RiffLoad("C:\YourProject\Sounds\forest_loop.mp3")

    If buf < 0 Then
        Debug.Print "Failed to load audio asset."
        Exit Sub
    End If

    ' 3. Play the asset (returns a voice handle)
    Dim v As Long
    v = RiffPlay(buf)

    If v >= 0 Then
        ' 4. Configure DSP and playback
        RiffVoiceLoop(v) = True
        RiffVoiceVolume(v) = 0.5
        RiffVoiceLowPass(v) = 0.8
        RiffVoiceReverbMix(v) = 0.2
        
        ' 5. Apply a smooth fade-in
        RiffFadeIn v, 2.5
    End If
End Sub
```

## Host Shutdown

Always include a call to `RiffClose` in your host's shutdown event (e.g., `Workbook_BeforeClose` in Excel) to ensure all native resources are released.
