# Troubleshooting Guide

This document provides solutions for common integration, playback, and stability issues encountered when using Riff.

---

## Common Diagnostic Path

If you are not hearing audio or encountering errors, follow these steps:

1. **Verify Initialization:** Ensure `RiffOpen` is called once and returns `True`.
2. **Check Error Code:** Immediately after a failure, inspect `Debug.Print RiffLastError`.
3. **Verify Paths:** Ensure file paths passed to `RiffLoad` are absolute and accessible.
4. **Check Device:** Confirm the Windows default playback device is active and not muted by another application.

---

## Specific Issues

### 1. No Sound During Playback
- **Auto-Suspend:** Riff may be in suspend mode to save CPU. Call `RiffWake` or ensure `RiffAutoSuspendTimer` is configured as desired.
- **Bus Volume:** Check if the voice is routed to a bus (e.g., `RiffBusMusic`) that is currently muted or has `RiffBusVolume = 0`.
- **Master Volume:** Ensure `RiffMasterVolume` is set to a audible level (default is 1.0).
- **Voice Exhaustion:** Riff supports 32 simultaneous voices. If you spawn many one-shots without stopping them, you may hit the limit. Use `RiffStopAll` to clear the pool for testing.

### 2. Host Crashes on Project Reset
- **Dangling Timer:** If you press the "Reset" button in the VBA Editor without calling `RiffClose`, the native timer thunk may still be active. While Riff includes `EbMode` guards to prevent this, the safest practice is to always call `RiffClose` in the `Workbook_BeforeClose` event.
- **Thunk Corruption:** Ensure `package/Riff.bas` is not modified manually, as the machine-code opcodes are sensitive to offset changes.

### 3. Loading Fails (`RiffLoad` returns -1)
- **Format Support:** Riff uses Windows Media Foundation. If a specific file (e.g., a `.flac` or `.aac`) won't load, ensure the necessary codecs are installed on the system (usually included in "Media Feature Packs").
- **File Access:** VBA may struggle with networked paths or cloud-synced folders (OneDrive) if the file is not "Always keep on this device."
- **Buffer Pool Full:** Riff supports 64 loaded buffers. Use `RiffUnload` to free slots if you are dynamically loading many assets.

### 4. Audio is Distorted or "Crackly"
- **Clipping:** If many loud voices are played simultaneously, the master bus may clip. Lower `RiffMasterVolume` or use the `RiffVoiceCompressor` on individual voices.
- **CPU Saturation:** While Riff is highly optimized, extremely complex DSP chains on all 32 voices simultaneously may stress older CPUs. Disable unused effects by setting their depth/mix to 0.

---

## Diagnostic Snippet

Use this procedure to verify the engine's health in the Immediate Window (`Ctrl + G`):

```vb
Public Sub RiffDiagnostic()
    Debug.Print "Initialized: "; RiffIsInitialized
    Debug.Print "Last Error:  "; RiffLastError
    Debug.Print "Active Voices: "; ' (Iterate handles to check)
    
    If Not RiffIsInitialized Then
        If RiffOpen() Then
            Debug.Print "Engine started successfully."
            RiffPlayOscillator RiffWaveSine, 440
            RiffFadeOut 0, 1.0
        Else
            Debug.Print "Engine failed to start. Error: "; RiffLastError
        End If
    End If
End Sub
```
