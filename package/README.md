# Package

This directory contains the production module that users import into their VBA projects.

| File | Purpose |
|---|---|
| `Riff.bas` | The complete Riff audio engine. |

## Importing Riff

1. Open the target Office file.
2. Open the VBA editor with `Alt+F11`.
3. Choose **File > Import File...**.
4. Select `package/Riff.bas`.
5. Save the host file as a macro-enabled workbook, document, database, or presentation.

Riff does not require references under **Tools > References**. The module declares and calls the required Windows APIs itself.

## First Run

Use this pattern before loading or playing audio:

```vb
Public Sub InitializeAudio()
    If Not RiffOpen() Then
        MsgBox "Riff could not start the audio engine.", vbCritical
        Exit Sub
    End If

    RiffMasterVolume = 0.8
End Sub
```

Release native resources when the host closes:

```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RiffClose
End Sub
```

For Word, PowerPoint, Access, or Outlook, place the same `RiffClose` call in the closest available shutdown or close event for that host.

## Packaging Notes

- Keep `Riff.bas` as a standard module.
- Do not split the module into VBA classes unless the engine architecture is intentionally redesigned.
- Do not add external DLL, ActiveX, typelib, or registry requirements.
- Keep public API changes documented in [docs/API_REFERENCE.md](../docs/API_REFERENCE.md).
- Keep low-level implementation changes documented in [docs/ARCHITECTURE.md](../docs/ARCHITECTURE.md).

## Related Files

- [Examples](../examples/README.md)
- [API Reference](../docs/API_REFERENCE.md)
- [Architecture](../docs/ARCHITECTURE.md)
