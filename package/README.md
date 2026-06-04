# Package

This directory contains the production-ready Riff module. This is the only file you need to integrate Riff into your own VBA project.

## Module Index

| File | Purpose |
|:---|:---|
| [**Riff.bas**](Riff.bas) | The complete, standalone audio engine. Includes WASAPI output, Media Foundation decoding, and the DSP pipeline. |

## Installation

Integrating Riff into a new or existing Office project takes only a few seconds:

1. **Download:** Get the latest `Riff.bas` from the [releases](https://github.com/vbacollective/riff/releases) page.
2. **Import:**
   - Open your Excel, Word, or Access file.
   - Press `Alt + F11` to open the VBA Editor.
   - Go to **File > Import File...** (or press `Ctrl + M`).
   - Select `Riff.bas`.
3. **Save:** Save your document as a macro-enabled file (e.g., `.xlsm`, `.docm`, `.accdb`).

## Configuration & Dependencies

- **No References:** Riff does not require any entries in **Tools > References**. It uses late-bound COM calls and direct API declarations.
- **No DLLs:** All required APIs (`kernel32`, `ole32`, `mfplat`, etc.) are built into every modern Windows installation (Windows 7 and later).
- **Architecture:** Riff automatically detects whether you are running 32-bit or 64-bit Office and adjusts its internal pointer logic and machine-code thunks accordingly.

## Maintenance

To update Riff in an existing project:
1. Right-click `Riff` in the Project Explorer and select **Remove Riff...**.
2. Select **No** when asked to export it.
3. Import the new version of `Riff.bas`.

## Related Resources

- [**API Reference**](../docs/API_REFERENCE.md)
- [**Architecture Design**](../docs/ARCHITECTURE.md)
- [**Interactive Examples**](../examples/README.md)
