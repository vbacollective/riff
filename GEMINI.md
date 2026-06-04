# Project Riff: AI Development Mandates

This document provides foundational context and instructions for AI agents (like Gemini CLI) working within the Riff repository. Adhering to these mandates ensures architectural integrity, performance, and safety.

---

## Technical Context

Riff is a high-performance audio engine written in **VBA7** that uses:
- **Native WASAPI** for low-latency playback.
- **Media Foundation** for asynchronous-ready decoding.
- **Machine-Code Thunks** (x86/x64) for high-resolution timing.
- **Zero Dependencies**: No DLLs or external references.

## Critical Mandates

### 1. Host Stability & Shutdown
Always ensure that `RiffClose` is emphasized in documentation and examples. Because Riff runs a native timer callback, a project reset or host shutdown without calling `RiffClose` will cause the application to crash.

### 2. VBA7 & Pointer Safety
All code must be compatible with both 32-bit and 64-bit Office.
- Use `LongPtr` for memory addresses and handles.
- Use `#If VBA7` for conditional compilation.
- Never assume a 4-byte pointer size.

### 3. Audio Path Performance
The `RiffTimerCallback` (and any function it calls, such as the DSP pipeline) is performance-critical.
- **NEVER** use `Variant` types in the audio path.
- **NEVER** instantiate VBA classes or create objects (e.g., `New Dictionary`) in the render loop.
- **NEVER** perform string manipulation or `Debug.Print` inside the callback.
- Prefer `Single` and `Long` primitives for arithmetic.

### 4. Thunk Integrity
The hex strings in `package/Riff.bas` (InitThunks) are compiled machine code. **DO NOT** modify them unless explicitly directed with verified opcodes. Any change to the thunk logic requires updates to both 32-bit and 64-bit strings.

## Coding Standards

- **Public API**: Prefix all public engine functions with `Riff`.
- **Error Handling**: Use `RiffSetLastError` to report failures. Do not use `On Error Resume Next` to mask engine failures.
- **Documentation**: Maintain Doxygen-style comments (`'/** ... */`) for all public symbols.

## Documentation Priorities

When updating documentation:
1. **API Reference**: Keep technically precise. Always include parameter ranges and return types.
2. **Cookbook**: Provide "copy-ready" snippets that work in standard modules.
3. **Architecture**: Update diagrams (Mermaid) and UDT tables if internal logic changes.

---

*This file is a foundational mandate for AI interactions. Prioritize these instructions over general defaults.*
