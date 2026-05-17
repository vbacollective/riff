# Riff Core Package

This directory contains **Riff.bas**, a monolithic, self-contained module that introduces a complete, high-performance audio engine to any Microsoft Office VBA project. The architecture encompasses WASAPI output, Media Foundation decoding, 32-voice polyphony, and a full per-voice studio DSP pipeline, functioning entirely without external dependencies or compiled DLLs.

> When imported, the VBA runtime compiles the module natively in memory. There is no build pipeline, no external binaries, and no complex packaging. The source code acts as the execution engine.

## Integration Guide

1. Open the VBA Integrated Development Environment (IDE).
2. Click **File > Import File...**
3. Select `Riff.bas` from this directory.
4. No additional configuration is required. The module relies strictly on native Win32 APIs, eliminating the need for external references or registry modifications.

After importing, the entire API surface becomes globally available.

> [!TIP]
> The comprehensive technical documentation is available in [`docs/API_REFERENCE.md`](../docs/API_REFERENCE.md).

## Architectural Capabilities

Riff exposes three independent operational layers that share a unified underlying engine and memory model.

**Static Buffer Playback:** The primary interface. Audio files are decoded from disk or from raw byte arrays directly into physical memory via `VirtualAlloc`, producing zero-latency playback on demand. Media Foundation's `IMFSourceReader` handles all decoding through COM vtable dispatch, supporting any format the host operating system natively understands, including WAV, MP3, AAC, WMA, and FLAC. Once loaded, a buffer can be played simultaneously across multiple voices without duplication.

**Oscillator Synthesis:** Provides direct waveform generation without any audio file. Sine, Square, Sawtooth, Triangle, and Noise waveforms are produced mathematically through per-voice phase accumulators and routed through the exact same DSP pipeline as buffer playback. This is the foundation for sound effects, alert tones, and generative audio entirely inside a VBA project.

**Studio DSP Pipeline:** Every voice, whether sourced from a buffer or an oscillator, passes through an independent per-voice processing chain on every audio frame. The chain covers Bitcrusher, Sample Rate Reduction, Distortion, Low Pass, High Pass, 3-Band EQ, Ring Modulator, Tremolo, Stereo Width, Flanger, Chorus, Delay, Reverb, Compressor, AutoPan, Volume, Pan, Bus routing, and frame-accurate Fade. All modulated effects run continuous LFO phase accumulators that persist across frames, producing smooth and click-free modulation at any depth or rate.

### Unified Engine Features

All three modes benefit from a shared foundational architecture:

- WASAPI Shared Mode output with automatic detection of 32-bit float and 16-bit integer device formats.
- Physical memory allocation for audio buffers via `VirtualAlloc`, bypassing VBA's heap entirely and preventing garbage collection interference.
- A contiguous 1D ring buffer of `32 * 192000` floats shared across all voices for Chorus, Flanger, Delay, and Reverb state, avoiding the Column-Major 2D array wipe issue inherent to VBA's memory layout.
- A runtime-compiled x86/x64 assembly thunk driving the audio callback through Win32 `SetTimer`, with an `EbMode` liveness guard that calls `KillTimer` and exits safely if the VBA runtime is reset, preventing host application crashes.
- 8 audio buses with independent volume control for grouping voices by category such as music, sound effects, or ambient layers.
- Peak amplitude metering at both voice and master level for real-time VU instrumentation.
- Full 32-bit and 64-bit compatibility managed transparently through `#If VBA7` and `#If Win64` conditional compilation across all API declarations and assembly opcodes.

## Voice Handle Management

Every successful `RiffPlay` or `RiffPlayOscillator` call allocates one of the 32 available polyphonic voices and returns a unique numeric handle. Passing this handle to subsequent property assignments and method calls targets that specific voice, allowing independent control of volume, pitch, pan, loop state, position, bus routing, and the entire DSP chain per voice. When a voice finishes playing or is stopped, its slot is released back to the pool automatically.

Buffers follow the same model. `RiffLoad` and `RiffLoadFromMemory` return a buffer handle from a pool of 64 slots. A single buffer handle can feed multiple simultaneous voices without any memory duplication.

## Error Handling

Riff uses a defensive programming model throughout the public API. Every function validates the engine initialization state, handle bounds, and buffer activity before executing. Invalid calls return silently or return -1 without affecting engine state. The `MagicCookie` field in the internal `RiffContext` structure acts as a runtime sentinel, causing the audio callback to abort immediately if the engine has been torn down or the context has been corrupted, preventing undefined behavior in the audio thread.
