# Release Checklist

Use this checklist before tagging or publishing a Riff release.

## Package Integrity

- Confirm `package/Riff.bas` is the only production module.
- Import `package/Riff.bas` into a blank VBA project.
- Compile the VBA project.
- Confirm no external **Tools > References** entries are required.
- Confirm public API changes are documented in [API_REFERENCE.md](API_REFERENCE.md).
- Confirm architecture changes are documented in [ARCHITECTURE.md](ARCHITECTURE.md).

## Runtime Smoke Tests

Run these on at least one Windows machine with a working output device:

- `RiffOpen` returns `True`.
- `RiffPlayOscillator(0, 440)` plays a sine tone.
- `RiffFadeOut` fades the oscillator without clicks.
- `RiffLoad` loads a WAV file.
- `RiffPlay` plays the loaded buffer.
- A looping voice can be stopped with `RiffStop` or `RiffStopAll`.
- `RiffClose` shuts down cleanly.

## Host Coverage

Recommended test matrix:

| Host | 32-bit | 64-bit |
|---|---:|---:|
| Excel | Test when available | Test when available |
| Word | Optional | Optional |
| PowerPoint | Optional | Optional |
| Access | Optional | Optional |

At minimum, test the release on the Office architecture most likely to be used by your target audience.

## Audio Feature Checks

- File decoding works with WAV.
- File decoding works with at least one compressed format supported by the test machine, such as MP3.
- `RiffLoadFromMemory` loads a valid encoded byte array.
- `RiffExportBufferWav` writes a playable WAV file.
- `RiffRenderOscillatorWav` writes a playable oscillator WAV file.
- Master peak metering returns changing values during playback.
- Bus volume affects only voices assigned to that bus.

## DSP Spot Checks

Use `examples/RiffShowcase.bas` or equivalent manual tests to confirm:

- Reverb tail is audible.
- Delay feedback repeats.
- Chorus and flanger modulation are audible.
- Low-pass and high-pass filters change tone.
- EQ changes bass, mid, and treble balance.
- Distortion and bitcrusher are audible.
- Tremolo and auto-pan modulate over time.
- Ring modulation creates metallic tones.
- Pitch changes playback speed or oscillator pitch.

## Documentation Checks

- Root [README](../README.md) points to the current package path.
- [docs/README.md](README.md) includes all major documentation files.
- [examples/README.md](../examples/README.md) describes how to import and run examples.
- Usage snippets use public API names that exist in `package/Riff.bas`.
- Paths in examples are clearly placeholders or host-derived paths.
- No copyrighted audio files are required for examples.

## Git Checks

- Review `git status --short`.
- Review `git diff --stat`.
- Confirm generated or local-only files are not included.
- Confirm no accidental edits were made to unrelated files.

## Release Notes

Include:

- Version number.
- New features.
- Bug fixes.
- Public API additions or breaking changes.
- Documentation changes.
- Known limitations.
- Tested Office and Windows versions.
