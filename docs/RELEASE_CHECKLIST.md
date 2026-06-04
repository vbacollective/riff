# Release Checklist

Use this checklist to ensure every Riff release meets professional quality standards before publishing.

---

## 1. Package Integrity
- [ ] `package/Riff.bas` is the only production module and has no external dependencies.
- [ ] No `Tools > References` are required in a clean VBA project.
- [ ] Code compiles without errors in both 32-bit and 64-bit Office hosts.
- [ ] `Public` API names match the documentation in `API_REFERENCE.md`.

## 2. Runtime Smoke Tests
- [ ] `RiffOpen` successfully acquires the default WASAPI device.
- [ ] `RiffPlayOscillator` produces a clean tone without artifacts.
- [ ] `RiffLoad` successfully decodes WAV and MP3 assets.
- [ ] `RiffClose` shuts down all resources and restores timer resolution.

## 3. DSP & Pipeline Verification
- [ ] **Modulation:** Verify Chorus, Flanger, and Tremolo sweep correctly.
- [ ] **Spatial:** Confirm Reverb tail and Delay repetitions are audible.
- [ ] **Dynamics:** Test Compressor gain reduction and Distortion clipping.
- [ ] **Filters:** Test Low-Pass, High-Pass, and 3-Band EQ tonal changes.
- [ ] **Presets:** Verify `RiffVoiceApplyPreset` correctly configures the DSP matrix.

## 4. Host & Platform Coverage
- [ ] Tested on **Excel** (32-bit & 64-bit).
- [ ] Tested on **Windows 10** or **Windows 11**.
- [ ] (Optional) Verified on Word, PowerPoint, or Access.

## 5. Documentation Review
- [ ] `README.md` version number and roadmap are current.
- [ ] All `docs/*.md` files are free of broken links and outdated snippets.
- [ ] `EFFECT_COOKBOOK.md` recipes are tested and copy-ready.

## 6. Git & Asset Management
- [ ] No temporary files (e.g., `.tmp`, `.log`) are staged for commit.
- [ ] `resources/` folder contains only required, optimized assets.
- [ ] Commit message follows project standards (professional and descriptive).

---

## 7. Version Tagging
- [ ] Update version constant (if any) in `Riff.bas`.
- [ ] Push tags to the repository.
- [ ] Verify CI/CD pipeline successfully builds release assets.
