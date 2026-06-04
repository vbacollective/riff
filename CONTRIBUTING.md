# Contributing to Riff

Thank you for your interest in contributing to Riff! We welcome contributions that improve the engine's performance, expand its capabilities, or enhance its documentation.

As Riff is a high-performance audio engine written in VBA with native thunks, it has unique architectural constraints. Please review these guidelines before submitting a pull request.

## Development Workflow

1. **Fork the Repository**: Create your own fork of the project.
2. **Create a Branch**: Use a descriptive name for your branch (e.g., `feat/asynchronous-loading` or `fix/wasapi-underrun`).
3. **Implement Changes**: Ensure your code follows the established [Coding Standards](#coding-standards).
4. **Test Thoroughly**: Use the `examples/RiffShowcase.bas` and create new test procedures to verify your changes in both 32-bit and 64-bit Office.
5. **Submit a Pull Request**: Provide a clear description of your changes and why they are necessary.

## Coding Standards

### 1. Compatibility
- All code **must** be compatible with both 32-bit and 64-bit Office (VBA7).
- Use `#If VBA7` and `LongPtr` where necessary for pointer safety.

### 2. Performance
- Riff's audio path is performance-critical. Avoid creating VBA objects, using `Variant` types, or performing heavy string manipulation inside the `RiffTimerCallback` or its called functions.
- Prefer static arrays and UDTs over collections or classes for internal state.

### 3. Style & Conventions
- **Naming**: Use `PascalCase` for public functions and `camelCase` for private ones. Prefix public functions with `Riff`.
- **Comments**: Use Doxygen-style comments (`'/** ... */`) for public API documentation.
- **Explicit**: Always use `Option Explicit` and explicitly declare all variable types.

## Native Thunks

Riff uses machine-code thunks for the high-resolution audio callback. **Do not modify the hex strings in `InitThunks` unless you are an expert in x86/x64 assembly.** If you need to change the callback logic, ensure you update both the 32-bit and 64-bit opcodes and verify them across different Windows versions.

## Testing Requirements

Every pull request must be verified against:
- **Architecture**: 32-bit and 64-bit Excel/Word.
- **DSP**: No audible glitches or regressions in existing effects.
- **Stability**: No crashes during high-volume playback or project reset.

## Documentation

If you add a new public function or property, you **must** update:
1. `API_REFERENCE.md`
2. `ARCHITECTURE.md` (if internal logic changed)
3. `EFFECT_COOKBOOK.md` (if it enables new sound design possibilities)

## License

By contributing to Riff, you agree that your contributions will be licensed under the project's [MIT License](LICENSE).
