# Resources

This directory contains visual assets used by repository documentation. These files are not required at runtime after `package/Riff.bas` has been imported into a VBA project.

## Contents

| Path | Purpose |
|---|---|
| `logo.png` | Project logo used by the root README. |
| [`svg/`](svg/) | Microsoft Office, Windows, status, and documentation icons. |

## Usage Guidelines

- Keep runtime VBA code out of this directory.
- Prefer stable asset filenames because README files link to these paths directly.
- Use PNG or SVG assets for documentation graphics.
- Add attribution notes when an asset is derived from another project or public source.

## Example README Reference

```md
![Riff logo](resources/logo.png)
```

For SVG icons under `resources/svg/`, use relative paths from the README that includes them.
