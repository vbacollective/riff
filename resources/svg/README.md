# SVG Assets

This directory contains SVG icons used throughout the Riff documentation to represent host compatibility, project status, and platform support.

## Icon Index

| File | Represents |
|:---|:---|
| `ms-excel.svg` | Microsoft Excel compatibility |
| `ms-word.svg` | Microsoft Word compatibility |
| `ms-powerpoint.svg` | Microsoft PowerPoint compatibility |
| `ms-access.svg` | Microsoft Access compatibility |
| `ms-outlook.svg` | Microsoft Outlook compatibility |
| `ms-office.svg` | General Microsoft Office / VBA support |
| `windows.svg` | Windows platform requirement |
| `planning.svg` | Roadmap items / In-progress features |
| `completed.svg` | Implemented features / Success states |

## Maintenance Guidelines

- **Format:** Icons should be optimized, standard SVG 1.1.
- **Style:** Maintain a consistent "flat" aesthetic that aligns with modern Office/Windows branding.
- **Accessibility:** Use descriptive filenames and ensure the SVG `viewBox` is correctly set for consistent scaling in Markdown.
- **Cleaning:** Use tools like `svgo` to remove unnecessary metadata and minimize file size before committing.
