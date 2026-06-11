# Agent Instructions

## UI/UX Consistency

- Keep UI and UX consistent throughout the project.
- Reuse the established layout, spacing, typography, colors, modal patterns, and interaction behavior.
- When adding new screens or controls, match the existing viewer shell rather than introducing a new visual language.
- Use icons instead of text for buttons where possible, but keep labels when an icon alone would be ambiguous.

## Persistent Data

- Use the configured MongoDB instance for any persistent application data.
- Do not add new file-based persistence for durable state unless it is explicitly requested.
- The only file-based persistence currently allowed is the platform audit log in `example/logs/activity.log`.
