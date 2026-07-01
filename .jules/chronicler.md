# Chronicler's Journal

## Insights
- The `MsgReader` logic is complex, especially the OLE parsing parts. Adding JSDoc to these internal methods helps clarify the "magic" for future contributors and AI models.
- The project structure is simple but effective. Keeping documentation aligned with the code is crucial since there are no automated tests to verify behavior (yet).

## Recent Updates
- Created `ROADMAP.md` to track project progress.
- Documented `extractRecipients` in `js/msgreader.js` to explain the reconciliation logic between directory entries and display strings.
- Published a missing 'Quick Start' setup matrix to `README.md` and injected an 'Architectural Map' to detail the `mailto-link-generator/` directory layout.
