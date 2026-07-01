# ⚡ MailTo Link Generator

[![status: active](https://img.shields.io/badge/status-active-brightgreen)](#)
[![execution: client-side](https://img.shields.io/badge/execution-client--side-blue)](#)

A zero-dependency, pure client-side engine for extracting, generating, and organizing `mailto` hyperlinks. No backend. Strict browser execution.

## 🏗️ Architecture

The active codebase is completely encapsulated within the `mailto-link-generator/` directory to support monorepo integration. There is no `package.json` or build system. The application functions entirely as a static site.

- **State Persistence:** Data is persisted purely via the browser's `localStorage` using the `mailto_generator_data` key.
- **Execution Paradigm:** 100% client-side logic.
- **Application Topology:**

  | Route/File                              | Domain     | Structural Purpose                                                                                                     |
  | :-------------------------------------- | :--------- | :--------------------------------------------------------------------------------------------------------------------- |
  | `mailto-link-generator/index.html`      | Entry      | The core application entry point orchestrating the UI.                                                                 |
  | `mailto-link-generator/js/mailto.js`    | Controller | The application controller managing logic, state, and user interactions.                                               |
  | `mailto-link-generator/js/msgreader.js` | Library    | A standalone parser library that extracts metadata (To, CC, BCC, Subject, Body) from `.eml`, `.msg`, and `.oft` files. |

## 🚀 Usage

Since the codebase relies entirely on static assets, there are no package managers or installation commands to run. However, due to strict browser security policies (CORS) surrounding ES6 Modules (`<script type="module">`), the file import feature requires a local HTTP server to function correctly.

1. Navigate into the encapsulated directory:
   ```bash
   cd mailto-link-generator
   ```
2. Boot up a local server. For example:
   ```bash
   python3 -m http.server 8000
   ```
   _or_
   ```bash
   npx serve .
   ```
3. Open `http://localhost:8000` in a modern web browser to execute. Do not open `index.html` directly from the file system.

## ⚙️ Capabilities

- **Raw Extraction:** Instantly parse `.msg`, `.eml`, and `.oft` files by dragging them into the upload zone.
  - _Benefit:_ Eliminates manual copy-pasting of complex email structures.
  - _Use Case:_ Rapidly convert fossilized Outlook templates into web-native formats without losing recipient routing or subject line context.
- **Library System:** Categorize templates into folders and manage state via CSV import and export.
  - _Benefit:_ Transforms ephemeral mailto links into a resilient, version-controlled asset library safely stored in local persistence.
  - _Use Case:_ Organize campaign-specific outreach, manage departmental contact lists, and seamlessly transfer workspace data across environments.
- **Deterministic Output:** Generates perfectly URL-encoded HTML links every time.
  - _Benefit:_ Guarantees cross-client rendering integrity by neutralizing special characters and broken spacing.
  - _Use Case:_ Deploy bulletproof mailto links in enterprise portals or static sites where malformed URLs would result in dead communication channels.
