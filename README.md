# ⚡ MailTo Link Generator

[![status: active](https://img.shields.io/badge/status-active-brightgreen)](#)
[![execution: client-side](https://img.shields.io/badge/execution-client--side-blue)](#)

A zero-dependency, pure client-side engine for extracting, generating, and organizing `mailto` hyperlinks. No backend. Strict browser execution.

## 🏗️ Architecture

The active codebase is completely encapsulated within the `mailto-link-generator/` directory to support monorepo integration. There is no `package.json` or build system. The application functions entirely as a static site.

* **State Persistence:** Data is persisted purely via the browser's `localStorage` using the `mailto_generator_data` key.
* **Execution Paradigm:** 100% client-side logic.
* **Application Topology:**
  * `mailto-link-generator/index.html` — The core application entry point.
  * `mailto-link-generator/js/mailto.js` — The application controller managing logic and user interactions.
  * `mailto-link-generator/js/msgreader.js` — A standalone library that parses `.eml`, `.msg`, and `.oft` files to extract metadata (To, CC, BCC, Subject, Body).

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
   *or*
   ```bash
   npx serve .
   ```
3. Open `http://localhost:8000` in a modern web browser to execute. Do not open `index.html` directly from the file system.

## ⚙️ Capabilities

* **Raw Extraction:** Instantly parse `.msg`, `.eml`, and `.oft` files by dragging them into the upload zone.
* **Library System:** Categorize templates into folders and manage state via CSV import and export.
* **Deterministic Output:** Generates perfectly URL-encoded HTML links every time.
