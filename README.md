# ⚡ MailTo Link Generator

## 1. Overview
The **MailTo Link Generator** is a zero-dependency, pure client-side web utility engineered specifically to streamline my personal communication workflows. Built as an encapsulated local application, it instantly extracts metadata from complex email files (`.msg`, `.eml`, `.oft`) and generates perfectly encoded, cross-client compatible `mailto:` hyperlinks. It operates entirely within the browser, requiring no backend or external server dependencies.

## 2. The Operational Catalyst
In my daily workflow, I routinely deal with fossilized Outlook templates and complex email structures that require manual extraction and conversion into web-native communication channels. The process of opening an `.msg` or `.eml` file, copying the subject line, body text, and complex recipient arrays (To, CC, BCC), and then manually URL-encoding them to create a functional `mailto:` link was a tedious, error-prone manual nightmare. This repetitive technical friction caused significant bottlenecks and increased the risk of data entry errors. I needed an autonomous, zero-touch solution to bypass this manual copy-pasting and rapidly convert these files into robust hyperlinks, strictly for my own localized use.

## 3. Under the Hood (Technical Architecture)
This project is architected as a static client-side application, utilizing vanilla HTML, CSS, and ES6 JavaScript modules. The execution pattern relies entirely on modern browser APIs to handle binary parsing and state management without a server.

- **Execution Paradigm:** Purely client-side execution. The entry point (`index.html`) orchestrates the UI, while application logic is split between a core controller (`mailto.js`) and a standalone OLE/MIME parser library (`msgreader.js`).
- **Binary Parsing & Extraction:** When an email file is dropped into the UI, the `FileReader` API reads it as an `ArrayBuffer`. The `MsgReaderParser` processes `.msg` files (OLE Compound Documents) by reading the File Allocation Table (FAT/MiniFAT) and directory entries using `DataView`. It extracts binary properties using MAPI Property Tags (e.g., `PROP_ID_SUBJECT`, `PROP_ID_BODY`) and decodes strings using `TextDecoder` (supporting UTF-8, UTF-16LE, and Windows-1252).
- **MIME Fallback:** For standard `.eml` files, it utilizes a custom regex-based MIME parser to extract headers and decode Quoted-Printable body text.
- **Data Sanitization:** The `DOMParser` API is leveraged to aggressively strip out HTML tags, CSS artifacts (especially Outlook-specific styles), and scripts, normalizing rich text into clean plain text for URL embedding.
- **State Persistence:** Parsed templates and custom folder structures are managed in memory and persisted purely via the browser's `localStorage` (`mailto_generator_data`), effectively creating a resilient local asset library.
- **Asynchronous & DOM Operations:** UI interactions are managed via a centralized controller pattern. Real-time preview updates are governed by debounced input listeners, ensuring the DOM updates performantly while generating deterministic, URL-encoded `mailto:` links. Data import/export operations use `Promise`-wrapped file reading APIs for non-blocking execution.

## 4. Robustness & Integrity
To ensure zero-touch execution and prevent data corruption within my personal workflow, the codebase is fortified with explicit fail-safes and error-handling routines:

- **Encoding Fallbacks:** The `MsgReader` incorporates robust fallback mechanisms for `TextDecoder`. If a specific character set fails or is unsupported by the environment, it gracefully degrades to manual byte-by-byte decoding routines to salvage string data.
- **Input Validation & Sanitization:** All CSV imports are strictly validated against required headers (`['name', 'path', 'to', 'cc', 'bcc', 'subject', 'body']`). The application actively rejects malformed data, presenting specific error lists rather than corrupting the local storage state.
- **Deterministic URL Encoding:** The core `MailTo.build` function guarantees rendering integrity by safely URL-encoding all special characters, replacing line breaks with `%0D%0A`, and ensuring valid query parameter structure, neutralizing the risk of broken links.
- **DOM Injection Protection:** Custom `Utils.escapeHTML` sanitization is applied across all dynamic list rendering and modal injections to prevent XSS and DOM disruption, even when parsing malformed or unexpected email data.

## 5. Localized ROI (Impact)
Developed exclusively to optimize my individual efficiency, this tool has fundamentally eliminated a daily operational bottleneck.

- **Time Reduction:** Condensed a multi-minute manual process of opening desktop email clients, extracting text, and URL-encoding characters into a 3-second drag-and-drop operation.
- **Error Elimination:** Fully automated the translation of complex recipient arrays and special character encoding, resulting in a 100% reduction in broken links and malformed email drafts.
- **Throughput Increase:** By building a searchable, localized template library with CSV export capabilities, I've created a version-controlled asset system that dramatically accelerates my daily communication output and workflow consistency.
