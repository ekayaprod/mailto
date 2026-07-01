# Technical Capability Matrix: MailTo Link Generator

## 1. Core Problem Definition & System Mechanics

**Systematic Friction Elimination:**
The system systematically eliminates the multi-domain collision between proprietary binary messaging formats (Microsoft OLE/MAPI Compound Documents) and web-native URI protocols. It bypasses backend parsing constraints, eliminating server-side ingestion delays, API rate limits, and network latency by executing a pure client-side binary extraction loop.

**Automated Constraint Resolution:**
It natively reconciles OLE directory structural drift and recipient type collisions (e.g., misaligned To/Cc routing) by executing a programmatic reconciliation between binary property structures and MAPI display strings. The system guarantees deterministic URL-encoded outputs, neutralizing special character corruption and malformed query strings that traditionally fracture inter-system communication links.

## 2. Granular Technical Architecture

| Domain | Component / Module | Execution Protocol / Specification |
| :--- | :--- | :--- |
| **Ingestion Pipeline** | `js/msgreader.js` | Parses `.msg` (OLE Compound Document) and `.eml` (MIME). Operates on `ArrayBuffer` via `DataView` for binary structural mapping. |
| **Binary Extraction** | OLE/FAT Traversal | Reads 512-byte sectors using FAT/MiniFAT sector chains. Navigates directory entries mapped by ID to extract start sectors, child IDs, and sibling relationships. |
| **Property Mappings** | MAPI Extraction | Extracts specific MAPI tags: `0x0037` (Subject), `0x1000` (Body), `0x1013` (HTML Body), `0x0E04` (Display To). Handles types: Integer32, Boolean, String/String8, Time, Binary. |
| **Decoding Engine** | `TextDecoder` / `DOMParser` | Handles UTF-16LE, Windows-1252, UTF-8. Decodes Quoted-Printable strings. Aggressive Regex + `DOMParser` strips HTML/CSS artifacts for plaintext normalization. |
| **State Controller** | `js/mailto.js` (App) | Central controller. Interfaces with UI, intercepts events, maintains nested tree object structures. |
| **Persistence Layer** | LocalStorage API | Serializes JSON structures to `localStorage` under the `mailto_generator_data` key. Constant time state retrieval. |
| **Protocol Generation** | MailTo Builder | `encodeURIComponent` wrapper for query string construction (`to`, `cc`, `bcc`, `subject`, `body`). Explicit `%0A` to `%0D%0A` transformation. |
| **Data Interchange** | CSV Serialization | Custom `parseCSVLine` handling quoted fields and escape sequences. Blobs and `URL.createObjectURL` for client-side file synthesis. |

## 3. Robustness & Operational Mandates

* **Runtime Integrity Checks:** The binary parser enforces strict boundary constraints (`entryOffset + entrySize > this.buffer.byteLength`) prior to sector reads, neutralizing buffer overrun exceptions on malformed files.
* **Error-Handling Routines:** Fallback decoders (`fatal: false` in `TextDecoder`) prevent thread panics on corrupted byte sequences, defaulting to `latin1` or generic ASCII traversal if strict UTF-8/16 parsing fails.
* **State-Validation Queries:** CSV ingestion mandates required header validation (`CONFIG.CSV_HEADERS`). Invalid structures abort injection, preserving `localStorage` state integrity.
* **Variable Collision Prevention:** Pseudo-random UUID synthesis (`crypto.randomUUID` with timestamp/math fallback) ensures distinct hierarchical node identities during DOM-tree mutation and data merges.
* **Asynchronous Resiliency:** Clipboard write operations and asynchronous imports are wrapped in granular `try/catch` blocks routed to non-blocking UI toast notifications, maintaining the execution loop context on localized failures.

## 4. Data Vectors & Quantifiable Impact

* **Transaction Processing Speed:** Zero network latency. The local extraction path (FileReader -> ArrayBuffer -> MsgReaderParser -> DOM Update) processes typical `.msg` files in milliseconds without backend dependency.
* **Data Replication Accuracy:** Guaranteed 100% deterministic mailto query string encoding, removing manual copy-paste mutation errors and URI corruption under complex load (e.g., massive CC lists).
* **Structural Maintainability:** The monorepo-ready encapsulated directory structure natively unifies the control logic and parsing library, ensuring immediate horizontal portability across static-site deployment vectors without a build chain constraint.
* **Volume Capacity:** Constrained only by browser V8 engine `localStorage` quotas (typically 5MB-10MB), sufficient for tens of thousands of flat or deeply nested templated routing structures.
