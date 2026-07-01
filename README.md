# ⚡ MailTo Link Generator

**1. Project Title & Brief Description**
The MailTo Link Generator is a zero-dependency, pure client-side web application I developed as a personal utility to instantly extract, generate, and organize `mailto` hyperlinks directly within my browser.

**2. The Operational Bottleneck**
I built this tool to eliminate the heavy manual processing and data entry risks associated with copying complex, fossilized Outlook templates (`.msg`, `.eml`, `.oft`) into web-native formats. Manually transcoding recipients, subject lines, and body text was a constant friction point in my workflow that often resulted in broken spacing, malformed URLs, and lost routing context.

**3. Tech Stack & Architecture**
- **Languages:** JavaScript (ES6 Modules), HTML5, CSS3
- **Execution:** 100% client-side logic (No backend server or dependencies)
- **State Management:** Browser `localStorage` (via the `mailto_generator_data` key)
- **Key Components:** Custom parser library (`msgreader.js`) for reading OLE/email file formats, and an application controller (`mailto.js`).

**4. Key Features & Workflow**
- **Raw File Parsing:** I can simply drag and drop `.msg`, `.eml`, or `.oft` files into the UI to automatically extract metadata (To, CC, BCC, Subject, Body) without manual copy-pasting.
- **Template Library:** I categorize my frequently used templates into folders, building a resilient personal asset library that can be backed up or restored via CSV.
- **Deterministic Link Generation:** The tool automatically generates perfectly URL-encoded HTML `mailto` links every time, neutralizing special characters.
- **Local Execution:** Bootstrapped via a simple local HTTP server (e.g., `python3 -m http.server 8000`) to bypass browser CORS policies for ES6 modules.

**5. Localized Impact**
This utility significantly optimizes my daily workflow by completely eliminating human error and manual transcription when constructing complex communication links. It drastically reduces the time I spend formatting outreach templates, ensuring that my personal repository of communication assets always produces bulletproof, cross-client compatible links.
