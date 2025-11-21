MailTo Link Generator

A client-side tool for generating, organizing, and managing mailto hyperlinks.

Features

File Import: Extracts To, CC, BCC, Subject, and Body from .msg, .eml, and .oft email files.

Library: Folder-based organization for email templates.

Data Management: CSV import and export functionality.

Output: Generates URL-encoded HTML links.

Processing: Client-side execution without server dependencies.

Installation

Download the source files.

Ensure the following directory structure:

/
├── index.html
├── style.css
└── js/
    ├── mailto.js
    └── msgreader.js


Open index.html in a modern web browser.

Usage

Creating Templates

Populate fields in the Editor panel.

Select Generate Link to create the mailto string.

Select Save to Library to persist the template.

Importing Files

Drag a supported email file into the upload zone.

Verify populated fields.

Edit and save as a template.

Managing Library

Folders: Create folders via the sidebar header button.

Move: Organize items using the arrow icon.

Edit: Load templates into the editor using the pencil icon.

Technical Notes

Persistence: Uses browser localStorage.

Encoding: Handles standard URL encoding for special characters.

Dependencies: No external libraries required.
