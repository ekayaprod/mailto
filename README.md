# **MailTo Link Generator**

A client-side tool for generating, organizing, and managing mailto hyperlinks.

## **Features**

* **File Import**: Extracts To, CC, BCC, Subject, and Body from .msg, .eml, and .oft email files.  
* **Library**: Folder-based organization for email templates.  
* **Data Management**: CSV import and export functionality.  
* **Output**: Generates URL-encoded HTML links.  
* **Processing**: Client-side execution without server dependencies.

## **Installation**

1. Download the source files.  
2. Ensure the following directory structure:  
   /  
   ├── index.html
   ├── css/ 
      ├── style.css  
   └── js/  
       ├── mailto.js  
       └── msgreader.js

4. Open index.html in a modern web browser.

## **Usage**

### **Creating Templates**

1. Populate fields in the **Editor** panel.  
2. Select **Generate Link** to create the mailto string.  
3. Select **Save to Library** to persist the template.

### **Importing Files**

1. Drag a supported email file into the upload zone.  
2. Verify populated fields.  
3. Edit and save as a template.

### **Managing Library**

* **Folders**: Create folders via the sidebar header button.  
* **Move**: Organize items using the arrow icon.  
* **Edit**: Load templates into the editor using the pencil icon.

## **Technical Notes**

* **Persistence**: Uses browser localStorage.  
* **Encoding**: Handles standard URL encoding for special characters.  
* **Dependencies**: No external libraries required.
