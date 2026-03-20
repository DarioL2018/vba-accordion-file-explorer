# VBA Accordion Template Explorer for Microsoft Word

A lightweight, zero-dependency **VBA UserForm** designed to navigate template directories directly within Microsoft Word. This project features a custom-built "Accordion" (Tree-View) UI that ensures 100% compatibility across both **32-bit and 64-bit Office environments**.

## 🚀 The Problem: The "TreeView" Compatibility Gap
Standard VBA development often relies on the `MSCOMCTL.OCX` TreeView control. However, this control is notorious for crashing on 64-bit Office versions and frequently requires administrative rights for registration. 

**This solution bypasses those issues** by simulating a hierarchical tree structure using a standard `ListBox` and recursive logic, making it a robust, "plug-and-play" tool for corporate environments.

## ✨ Key Features
*   **Dynamic Accordion UI:** Inline expand/collapse logic for folders using `[+]` and `[-]` indicators.
*   **Cross-Bitness Compatible:** Works flawlessly on Word 2013, 2016, 2019, 2021, and Office 365 (both 32/64-bit).
*   **External Configuration:** Managed via a `config.ini` file, allowing users to update root paths without touching the VBA source code.
*   **Recursive Folder Loading:** Efficiently handles deep directory structures.
*   **Smart Filtering:** Automatically filters for Word documents (`.doc`, `.docx`, `.dot`, `.dotm`) and ignores temporary Office files (`~$`).

## 📂 Project Structure
```text
├── bin/
│   └── TemplateExplorer.docm       # Ready-to-use Word Document with Macro
├── src/
│   ├── frmExplorer.frm            # UserForm Source (Exported)
│   ├── frmExplorer.frx            # UserForm Binary (Exported)
│   └── Module1.bas                # Launcher and Auto-run logic
├── config.ini                     # Configuration file for root path
└── README.md
```

## 🛠️ Setup & Installation

1.  **Placement:** Place the `.docm` (or `.dotm`) file and the `config.ini` in the same directory.
2.  **Configuration:** Edit the `config.ini` file to set your desired template root path. Example:
    ```ini
    [Settings]
    RootFolder=C:\Templates\
    ```
3.  **VBA Import (Manual):** If you wish to integrate this into an existing project:
    *   Open Microsoft Word and press `ALT + F11` to open the VBA Editor.
    *   Right-click on your Project (in the Project Explorer) -> **Import File...**
    *   Select the `.frm` and `.bas` files from the `src/` folder of this repository.
4.  **Activation:** By default, the form is set to launch automatically via the `Document_Open` event in `ThisDocument`.

## 💻 Technical Implementation

As a Software Engineering solution, this project prioritizes **stability and maintainability**:

*   **Recursive Navigation:** Uses a robust recursive function to map the file system into the ListBox without depth limits.
*   **Hidden Metadata Columns:** The ListBox is configured with 3 columns (`Widths: 300;0;0`). Columns 2 and 3 store the absolute file paths and the item type (FOLDER/FILE) invisibly.
*   **State Management:** Utilizes a `Scripting.Dictionary` object to track which folders are expanded, ensuring the "tree" state is preserved during UI refreshes.
*   **Optimized Event Handling:** 
    *   `ListBox_Click`: Handles the logic for expanding/collapsing folders (Accordion toggle).
    *   `ListBox_DblClick`: Dedicated to executing/opening the selected Word template.
*   **Win32 API Integration:** Uses `kernel32` calls (`GetPrivateProfileStringA`) for efficient and fast reading of the `.ini` configuration file.

## 📄 License

This project is open-source and available under the **MIT License**. Feel free to use, modify, and distribute it in your own professional projects.

---
Developed by **Dario Larenas** - *Software Engineer*