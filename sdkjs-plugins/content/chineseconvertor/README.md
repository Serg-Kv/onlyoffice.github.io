# Chinese Converter — Simplified ⇄ Traditional & Pinyin (ONLYOFFICE Plugin)

A lightweight and practical ONLYOFFICE plugin for Chinese text conversion. It enables instant conversion between Simplified and Traditional Chinese, and supports adding or removing Pinyin annotations — directly inside the editor.

> Current version: see `CHANGELOG.md`

---

## 🚀 Feature Overview

- **Simplified → Traditional**
  - Convert selected Simplified Chinese text to Traditional Chinese.

- **Traditional → Simplified**
  - Convert selected Traditional Chinese text to Simplified Chinese.

- **Add Pinyin**
  - Automatically generates Pinyin for each Chinese character in the selected text.
  - **Inline** inserts pinyin after each character in parentheses.
  - **Reading line** puts the readings on a separate line above the original text without padding or per-character alignment.
    - Example: 你好 → 你(nǐ)好(hǎo)
  - Useful for language learners, teaching materials, and pronunciation guidance.

- **Remove Pinyin**
  - Removes generated pinyin. In documents, select the Chinese paragraph to remove or replace its reading line. For older annotations with alignment spaces, select the whole Chinese paragraph.
  - In spreadsheets, readings and Chinese text stay within each original cell. Reading line enables wrapping; formula cells are skipped.

- **Right-Click Quick Access**
  - Access conversion features directly from the context menu after selecting text.

- **Multi-Editor Support**
  - Works in:
    - Document Editor
    - Spreadsheet Editor

---

## 🖼️ UI Preview

📌 **Left Side Panel**  
_The main plugin interface displayed in the left window panel._  
![Left Panel](resources/store/screenshots/screenshot1.png)

📋 **Inline layout**

_Choose Inline to put readings next to each character._

![Inline layout](resources/store/screenshots/screenshot2.png)

---

## 🛠️ How to Use

1. Clone the repository and deploy it to an ONLYOFFICE environment that supports plugins.
2. Open a document or spreadsheet file.
3. Select the Chinese text you want to modify.
4. Use one of the following methods:
   - Open **Chinese Converter** from the toolbar (left side panel).
   - Right-click the selected text and choose the desired action.
5. Choose one of the following actions:
   - `Simplified → Traditional`
   - `Traditional → Simplified`
   - `Add Pinyin`
   - `Remove Pinyin`

💡 Tip: You can also right-click after selecting text to access actions quickly.

---

## ✅ Supported Editors

- ✅ ONLYOFFICE Document Editor  
- ✅ ONLYOFFICE Spreadsheet Editor
- ✅ Desktop Editors & Web Document Server  

---

## 💡 For Developers

- Built using the ONLYOFFICE plugin framework (`sdkjs-plugins`).
- Uses:
  - Document selection and `Api.ReplaceTextSmart()` for text edits in documents.
  - Cell addresses and `ApiRange.SetValue()` for in-place changes to spreadsheet cells.
- The panel and context menu are localized in English, Simplified Chinese, and Russian.
- Plugin configuration defined in `config.json`.
- Supports responsive side panel UI.

---

## 📦 Compatibility

- ✅ ONLYOFFICE Desktop Editors  
- ✅ ONLYOFFICE Document Server (Web)  
- ✅ Modern browser environments  

---

If you find this plugin helpful, please ⭐ Star, 🍴 Fork, and open Issues!