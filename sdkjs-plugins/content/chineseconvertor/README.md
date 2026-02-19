## Overview

This plugin can help you convert between simplified and traditional Chinese characters, as well as add or remove pinyin from text.

## Changelog

**Version 0.0.1 Release date: 2025/07/23**

- The plugin supports selecting text and right clicking to use the corresponding function through the context menu.

**Version 0.0.2 Release date: 2025/09/11**

- Add a side panel to the plugin.
- Fix plugin file directory format related issues.
- Localization of referenced JS packages.

**Version 0.0.3 Release date: 2025/09/30**

- Add icons to plugin.
- The plugin supports both Chinese and English languages.
- Replace appropriate screenshots for plugin introduction.
- Optimize the UI of plugins.
- Fix the issue of adding incorrect pinyin for polyphonic characters.

**Version 0.0.4 Release date: 2025/10/08**

- Supplement and improve the changelog of the plugin.
- Fix incorrect display in side panel.
- Add “offered by” information to plugin.
- remove the excess console log.

**Version 0.0.5 Release date: 2025/12/01**

- Completely refactored the core logic for processing selected text to ensure more accurate translation and pinyin generation.
- Added full support for selecting and converting text inside tables and lists within the Document Editor.
- Ensured compatibility with the Spreadsheet Editor, allowing selected cells to be translated or have pinyin added without issues.
- Removed support for the Slide Editor, as ReplaceTextSmart is not functional for text contained within slide tables.
- Improved stability of selection handling across editors and optimized the internal translation flow.