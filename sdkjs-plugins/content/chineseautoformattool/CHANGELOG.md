# Change Log

## 1.1.5

* Fixed Apply & Save in web editors so the report falls back to the original selection when no explicit base lines are provided.

## 1.1.4

* Removed legacy slide/PPT code paths so the plugin explicitly targets only Documents and Spreadsheets.

## 1.1.3

* Refined Force Full-width context detection so only punctuation in Chinese or mixed Chinese text is counted as converted, while plain English selections stay untouched.

## 1.1.2

* Limited Force Full-width conversion to punctuation that is adjacent to CJK text, so plain English lines are no longer reported as converted.

## 1.1.0

* Fixed web modal initialization and toolbar refresh handling.
* Improved punctuation conversion flow for DOCs and Spreadsheets.
* Updated report and spacing dialogs to follow editor theme changes.

## 1.0.0

* Initial release.
