# Change Log

## 4.0.4

* Publish the student's fourth Insert Note revision unchanged for web testing.
* Include the current-slide prompt example and cleanup on group-action startup failure.
* Known limitations: content-only table cells are omitted and slide-number validation is not integer-safe.

## 4.0.3

* Expand Insert Note metadata for talking points, speaker notes and the current slide.
* Insert literal notes in one editor call and keep the resolved slide index during AI generation.
* Read presentation table rows and report AI failures while releasing editor actions.

## 4.0.2

* Integrate the revised mock data generator with bounded row counts, explicit range errors and single-row headers.
* Protect existing target values and formulas, close actions on request failures, and treat header text as inert data.
* Handle single-cell headers and retain test diagnostics.

## 4.0.1

* Add Insert Note to presentation AI tools, with console diagnostics for both literal and generated notes.
* Report invalid inputs and unavailable note APIs; release editor actions on AI request failure.

## 4.0.0

* Include Generate mock data in the spreadsheet AI tools bundle.
* Validate the generated row count and release editor actions on request failure.

## 1.0.0

* Initial release.

## 1.0.1

* Add new model.

## 1.0.2

* Change plugin structure. Now it's non visual plugin and it can work by context menu.
* Add new model for chat variation.

## 1.1.0

* Commont improvements.

## 1.1.1

* Add new languages for translating and add translations for it.
* Change plugin type from "system" to "background".
* Increase minimal editor version to "7.5.0".
* Change type for chat window (now it's not a modal window).

## 1.1.2

* Disable plugin for IE.

## 1.1.3

* Add notification, that service isn't enable in user region (if we can't load list of models).

## 1.1.4
* Change list of models for custom requests (add new models and remove old).
* Change default model for request by context menu (now we use gpt-3.5-turbo-16k model for chat and gpt-4 for other request).
* Add new functions: "Fix spelling & grammar", "Rewrite differently", "Make longer", "Make shorter", "Make simpler" (it has restriction by gpt-4 model: 8k tokens)

## 2.0.0
* The plugin has been completely redesigned.

## 2.1.0
* Bug fix.

## 2.1.1
* Bug fix.

## 2.1.2
* Fix add provider.

## 2.1.3
* Bug fix. Remove v1 suffix for endpoints.

## 2.1.4
* Add proxy for together.ai.
* Add Groc as internal provider.
* Bug fix

## 2.1.5
* Bug fix.

## 2.2.0
* Refactoring

## 2.2.1
* Bug fix.

## 2.2.2
* Add xAI as internal provider.

## 2.2.3
* Fix translations.

## 2.2.4
* Refactoring chat. Add docked mode for chat window.

## 2.2.5
* Bug fix.

## 2.2.6
* Add interface for system role detection.
* Fixed the work of providers antropic and gemini when working with a system role.
* Fix bug with custom functions in spreadsheets editor.

## 2.2.7
* Add image actions.
* Add Stability AI provider.

## 2.2.8
* Fix image actions. Add "OCR" and "Image to text" support.

## 2.2.9
* Add support server settings.
* Bug fix.

## 2.3.0
* Add license.

## 2.3.1
* Bug fix.

## 2.3.2
* Fix offline work.

## 2.3.3
* Refactoring custom providers.

## 2.3.4
* Refactoring custom providers.

## 2.4.0
* Bug fix.

## 2.4.1
* Bug fix.

## 2.4.2
* Add agent for simple functions.

## 2.5.0
* Add events for another plugins.
* Bug fix.

## 2.5.1
* Bug fix.

## 3.0.0
* Bug fix.
* Add Grammar & Spelling functionality.
* Add OpenRouter provider.
* Refactoring helper functions.

## 3.0.1
* Bug fix.
* Translations added.

## 3.0.2
* Bug fix.

## 3.0.5
* New Feature: Build Your Own AI Assistants
* Known issues: immediately after adding the assistant in the desktop application (and only there), the Edit and Delete buttons will not work. After restarting the plugin or reopening the file, they will work correctly. This will be fixed in version 9.3.0.

## 3.0.6
* Fix bugs with external providers/models.

## 3.0.7
* Bug fix.

## 3.0.8
* Fix bug with external fetch error.

## 3.1.0
* Bug fix.

## 3.2.0
* Chat has been redesigned.

## 3.2.1
* Update storage version.

## 3.2.2
* Add support 9.4.0+ server settings.

## 3.2.3
* Bug fix.
