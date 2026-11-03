# Serbian Text Converter (ONLYOFFICE plugin)

## Overview

This plugin converts selected Serbian text between Cyrillic and Latin script. It works in documents, spreadsheets, presentations, and PDF, from the side panel or the right-click menu.

## Features

- **Cyrillic → Latin** — correct digraph casing: `ЉУБАВ` → `LJUBAV`, `Љubav` → `Ljubav`, `КОЊ` → `KONJ`.
- **Latin → Cyrillic** — digraph exception list: `injekcija`, `konjunkcija`, `nadživeti`, `Tanjug`, `podžanr`.
- **Automatic swap** — detects the script of the selection and converts it to the other one.
- **Keep foreign words and brands** — words with `q`, `w`, `x`, `y` and known names (`Windows`, `Google`, `PDF`…) stay unchanged.

The panel follows the editor light/dark theme. The UI is available in Serbian and English.

## Structure

| File | Role |
| --- | --- |
| `config.json` | Plugin manifest |
| `index.html` | Side panel |
| `code.js` | Selection, text replace, context menu |
| `resources/cyrlatconverter.js` | Transliteration engine (browser and Node) |
| `translations/` | Serbian UI strings (panel default language is English) |

## Digraphs

Latin → Cyrillic is not a letter-by-letter replace, because `n+j`, `l+j`, and `d+ž` at a morpheme boundary are not `њ`, `љ`, and `џ`. `resources/cyrlatconverter.js` keeps a stem list (`DIGRAPH_EXCEPTIONS`) where those letters stay split. A stem matches only at the start of a word, so `kanjon` stays `кањон` even though it contains `anjon`. Only the first digraph in the stem is blocked, so `nadživljavanje` becomes `надживљавање` (split `дж`, kept `љ`).
