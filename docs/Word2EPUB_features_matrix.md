---
title: Features comparison between WordToEPUB and SaveAsDAISY
layout: my-default
---
# Features matrices

List and comparison of features between Word2EPUB and SaveAsDAISY.

- ✅ = supported
- ❌ = unsupported
- ⛔ = wont support
- 🟨 = partially supported
- ❓= to be checked

This is to track the current state of portage of the features from WordToEPUB into SaveAsDAISY.

## User settings

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| Save settings | ✅ | ✅ | |
| Load settings | ✅ | ❌ | |
| Reset settings | ✅ | ❌ | |
| Set defaults conversion options | ✅ | 🟨 | SAD: a limited set of options are available in settings for conversion options |
| Set defaults metadata | ✅ | ❌ | |
|  |  |  |  |
| **Set special textual markers to convert** |  |  | |
| + For `<aside></aside>` elements | ✅ | ❌ | |
| + For `<details></details>` elements | ✅ | ❌ | |
| + For `<hr />` element | ✅ | ❌ | |
| + For page numbers | ✅ | ❌ | |
| + For landmarks | ✅ | ❌ | |
|  |  |  |  |
| **Footnotes customization** |  |  | SAD: was initially requested by Gautier Chomel for AVH, not sure it is used elsewhere, it is used at AVH only to circumvent a bug in the production chain |
| + Set notes position in dtbook XML | ❌ | ✅ | |
| + Set custom numbering and prefixes | ❌ | ✅ | Not used anymore by AVH |

## Conversions to Word document

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| EPUB | ✅ | ❌ | W2E: does not import directly in word, but provides an EpubToWord tool to convert epubs into docx file using pandoc |
| DTBook XML | ❌ | ✅ | SAD: can be done through 2 scripts, dtbook-to-odt and dtbook-to-rtf, exposed in the ribbon |
| PDF | ❌ | ✅ | SAD: using the new pdf-to-word-mistral-ocr script, exposed in the ribbon, requires a mistral API key|

## Word document modifications

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| Insert and manage acronyms  | ❌ | ✅ | |
| Insert and manage abbreviations  | ❌ | ✅ | |
| Change document and selection language  | ❌ | ✅ | SAD: could normally be done using the "Review/Language" dialog, but I personally add inconsistent results with the word original dialog, while the one in SAD worked as expected|
| Load DAISY styles into to the document  | ❌ | ✅ | |
| **Set/Update Metadata in document** |  |  | (both tools use the same custom properties for interoperability) |
| + Title   | ✅ | ✅ | |
| + Subtitle    | ✅ | ✅ | |
| + Author  | ✅ | ✅ | |
| + Contributor     | ✅ | ✅ | |
| + Publisher   | ✅ | ✅ | |
| + Rights  | ✅ | ✅ | |
| + ISBN    | ✅ | ✅ | |
| + EPUB Date   | ✅ | ✅ | |
| + Source  | ✅ | ✅ | |
| + Source Date     | ✅ | ✅ | |
| + Book summary    | ✅ | ✅ | |
| + Subject     | ✅ | ✅ | |
| + Accessibility summary   | ✅ | ✅ | |

## Checks before conversion

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| Headings report | ✅ | ❌ | |


## Word document conversion

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| **Pagination compute** |  |  | |
| + Using word pagination | ✅ | 🟨 |(SAD: attempt in the script side currently not working, script uses `lastRenderedPageBreak` but those seems computed on last pages rendered on user's screen)|
| + Using headers and footers | ✅ | ❌ | |
| + Using Heading6 | ✅ | ❌ | |
| + Using DAISY style "Pagenum" | ✅ | ✅ | |
| + Using textual marker | ✅ | ❌ | |
|  |  |  |  |
| **Objects/Content conversion** |  |  | |
| + Shapes (vectorial images) | ✅ | ✅ | |
| + Extract images | ✅ | 🟨 | SAD only extract images (no conversion to other format is done) |
| + + Convert to a specified format | ✅ | ❌ | |
| + + Take image transformations in account | ✅ | ❌ | |
| + Office math equations | ✅ | ✅ | |
| + MathType equations | ❓ | ✅ | |
| + footnotes / endnotes | ❌ | ✅ | |
| + + With customization from settings | ❌ | ✅ | |
|  |  |  |  |
| **EPUB rendering customization (CSS)** |  |  |  |
| + Embedding a CSS provided by the user | ✅ | ❌ | |
| + Embedding a CSS computed from options | ✅ | ❌ | |
| + + Set font (+ embedding in epub) | ✅ | ❌ | |
| + + Set text alignment | ✅ | ❌ | |
| + + Set paragraphs indentation | ✅ | ❌ | |
| + + Set tables style | ✅ | ❌ | |
|  |  |  |  |
| **EPUBs Metadata for accessibility** | | | SAD: output to be validated in word-to-* scripts |
| + Title   | ✅ | ✅ | |
| + Subtitle    | ✅ | ❓ | |
| + Author  | ✅ | ✅ | |
| + Contributor     | ✅ | ❓ | |
| + Publisher   | ✅ | ✅ | |
| + Rights  | ✅ | ❓ | |
| + ISBN    | ✅ | ❓ | |
| + EPUB Date   | ✅ | ❓ | |
| + Source  | ✅ | ✅ | |
| + Source Date     | ✅ | ❓ | |
| + Book summary    | ✅ | ❓ | |
| + Subject     | ✅ | ❓ | |
| + Accessibility summary   | ✅ | ❓ | |
|  |  |  |  |
| **EPUB content options** |  |  |  |
| + Compute a cover image | ✅ | ❌ | |
| + Set a cover image | ✅ | ❌ | |
| + + In metadata | ✅ | ❌ | |
| + + In content | ✅ | ❌ | |
| + Detect and/or set document language  | ✅ | ✅ | SAD : only in word-to-epub script |
| + Change text direction | ✅ | ❌ | |
| + Modify table of content | ✅ | ❌ | |
| + + Change depth of the TOC | ✅ | ❌ | |
| + + Include TOC in content | ✅ | ❌ | |
| + + Delete word TOC | ✅ | ❌ | |
| + Add a metadata page at the end | ✅ | ❌ | |
| + Splitting epub based on a level | ✅ | ❌ |  |
| + Include visible page numbers | ✅ | ❌ | |
| + Include title page inline | ✅ | ⛔ | From Prashant : "Title page should not be kept in-line because its visual appearance is not great as it has only title and author on the whole page. |
|  |  |  |  |
| **Textual markers conversion** |  |  | W2E: markers are process afterward (additionnal markers are used in conversion, but they seem internal to the conversion chain for pre and post process)|
| + For `<aside></aside>` elements | ✅ | ❌ | |
| + For `<details></details>` elements | ✅ | ❌ | |
| + For `<hr />` element | ✅ | ❌ | |
| + For page numbers | ✅ | ❌ | |
| + For landmarks | ✅ | ❌ | |
|  |  |  |  |
| **Additional export formats** |  |  |  |
| + Export to html | ✅  | ❌ |noted as not necessary to include in the portage in meetings note|
| + Export to MOBI | ✅  | ❌ |  |
| + Export to DTBook XML | ❌  | ✅ |  |
| + Export to DAISY 3 | ❌  | ✅ |  |
| + Export to DAISY 2.02 | ❌  | ✅ |  | 
| + Export to MP3s | ❌  | ✅ |  |

## User interface

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| **Localization**  |  |  | W2E: extended work done on that part, SAD: only partial french translations available for the ribbon |
| + User interface | ✅ | 🟨 | |
| + documentation | ✅ | 🟨 | |
| **User documentation**  |  |  | SAD: documentation might need some update|
| + User manual | ✅ | 🟨 | |
| + Guidelines | ✅ | 🟨 | |
| Updater | ✅ | ❌ | |