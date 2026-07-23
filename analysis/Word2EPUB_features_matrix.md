# Features matrices

List and Comparison of features bewteen Word2EPUB and SaveAsDAISY for the project of portage.

## User settings

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| Save settings | [x] | [x] | |
| Load settings | [x] | [ ] | |
| Reset settings | [x] | [ ] | |
| Set defaults conversion options | [x] | [-] | |
| Set defaults metadata | [x] | [ ] | |
|  |  |  |  |
| **Set special textual markers to convert** | [x] | [ ] | |
| + For `<aside></aside>` elements | [x] | [ ] | |
| + For `<details></details>` elements | [x] | [ ] | |
| + For `<hr />` element | [x] | [ ] | |
| + For page numbers | [x] | [ ] | |
| + For landmarks | [x] | [ ] | |
|  |  |  |  |
| **Footnotes customization** | [ ] | [x] | SAD: was initially requested by Gautier Chomel for AVH, not sure it is used elsewhere, it is used at AVH only to circumvent a bug in the production chain |
| + Set notes position in dtbook XML | [ ] | [x] | |
| + Set custom numbering and prefixes | [ ] | [x] | Not used anymore by AVH |

## Conversions to Word document

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| EPUB | [x] | [ ] | W2E: does not import directly in word, but provides an EpubToWord tool to convert epubs into docx file using pandoc |
| DTBook XML | [ ] | [x] | SAD: can be done through 2 scripts, dtbook-to-odt and dtbook-to-rtf, exposed in the ribbon |
| PDF | [ ] | [x] | SAD: using the new pdf-to-word-mistral-ocr script, exposed in the ribbon, requires a mistral API key|

## Word document modifications

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| Insert and manage acronyms  | [ ] | [x] | |
| Insert and manage abbreviations  | [ ] | [x] | |
| Change document and selection language  | [ ] | [x] | SAD: could normally be done using the "Review/Language" dialog, but I personally add inconsistent results with the word original dialog, while the one in SAD worked as expected|
| Load DAISY styles into to the document  | [ ] | [x] | |
| **Set/Update Metadata in document** | [x] | [x] | (both tools use the same custom properties for interoperability) |
| + Title   | [x] | [x] | |
| + Subtitle    | [x] | [x] | |
| + Author  | [x] | [x] | |
| + Contributor     | [x] | [x] | |
| + Publisher   | [x] | [x] | |
| + Rights  | [x] | [x] | |
| + ISBN    | [x] | [x] | |
| + EPUB Date   | [x] | [x] | |
| + Source  | [x] | [x] | |
| + Source Date     | [x] | [x] | |
| + Book summary    | [x] | [x] | |
| + Subject     | [x] | [x] | |
| + Accessibility summary   | [x] | [x] | |

## Checks before conversion

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| Headings report | [x] | [ ] | |


## Word document conversion

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| **Pagination compute** | [x] | [ ] | |
| + Using word pagination | [x] | [-] |(SAD: attempt in the script side currently not working, script uses `lastRenderedPageBreak` but those seems computed on last pages rendered on user's screen)|
| + Using headers and footers | [x] | [ ] | |
| + Using Heading6 | [x] | [ ] | |
| + Using DAISY style "Pagenum" | [x] | [x] | |
| + Using textual marker | [x] | [ ] | |
|  |  |  |  |
| **Objects/Content conversion** | [x] | [ ] | |
| + Shapes (vectorial images) | [x] | [X] | |
| + Images | [x] | [-] | SAD only extract images (no conversion to other format is done) |
| + + including image transformations | [x] | [ ] | |
| + Office math equations | [x] | [x] | |
| + Mathtype equations | [ ] | [x] | |
| + footnotes / endnotes | [ ] | [x] | |
| + + With customization from settings | [ ] | [x] | |
|  |  |  |  |
| **EPUB rendering customization (CSS)** | [x] | [ ] | W2E |
| + Embedding a CSS provided by the user | [x] | [ ] | |
| + Embedding a CSS computed from options | [x] | [ ] | |
| + + Set font (+ embedding in epub) | [x] | [ ] | |
| + + Set text alignment | [x] | [ ] | |
| + + Set paragraphs indentation | [x] | [ ] | |
| + + Set tables style | [x] | [ ] | |
|  |  |  |  |
| **EPUBs Metadata for accessibility** | [x] | [-] | SAD: output to be validated in word-to-* scripts |
| + Title   | [x] | [x] | |
| + Subtitle    | [x] | [ ] | |
| + Author  | [x] | [x] | |
| + Contributor     | [x] | [ ] | |
| + Publisher   | [x] | [x] | |
| + Rights  | [x] | [ ] | |
| + ISBN    | [x] | [ ] | |
| + EPUB Date   | [x] | [ ] | |
| + Source  | [x] | [x] | |
| + Source Date     | [x] | [ ] | |
| + Book summary    | [x] | [ ] | |
| + Subject     | [x] | [ ] | |
| + Accessibility summary   | [x] | [ ] | |
|  |  |  |  |
| **EPUB content options** |  |  |  |
| + Compute a cover image | [x] | [ ] | |
| + Set a cover image | [x] | [ ] | |
| + + In metadata | [x] | [ ] | |
| + + In content | [x] | [ ] | |
| + Detect and/or set document language  | [x] | [x] | SAD : only in word-to-epub script |
| + Change text direction | [x] | [ ] | |
| + Modify table of content | [x] | [ ] | |
| + + Change depth of the TOC | [x] | [ ] | |
| + + Include TOC in content | [x] | [ ] | |
| + + Delete word TOC | [x] | [ ] | |
| + Add a metadata page at the end | [x] | [ ] | |
| + Splitting epub based on a level | [x] | [ ] |  |
| + Include visible page numbers | [x] | [ ] | |
| + Include title page inline | [x] | [ ] | From Prashant : "Title page should not be kept in-line because its visual appearance is not great as it has only title and author on the whole page. |
|  |  |  |  |
| **Textual markers conversion** | [x] | [ ] | W2E: markers are treated in post-process (additionnal markers are used in conversion|
| + For `<aside></aside>` elements | [x] | [ ] | |
| + For `<details></details>` elements | [x] | [ ] | |
| + For `<hr />` element | [x] | [ ] | |
| + For page numbers | [x] | [ ] | |
| + For landmarks | [x] | [ ] | |
|  |  |  |  |
| **Additional export formats** |  |  |  |
| + Export to html | [x]  | [ ] | SAD: could be added through dtbook-to-html script if requested |
| + Export to MOBI | [x]  | [ ] |  |
| + Export to DTBook XML | [ ]  | [x] |  |
| + Export to DAISY 3 | [ ]  | [x] |  |
| + Export to DAISY 2.02 | [ ]  | [x] |  | 
| + Export to MP3s | [ ]  | [x] |  |

## User interface

| Feature | WordToEPUB (W2E) | SaveAsDAISY (SAD) | Comment |
| ------- | ------- | ------- | ------- |
| UI Localization  | [x] | [-] | W2E: extended work done on that part, SAD: only partial french translations available for the ribbon |
| **User documentation**  | [x] | [-] | |
| + User manual | [x] | [x] | |
| + Guidelines | [x] | [x] | |
| Updater | [x] | [x] | |