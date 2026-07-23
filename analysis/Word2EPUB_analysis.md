# Word2EPUB - UI and code analysis notes

Main "WordToEPUB" form contains the following elements (to be confirmed for my English translations as my UI is in French) :
- "Convert the document" label, followed by the document name as a label,
- "Save as" label, followed by a text box with the default value being the document name without extension, an ".epub" label
- "In folder" label, followed by a label with the default value being the document folder, and a "Browse" button to select another folder
- "Help" button, opens a "Version of WordToEPUB" dialog with the following elements:
  - read-only text box with a quick conversion workflow (some caracter issues in the localized text displayed to report to Richard : "Vous pouvez d�sormais")
  - "Documentation" button opens a "localized folder" with localized word document, i.e in french a folder "%PROGRAMFILES%DAISY\WordToEPUB\Help\fr" that contains the following documentations :
    - "Créer_des_documents_Word_accessibles.docx"
    - "Guide_d_utilisation_avancée_de_WordToEpub.docx"
    - "Guide_de_démarrage_rapide_de_WordToEpub.docx"
  - "Version notes" button opens a "Release_notes.txt" file
  - "Licences" button opens a "Licences.txt" file 
  - "Send comments" button opens a "mailto" link to send an email to "wordtoepub@daisy.org" with subject "WordToEPUB"
  - "Open errors report (log)" opens a "wordtoepublog.txt" file
  - "About DAISY" opens a "About DAISY Consortium" dialog with a localized text presenting the DAISY mission statement
  - A label in bold text stating the number of pages converted to epub with the tool
  - A label in bold text stating the number of files converted to epub with the tool
  - A "OK" button closes the dialog

- "Settings" buttons opens a "Settings" dialog containing a group of tabs with the following elements:
  - "Introduction" tab
    - Localized read-only text area, presenting the settings tabs usage
    - "Save settings" button
    - "Load settings" button
    - "Default settings" button
  - "Default conversion options" tab :
    - "Replace spaces par underscores in file names" checkbox (default to true)
    - "Include the title page inline" checkbox (default to false)
    - "Remove the Word's table of content" checkbox (default to true)
    - "Include the TOC in the content" checkbox (default to false)
    - "Document language" group with the following radio buttons:
      - "Detected language" (default)
      - "Selected language", activates a dropdown list with a list of possible language (Name and code) + an empty entry by default + a "Detect language automatically" entry
    - "Text direction" radio group with the following options:
      - "Left to right" (default)
      - "Right to left"
    - "Split the file based on level" label, followed by a dropdown list with possible values being an integer in range 1 to 3, default value is 2
    - "Depth of TOC" label, followed by a dropdown list with possible values being an integer in range 1 to 6, default value is 3
    - "Images" group:
      - "WebP" radio button
      - "JPG" radio button (default)
      - "Convert all images" checkbox (default to true)
      - "Use original file image dimensions" checkbox (default to false)
    - "XML" group:
      - "Format the XHTML (ease the code review)" radio button (default)
      - "Minify the XHTML (reduce the file size)" radio button
      - "Clean the XML OOML source (recommended)" checkbox (default to true)
  - "Special markers" tab with a group "Special text markers" :
    - `<details>` label, followed by a text box with the default value being "@begin details@"
    - `</details>` label, followed by a text box with the default value being "@end details@"
    - `<aside>` label, followed by a text box with the default value being "@begin aside@"
    - `</aside>` label, followed by a text box with the default value being "@end aside@"
    - `<hr />` label, followed by a text box with the default value being "@hr@"
    - "Landmark" label, followed by a text box with the default value being "@landmark@"
  - "Metadata" tab
    - "Default metadata" group of editable fields :
      - "Publisher" label followed by a text box
      - "Rights" label followed by a text box
      - "Source" label followed by a text box
      - "Accessibility summary" label followed by a text box
    - "Add a metadata page at the end of the EPUB" checkbox (default to false)
  - "Default pagination" tab
    - "To generate pagination, use" option with the following element :
      - "Word document pagination" radio button (default)
      - "Page number form headers and footers" radio button
      - "From style Heading 6" radio button
      - "From style Page Numbers (available in the DAISY template)" radio button
      - "Pagination included in the text" radio button
      - "Page numbers are marked by" label, followed by a text box with the default value being "PRINT PAGE " (with a space at the end)
      - "No page numbers" radio button
    - "Include visible page numbers" checkbox (default to false)
    - ""Text before the page numbers" label, followed by a text box
    - "Text after the page numbers" label, followed by a text box
  - "Default coverpage options" tab
    - "Coverpage image source" option with the following choices :
      - "Use document first page" radio button
      - "Select an image in the document" radio button
      - "Select an image file" radio button - on selection :
        - shows an image preview, a Browse button to select an image file, and a Reset button to reset the selection
        - Radio button text is replaced by the name of the selected image file
          - defaults to "DefaultCoverImage.png"
      - "Title and author" radio button (default)
  - "Appearance" tab:
    - "Default appearance" group:
      - "Select a stylesheet" radio button,
      - "Appearance based on selected options" radio button
    - "CSS file" label, followed by a text box with a "Browse" button and a "Reset" button
    - "Appearance options" group:
      - "Text alignment" option with the following choices :
        - "Not specified" radio button (default)
        - "Justify" radio button
        - "Left aligned" radio button
        - "Right aligned" radio button
      - "Paragraphe style" option with the following choices :
        - "Indent first line" radio button (default)
        - "No indentation" radio button
      - "Table style" option with the following choices :
        - "Grid" radio button (default)
        - "Horizontal lines" radio button
      - "Font" option with a dropdown list that contains the following list of values : 
        - "None" (default)
        - "APHfont"
        - "AtkinsonHyperlisible"
        - "Delicious"
        - "Luciole"
        - "OpenDyslexic"
     - "Include colorized text" checkbox (default to true)
     - "Map Word styles to CSS classes" group :
       - "Paragraph" label, followed by a text box
       - "Text" label, followed by a text box
  - "User interface" tab:
    - "User interface language" option with a dropdown list with a list of supported languages
    - "Display a help message if no file is selected at startup" checkbox (default to true)
    - "Display advanced options at startup" checkbox (default to false)
    - "Display informations during conversion (partially in english)" checkbox (default to true)
    - "Save conversion messages in a log file" checkbox (default to false)
    - "Highlight warning messages" checkbox (default to true)
    - "Propose conversion in HTML format" checkbox (default to true)
    - "Propose conversion to MOBI format after creating the EPUB" checkbox (default to false)
    - "Suggest SIGIL after the conversion" checkbox (default to false)
    - "Quality assurance assistant after conversion" checkbox (default to false)
  - "Updates" tab:
    - "Check for updates automatically" checkbox (default to true)
    - "Check for updates now" button
 - "Add word extension" tab
   - If word is installed :
     - "Add a WordToEpub button to word" label
     - "Install" button
     - "Uninstall" button
  - else (I suppose) :
    - "Word software not found"
    - "Install" button disabled
- "Add libreoffice extension" tab
  - If libreoffice is installed (i assume as i don't have it installed) :
    - "Add a WordToEpub button to libreoffice" label
    - "Install" button
    - "Uninstall" button
  - else :
    - "Libreoffice software not found"
    - "Install" button disabled

- "Advanced mode"/"Simple mode" button, exposes/hides a group of tabs with the following tabs:
  - "Metadata" :
      - Group of editable fields :
        - "Title" text box
        - "Subtitle" text box
        - "Author" text box
        - "Contributor" text box
        - "Publisher" text box
        - "Rights" text box
        - "ISBN" text box
        - "EPUB Date" text box
        - "Source" text box, with a localized default text stating the source of pagination (i.e. "La pagination du document Word est utilisé pour numéroter les pages")
        - "Source Date" text box,
        - "Book summary" text box
        - "Subject" text box
        - "Accessibility summary" text box, with localized default text stating the accessibility of the document (i.e. "Le document permet une navigation par structure et par pages. Les descriptions d'images sont incluses si elles étaient présentes dans le document Word.")
      - "Add a metadata page at the end of the EPUB" checkbox
      - "Add the metadata in the DOCX file" button
  - "Cover" tab:
    - a "Cover" option with the following choices :
      - "Don't use a cover image", 
      - "Include cover in metadata" 
      - "Include cover in content"
    - A "Source of cover image" option 
      - "Use document first page", 
      - "Select an image in the document" : that selects the first image available, displays it preview, and it activates 2 buttons to select the next and previous document image to be used and previewed, 
      - "DefaultCoverImage.png" : enables a "Select a cover image" button
      - "Title and author" : enables a "Generate a cover" button that opens a cover page generator form
  
  - "EPUB Options" tab:
    - "Language"
      - "Detected language" (displays the main language code detected folowwed by the list of other languages avaiables)
      - "Selected language" with a dropdown list with a list of possible language (Name and code)
    - "Text direction"
      - "Left to right"
      - "Right to left"
    - "Inlcude title page in line" (not sure of what this option do)
    - "Table of content"
      - "Depth" dropdown list, starts with value 3, values are in range 1 to 6.
      - "Insert TOC in content" checkbox, that activates a "TOC title" text box
        - "TOC title" text box
      - "Delete TOC of the word file"
    - "Split epub based on level"
      - possible values are an integer in range 1 to 3, default value is 2
  
  - "Pagination" tab:
    - "To générate pagination, use" option with the following choices :
      - "Word document pagination"
      - "Page number form headers and footers"
      - "From style Heading 6"
      - "From style Page Numbers (DAISY)"
      - "Included as PRINT PAGE #"
      - "No page numbers"'
    - "Include visible page numbers" checkbox that activate the 2 following text field
      - Text before the page numbers
      - Text after the page numbers
  
  - "Appearance" tab : 
    - "Type of appeareance" option has 2 choices:
      - "Select a stylesheet" activate an "Stylesheet" section of the form,
      - "Select option" value activate a "Style options" section of the form
    - "Stylesheet" section allows to select a css file with a "Browse" button and a "Reset" button
    -  "Style options" contains a set of selectable options:
      - "Text alignment" that can be one of "Justify", "Left aligned", "Right aligned" and "Not specified"
      - "Paragraphe style" can be either "Indent first line" of "No indentation"
      - "Table style" can be either "Grid lines" or "Horizontal lines"
      - "Font" is a dropdown list that contains the following list of values : 
        - "None"
        - "APHfont"
        - "AtkinsonHyperlisible"
        - "Delicious"
        - "Luciole"
        - "OpenDyslexic"
  
  - Headings tab:
    - provides a "heading report" that retrieve all levels and text of the headings in the content. with a button to open the report in a browser.

"Mathematic" tab would normally allows to select between mathml format and images with latex format as alt text, but the form states that only mathml is currently handled, with more options planned for futur versions.

As stated in a meeting, the "HTML" tab was included for debugging at first and is not required to be ported. 

## analysis notes

WordToEPUB takes in account more visual elements that are available in the epub but that are not present in the dtbook format, i.e : 
- "hr", "details" and "aside"  elements can be added using a textual marker "@hr@", "@begin details@", "@end details@", "@begin aside@" and "@end aside@" in the word document, 
that will be converted to `<hr />`, `<details></details>` and `<aside></aside>` elements in the epub, but "hr", "details" and "aside" elements are not available in the dtbook format
  - A possibility would be to replace those by "divs" with classes "hr", "details" and "aside" in the dtbook format.
  - This could be done with the following chain : 
    - Markers would be matched by paragraph styles, same as the "page number (DAISY)" style, and would be converted to a "div" with the appropriate class in the dtbook xml, and then converted to the appropriate element in the epub. 
      - "Horizontal ruler (DAISY)", that could be converted to `<div class="hr"></div>` in the dtbook xml
      - "Details begin (DAISY)" that could be converted to `<div class="details">` in the dtbook xml
      - "Details summary (DAISY)" that could be converted to `<p class="details-summary"> ... </p>` in the dtbook xml
      - "Details end (DAISY)" that could be converted to `</div>` in the dtbook xml
      - "Aside Begin (DAISY)" that could be converted to `<div class="aside">` in the dtbook xml
      - "Aside End (DAISY)" that could be converted to `</div>` in the dtbook xml

## Conversion process analysis 

As of 2026/07/21, from branch master of WordToEPUB repo.

- conversion is done in WordToEPUB.vb/FrmWordToEPUB/ConvertFromWord(BooIsBatchConversion As Boolean)




# Developements proposal

Work in progress

## In SaveAsDAISY

List of possible features to be ported to the addin side

- Update of the DAISY Model with new styles for textual markers support 
  - "Horizontal ruler (DAISY)" paragraph style
  - "Aside Start (DAISY)" paragraph style
  - "Aside End (DAISY)" paragraph style
  - "Details Start (DAISY)" paragraph style
  - "Details summary (DAISY)" paragraph style
  - "Details End (DAISY)" paragraph style

- Settings form updates :
  - Switch to a tabbed interface with the following additionnal or updated tabs:
    - "Conversion options" tab
      - Take back the general settings from the current form
      - Add the other 
    - "Footnotes customization" tab
      - Take back the settings fields defined for footnotes customization
    - "Special Markers" tab
      - Define the same fields as the one in WordToEPUB settings
    - "Default Metadata" tab
      - Define the same fields as the one in WordToEPUB settings
    - "Default coverpage options" tab
    - "Default appearance" tab

Notes : A tabbed interface might improve accessibility as the current "single section" form, that starts to be a bit bloated for keyboard navigation with newer settings being added.

In the appearance settings
- If no css is provided, use a default css (probably the same as WordToEPUB) and allow the user to:
  - Select a replacement font in the following list
    - "None" (default)
    - "APHfont"
    - "AtkinsonHyperlisible"
    - "Delicious"
    - "Luciole"
    - "OpenDyslexic"
  - Select the "Text alignment" to be one of the following list:
    - "Not specified" (default)
    - "Justify", 
    - "Left aligned", 
    - "Right aligned" 
  - Select the "Paragraphe style" to be one of the following list:
    - "No indentation" (default)
    - "Indent first line"
  - Select the "Table style" to be either 
    - "Grid lines" (default) 
    - "Horizontal lines" 


- Conversion form updates for the epub3 export
  - in the options expandable section of the form, 
    - Wrap the script options in an "Export options" tab
    - Add the following tabs in the section
      - "Cover" tab:
        - a "Cover" option with the following choices :
          - "Don't use a cover image", 
          - "Include cover in metadata" 
          - "Include cover in content"
        - A "Source of cover image" option 
          - "Use document first page", 
          - "Select an image in the document" : that selects the first image available, displays it preview, and it activates 2 buttons to select the next and previous document image to be used and previewed, 
          - "Select a cover image" option with a browse button and label specifying the file to include
          - ~~"Title and author" : would enables a "Generate a cover" button that opens a cover page generator form (this might be proposed as a shared utility between the addins)~~ 
      
      - "EPUB layout Options" tab:
        - "Text direction"
          - "Left to right"
          - "Right to left"
        - ~~"Inlcude title page in line" (Prashant mentionned "Title page should not be kept in-line because its visual appearance is not great as it has only title and author on the whole page.")~~
        - "Table of content"
          - "Depth" dropdown list, starts with value 3, values are in range 1 to 6.
          - "Insert TOC in content" checkbox, that activates a "TOC title" text box
            - "TOC title" text box
          - "Delete TOC of the word file"
        - "Split epub based on level"
          - possible values are an integer in range 1 to 3, default value is 2
      
      - "Pagination" tab:
        - "To générate pagination, use" option with the following choices :
          - "Word document pagination"
          - "Page number form headers and footers"
          - "From style Heading 6"
          - "From style Page Numbers (DAISY)"
          - "Included as PRINT PAGE #"
          - "No page numbers"'
        - "Include visible page numbers" checkbox that activate the 2 following text field
          - Text before the page numbers
          - Text after the page numbers
      
      - "Appearance" tab : 
        - "Type of appeareance" option has 2 choices:
          - "Select a stylesheet" activate an "Stylesheet" section of the form,
          - "Select option" value activate a "Style options" section of the form
        - "Stylesheet" section allows to select a css file with a "Browse" button and a "Reset" button
        -  "Style options" contains a set of selectable options:
          - "Text alignment" that can be one of "Justify", "Left aligned", "Right aligned" and "Not specified"
          - "Paragraphe style" can be either "Indent first line" of "No indentation"
          - "Table style" can be either "Grid lines" or "Horizontal lines"
          - "Font" is a dropdown list that contains the following list of values : 
            - "None"
            - "APHfont"
            - "AtkinsonHyperlisible"
            - "Delicious"
            - "Luciole"
            - "OpenDyslexic"
      
      - Headings tab:
        - provides a "heading report" that retrieve all levels and text of the headings in the content. with a button to open the report in a browser.

- Document preprocessing, between the form validation and the conversion script exeuction :
  - Recompute pages based on pagination settings, 
    - If "Word document pagination" is selected, compute and inject pagenum text at the end of each page in the document, using the "page number (DAISY)" style
    - If "Page number form headers and footers" is selected, use the page numbers from headers and footers to inject pagenum text at the end of each page in the document, using the "page number (DAISY)" style
    - If "From style Heading 6" is selected, replace Heading 6 styled elements by "page number (DAISY)" styled text, using the Heading 6 text as page numbers
    - If "From style Page Numbers (available in the DAISY template)" is selected, this is the default case of SaveAsDAISY, don't do anything
    - If "Pagination included in the text" is selected, use the text of textual markers as page numbers to inject pagenum text at the end of each page in the document, using the "page number (DAISY)" style
    - If "No page numbers" is selected, don't parse page numbers and possibly remove any styled text that is using the "page number (DAISY)" style
  - Replace textual markers found in the document by an empty paragraph with the correspondig style
  - Recompute/package the css and it's required resources to be passed as settings to the epub3 export script
  - Based on cover image options, 
    - If no cover requested, do nothing
    - If cover is first page of the document, create an temoporary image by rendering the first page
      - In that case, the cover should only be registered in the output metadata


- Idea of additional ribbon buttons (not validated, it's more of a set of reflections)
  - Prepaginate button, that opens a "Pagination" dialog with the following elements:
    - A label stating "This action will add page numbers as "page number (DAISY)" styled text. Note that H6 or textual markers will be replaced in content."
    - "Source of pagination" option with the following choices :
      - "Word document pagination"
      - "Page number form headers and footers"
      - "From style Heading 6"
      - "From textual markers 'PRINT PAGE #'" (with 'print page ' text being editable)

## In DAISY Pipeline 2

Under investigation : how to include the cover image / cover page in the document and the conversion to dtbook and/or to epub
- If it is to be kept only for epub exports, the following additionnal settings could be considered :
  - add a "cover-image-file" settings that take a file uri to the image
  - add a "cover-image-inclusion" settings with possible values being "in metadata" or "in metadata and content"

- If it is to be kept also in the dtbook, we might consider the following setting in the word to dtbook script :
  - Add a "cover-image-file" setting that take a file uri to an image
  - add a "cover-image-inclusion" settings with possible values being "in metadata" or "in metadata and content"
    - If the cover should be only be declared in metadata, add the cover-image uri in relative link or a head/meta element ?
    - If it is to be declared in content, also add the image under a level1 with id "cover" in the front matter
  - then in the dtbook to epub chain, convert accordingly

Changes possibly required in the word to epub3 chain :
- In word to dtbook conversion, 
  - the following styles would need to be handled in conversion
    - "Horizontal ruler (DAISY)" style, that would be converted to `<p class="hr"></p>` (to be defined, maybe div is more appropriate, or something else ?) in the dtbook xml
    - "Landmark (DAISY)" style, that would be converted to `<a class="landmark"/>` in the dtbook xml
    - "Aside Start (DAISY)" style, that would be converted to `<div class="aside">` in the dtbook xml
    - "Aside End (DAISY)" style, that would be converted to `</div>` in the dtbook xml
    - "Details Start (DAISY)" style, that would be converted to `<div class="details">` in the dtbook xml
    - "Details summary (DAISY)" paragraph style that would be converted to `<bridgehead class="details-summary">..</bridgehead>` in the dtbook xml
    - "Details End (DAISY)" style, that would be converted to `</div>` in the dtbook xml

- In dtbook to epub conversion: 
  - the following element and class might be required to be handled in the conversion
    - `<div class="hr"></div>` in the dtbook xml would be converted to `<hr />` in the epub
    - `<div class="aside">..</div>` in the dtbook xml would be converted to `<aside>..</aside>` in the epub
    - `<div class="details">..</div>` in the dtbook xml would be converted to `<details>..</details>` in the epub
  - Allow to set a custom css to be used in the epub for its rendering
  - Allow to embed css resources like fonts


## Features matrix

### UI

| Feature | WordToEPUB + Pandoc | SaveAsDAISY + DAISY Pipeline 2 | Comment |
| ------- | ------- | ------- | ------- |
| Settings |  |  |  |
| "Save settings" | | | |
| "Load settings" | | | |
| "Reset settings" | | | |
|  - "Default conversion options" : | | | |
|    - "Replace spaces by underscores in file names" checkbox (default to true) | | | |
|    - "Include the title page inline" checkbox (default to false) | | | |
|    - "Remove the Word's table of content" checkbox (default to true) | | | |
|    - "Include the TOC in the content" checkbox (default to false) | | | |
|    - "Document language" group with the following radio buttons: | | | |
|      - "Detected language" (default) | | | |
|      - "Selected language", activates a dropdown list with a list of possible language (Name and code) + an  empty entry by default + a "Detect language automatically" entry | | | |
|    - "Text direction" options: | | | |
|      - "Left to right" (default) | | | |
|      - "Right to left" | | | |
|    - "Split the file based on level" label, followed by a dropdown list with possible values being an integer  in range 1 to 3, default value is 2 | | | |
|    - "Depth of TOC" label, followed by a dropdown list with possible values being an integer in range 1 to 6,  default value is 3 | | | |
|    - "Images" group: | | | |
|      - "WebP" radio | | | |
|      - "JPG" radio (default) | | | |
|      - "Convert all images" checkbox (default to true) | | | |
|      - "Use original file image dimensions" checkbox (default to false) | | | |
|    - "XML" group: | | | |
|      - "Format the XHTML (ease the code review)" radio (default) | | | |
|      - "Minify the XHTML (reduce the file size)" radio | | | |
|      - "Clean the XML OOML source (recommended)" checkbox (default to true) | | | |
|  - "Special markers" with a group "Special text markers" : | | | |
|    - `<details>` label, followed by a text box with the default value being "@begin details@" | | | |
|    - `</details>` label, followed by a text box with the default value being "@end details@" | | | |
|    - `<aside>` label, followed by a text box with the default value being "@begin aside@" | | | |
|    - `</aside>` label, followed by a text box with the default value being "@end aside@" | | | |
|    - `<hr />` label, followed by a text box with the default value being "@hr@" | | | |
|    - "Landmark" label, followed by a text box with the default value being "@landmark@" | | | |
|  - "Metadata" | | | |
|    - "Default metadata" group of editable fields : | | | |
|      - "Publisher" label followed by a text box | | | |
|      - "Rights" label followed by a text box | | | |
|      - "Source" label followed by a text box | | | |
|      - "Accessibility summary" label followed by a text box | | | |
|    - "Add a metadata page at the end of the EPUB" checkbox (default to false) | | | |
|  - "Default pagination" | | | |
|    - "To generate pagination, use" option with the following element : | | | |
|      - "Word document pagination" radio (default) | | | |
|      - "Page number form headers and footers" radio | | | |
|      - "From style Heading 6" radio | | | |
|      - "From style Page Numbers (available in the DAISY template)" radio | | | |
|      - "Pagination included in the text" radio | | | |
|      - "Page numbers are marked by" label, followed by a text box with the default value being "PRINT PAGE "  (with a space at the end)| | | |
|      - "No page numbers" radio | | | |
|    - "Include visible page numbers" checkbox (default to false) | | | |
|    - ""Text before the page numbers" label, followed by a text box | | | |
|    - "Text after the page numbers" label, followed by a text box | | | |
|  - "Default coverpage options" | | | |
|    - "Coverpage image source" option with the following choices : | | | |
|      - "Use document first page" radio | | | |
|      - "Select an image in the document" radio | | | |
|      - "Select an image file" radio - on selection : | | | |
|        - shows an image preview, a Browse to select an image file, and a Reset to reset the selection | | | |
|        - Radio text is replaced by the name of the selected image file | | | |
|          - defaults to "DefaultCoverImage.png" | | | |
|      - "Title and author" radio (default) | | | |
|  - "Appearance" tab: | | | |
|    - "Default appearance" group: | | | |
|      - "Select a stylesheet" radio button, | | | |
|      - "Appearance based on selected options" radio | | | |
|    - "CSS file" label, followed by a text box with a "Browse" and a "Reset" | | | |
|    - "Appearance options" group: | | | |
|      - "Text alignment" option with the following choices : | | | |
|        - "Not specified" radio (default) | | | |
|        - "Justify" radio | | | |
|        - "Left aligned" radio | | | |
|        - "Right aligned" radio | | | |
|      - "Paragraphe style" option with the following choices : | | | |
|        - "Indent first line" radio (default) | | | |
|        - "No indentation" radio | | | |
|      - "Table style" option with the following choices : | | | |
|        - "Grid" radio (default) | | | |
|        - "Horizontal lines" radio | | | |
|      - "Font" option with a dropdown list that contains the following list of values :  | | | |
|        - "None" (default) | | | |
|        - "APHfont" | | | |
|        - "AtkinsonHyperlisible" | | | |
|        - "Delicious" | | | |
|        - "Luciole" | | | |
|        - "OpenDyslexic" | | | |
|     - "Include colorized text" checkbox (default to true) | | | |
|     - "Map Word styles to CSS classes" group : | | | |
|       - "Paragraph" label, followed by a text box | | | |
|       - "Text" label, followed by a text box | | | |
|  - "User interface" tab: | | | |
|    - "User interface language" option with a dropdown list with a list of supported languages | | | |
|    - "Display a help message if no file is selected at startup" checkbox (default to true) | | | |
|    - "Display advanced options at startup" checkbox (default to false) | | | |
|    - "Display informations during conversion (partially in english)" checkbox (default to true) | | | |
|    - "Save conversion messages in a log file" checkbox (default to false) | | | |
|    - "Highlight warning messages" checkbox (default to true) | | | |
|    - "Propose conversion in HTML format" checkbox (default to true) | | | |
|    - "Propose conversion to MOBI format after creating the EPUB" checkbox (default to false) | | | |
|    - "Suggest SIGIL after the conversion" checkbox (default to false) | | | |
|    - "Quality assurance assistant after conversion" checkbox (default to false) | | | |
|  - "Updates" tab: | | | |
|    - "Check for updates automatically" checkbox (default to true) | | | |
|    - "Check for updates now" | | | |
| - "Add word extension" | | | |
|   - If word is installed : | | | |
|     - "Add a WordToEpub to word" label | | | |
|     - "Install" | | | |
|     - "Uninstall" | | | |
|  - else (I suppose) : | | | |
|    - "Word software not found" | | | |
|    - "Install" disabled | | | |
|- "Add libreoffice extension" | | | |
|  - If libreoffice is installed (i assume as i don't have it installed) : | | | |
|    - "Add a WordToEpub to libreoffice" label | | | |
|    - "Install" | | | |
|    - "Uninstall" | | | |
|  - else : | | | |
|    - "Libreoffice software not found" | | | |
|    - "Install" disabled | | | |






- "Help" button, opens a "Version of WordToEPUB" dialog with the following elements:
  - read-only text box with a quick conversion workflow (some caracter issues in the localized text displayed to report to Richard : "Vous pouvez d�sormais")
  - "Documentation" button opens a "localized folder" with localized word document, i.e in french a folder "%PROGRAMFILES%DAISY\WordToEPUB\Help\fr" that contains the following documentations :
    - "Créer_des_documents_Word_accessibles.docx"
    - "Guide_d_utilisation_avancée_de_WordToEpub.docx"
    - "Guide_de_démarrage_rapide_de_WordToEpub.docx"
  - "Version notes" button opens a "Release_notes.txt" file
  - "Licences" button opens a "Licences.txt" file 
  - "Send comments" button opens a "mailto" link to send an email to "wordtoepub@daisy.org" with subject "WordToEPUB"
  - "Open errors report (log)" opens a "wordtoepublog.txt" file
  - "About DAISY" opens a "About DAISY Consortium" dialog with a localized text presenting the DAISY mission statement
  - A label in bold text stating the number of pages converted to epub with the tool
  - A label in bold text stating the number of files converted to epub with the tool
  - A "OK" button closes the dialog

- "Settings" buttons opens a "Settings" dialog containing a group of tabs with the following elements:
  - "Introduction" tab
    - Localized read-only text area, presenting the settings tabs usage
    - "Save settings" button
    - "Load settings" button
    - "Default settings" button
  - "Default conversion options" tab :
    - "Replace spaces par underscores in file names" checkbox (default to true)
    - "Include the title page inline" checkbox (default to false)
    - "Remove the Word's table of content" checkbox (default to true)
    - "Include the TOC in the content" checkbox (default to false)
    - "Document language" group with the following radio buttons:
      - "Detected language" (default)
      - "Selected language", activates a dropdown list with a list of possible language (Name and code) + an empty entry by default + a "Detect language automatically" entry
    - "Text direction" radio group with the following options:
      - "Left to right" (default)
      - "Right to left"
    - "Split the file based on level" label, followed by a dropdown list with possible values being an integer in range 1 to 3, default value is 2
    - "Depth of TOC" label, followed by a dropdown list with possible values being an integer in range 1 to 6, default value is 3
    - "Images" group:
      - "WebP" radio button
      - "JPG" radio button (default)
      - "Convert all images" checkbox (default to true)
      - "Use original file image dimensions" checkbox (default to false)
    - "XML" group:
      - "Format the XHTML (ease the code review)" radio button (default)
      - "Minify the XHTML (reduce the file size)" radio button
      - "Clean the XML OOML source (recommended)" checkbox (default to true)
  - "Special markers" tab with a group "Special text markers" :
    - `<details>` label, followed by a text box with the default value being "@begin details@"
    - `</details>` label, followed by a text box with the default value being "@end details@"
    - `<aside>` label, followed by a text box with the default value being "@begin aside@"
    - `</aside>` label, followed by a text box with the default value being "@end aside@"
    - `<hr />` label, followed by a text box with the default value being "@hr@"
    - "Landmark" label, followed by a text box with the default value being "@landmark@"
  - "Metadata" tab
    - "Default metadata" group of editable fields :
      - "Publisher" label followed by a text box
      - "Rights" label followed by a text box
      - "Source" label followed by a text box
      - "Accessibility summary" label followed by a text box
    - "Add a metadata page at the end of the EPUB" checkbox (default to false)
  - "Default pagination" tab
    - "To generate pagination, use" option with the following element :
      - "Word document pagination" radio button (default)
      - "Page number form headers and footers" radio button
      - "From style Heading 6" radio button
      - "From style Page Numbers (available in the DAISY template)" radio button
      - "Pagination included in the text" radio button
      - "Page numbers are marked by" label, followed by a text box with the default value being "PRINT PAGE " (with a space at the end)
      - "No page numbers" radio button
    - "Include visible page numbers" checkbox (default to false)
    - ""Text before the page numbers" label, followed by a text box
    - "Text after the page numbers" label, followed by a text box
  - "Default coverpage options" tab
    - "Coverpage image source" option with the following choices :
      - "Use document first page" radio button
      - "Select an image in the document" radio button
      - "Select an image file" radio button - on selection :
        - shows an image preview, a Browse button to select an image file, and a Reset button to reset the selection
        - Radio button text is replaced by the name of the selected image file
          - defaults to "DefaultCoverImage.png"
      - "Title and author" radio button (default)
  - "Appearance" tab:
    - "Default appearance" group:
      - "Select a stylesheet" radio button,
      - "Appearance based on selected options" radio button
    - "CSS file" label, followed by a text box with a "Browse" button and a "Reset" button
    - "Appearance options" group:
      - "Text alignment" option with the following choices :
        - "Not specified" radio button (default)
        - "Justify" radio button
        - "Left aligned" radio button
        - "Right aligned" radio button
      - "Paragraphe style" option with the following choices :
        - "Indent first line" radio button (default)
        - "No indentation" radio button
      - "Table style" option with the following choices :
        - "Grid" radio button (default)
        - "Horizontal lines" radio button
      - "Font" option with a dropdown list that contains the following list of values : 
        - "None" (default)
        - "APHfont"
        - "AtkinsonHyperlisible"
        - "Delicious"
        - "Luciole"
        - "OpenDyslexic"
     - "Include colorized text" checkbox (default to true)
     - "Map Word styles to CSS classes" group :
       - "Paragraph" label, followed by a text box
       - "Text" label, followed by a text box
  - "User interface" tab:
    - "User interface language" option with a dropdown list with a list of supported languages
    - "Display a help message if no file is selected at startup" checkbox (default to true)
    - "Display advanced options at startup" checkbox (default to false)
    - "Display informations during conversion (partially in english)" checkbox (default to true)
    - "Save conversion messages in a log file" checkbox (default to false)
    - "Highlight warning messages" checkbox (default to true)
    - "Propose conversion in HTML format" checkbox (default to true)
    - "Propose conversion to MOBI format after creating the EPUB" checkbox (default to false)
    - "Suggest SIGIL after the conversion" checkbox (default to false)
    - "Quality assurance assistant after conversion" checkbox (default to false)
  - "Updates" tab:
    - "Check for updates automatically" checkbox (default to true)
    - "Check for updates now" button
 - "Add word extension" tab
   - If word is installed :
     - "Add a WordToEpub button to word" label
     - "Install" button
     - "Uninstall" button
  - else (I suppose) :
    - "Word software not found"
    - "Install" button disabled
- "Add libreoffice extension" tab
  - If libreoffice is installed (i assume as i don't have it installed) :
    - "Add a WordToEpub button to libreoffice" label
    - "Install" button
    - "Uninstall" button
  - else :
    - "Libreoffice software not found"
    - "Install" button disabled

- "Advanced mode"/"Simple mode" button, exposes/hides a group of tabs with the following tabs:
  - "Metadata" :
      - Group of editable fields :
        - "Title" text box
        - "Subtitle" text box
        - "Author" text box
        - "Contributor" text box
        - "Publisher" text box
        - "Rights" text box
        - "ISBN" text box
        - "EPUB Date" text box
        - "Source" text box, with a localized default text stating the source of pagination (i.e. "La pagination du document Word est utilisé pour numéroter les pages")
        - "Source Date" text box,
        - "Book summary" text box
        - "Subject" text box
        - "Accessibility summary" text box, with localized default text stating the accessibility of the document (i.e. "Le document permet une navigation par structure et par pages. Les descriptions d'images sont incluses si elles étaient présentes dans le document Word.")
      - "Add a metadata page at the end of the EPUB" checkbox
      - "Add the metadata in the DOCX file" button
  - "Cover" tab:
    - a "Cover" option with the following choices :
      - "Don't use a cover image", 
      - "Include cover in metadata" 
      - "Include cover in content"
    - A "Source of cover image" option 
      - "Use document first page", 
      - "Select an image in the document" : that selects the first image available, displays it preview, and it activates 2 buttons to select the next and previous document image to be used and previewed, 
      - "DefaultCoverImage.png" : enables a "Select a cover image" button
      - "Title and author" : enables a "Generate a cover" button that opens a cover page generator form
  
  - "EPUB Options" tab:
    - "Language"
      - "Detected language" (displays the main language code detected folowwed by the list of other languages avaiables)
      - "Selected language" with a dropdown list with a list of possible language (Name and code)
    - "Text direction"
      - "Left to right"
      - "Right to left"
    - "Inlcude title page in line" (not sure of what this option do)
    - "Table of content"
      - "Depth" dropdown list, starts with value 3, values are in range 1 to 6.
      - "Insert TOC in content" checkbox, that activates a "TOC title" text box
        - "TOC title" text box
      - "Delete TOC of the word file"
    - "Split epub based on level"
      - possible values are an integer in range 1 to 3, default value is 2
  
  - "Pagination" tab:
    - "To générate pagination, use" option with the following choices :
      - "Word document pagination"
      - "Page number form headers and footers"
      - "From style Heading 6"
      - "From style Page Numbers (DAISY)"
      - "Included as PRINT PAGE #"
      - "No page numbers"'
    - "Include visible page numbers" checkbox that activate the 2 following text field
      - Text before the page numbers
      - Text after the page numbers
  
  - "Appearance" tab : 
    - "Type of appeareance" option has 2 choices:
      - "Select a stylesheet" activate an "Stylesheet" section of the form,
      - "Select option" value activate a "Style options" section of the form
    - "Stylesheet" section allows to select a css file with a "Browse" button and a "Reset" button
    -  "Style options" contains a set of selectable options:
      - "Text alignment" that can be one of "Justify", "Left aligned", "Right aligned" and "Not specified"
      - "Paragraphe style" can be either "Indent first line" of "No indentation"
      - "Table style" can be either "Grid lines" or "Horizontal lines"
      - "Font" is a dropdown list that contains the following list of values : 
        - "None"
        - "APHfont"
        - "AtkinsonHyperlisible"
        - "Delicious"
        - "Luciole"
        - "OpenDyslexic"
  
  - Headings tab:
    - provides a "heading report" that retrieve all levels and text of the headings in the content. with a button to open the report in a browser.


## Conversion

| Feature | WordToEPUB + Pandoc | SaveAsDAISY + DAISY Pipeline 2 |
| ------- | ------- | ------- |
| Preprocessing |  |  |
| - Vectorial shapes |  |  |
| - MathType equations |  |  |
| - textual markers |  |  |
| - Pagination options |  |  |
| Conversion to epub3 |  |  |
|  - Metadata |  |  |
|  - Content into epub3 |  |  |
| PostProcessing |  |  |
