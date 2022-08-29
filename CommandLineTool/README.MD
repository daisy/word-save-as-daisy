Daisy Command Line Translator
===========================

Thank you for downloading Daisy Command Line Translator(DaisyConverterTest)

The purpose of the software is to enable users to open Daisy Specification based documents in Word 2007 and demonstrate the interoperability between OpenXML & Daisy specifications. 

The Daisy Command Line Translator is a tool for translating  Office Open XML format documents from the command line to Daisy format.

This software is still under development (see the roadmap on http://sourceforge.net/projects/openxml-daisy for more details about planned releases).

We will provide stable nightly builds. You can install and test them at your own risk!


Software Requirements
---------------------

To test the program, you need :

* .NET framework 3.0 


Command line Usage
----------------------

DaisyConverter.exe /I PathOrFilename [/O PathOrFilename] [BATCH-DOCX] [/REPORT Filename][/TITLE] [/CREATOR] [/PUBLISHER][/UID]		
    Where options are:
       /I PathOrFilename  Name of the file to transform (or input folder in case of batch conversion)(ex: D:\newfolder or D:\newfloder\input.xml)
       /O PathOrFilename  Name of the output file (or output folder)(ex: D:\newfolder or D:\newfloder\output.xml)
       /BATCH-DOCX        Do a batch conversion over every DOCX file in the input folder (note: existing files will be replaced)(ex: D:\newfolder)
       /REPORT Filename   Name of the report file that must be generated (existing files will be replaced)(ex: D:\newfloder\sample.txt)
       /TITLE      Data   Title of the output file that must be generated 
       /CREATOR    Data   Creator of the output file that must be generated 
       /PUBLISHER  Data   Publisher of the output file that must be generated 
       /UID        Data   Uid of the report output that must be generated  

Known Issues
------------

The DaisyTranslator command line tool only works on local drives.



OpenXML to Daisy conversion currently supports
--------------------------------------------
Dublin Core metadata: dc:Creator, dc:Title, dc:Publisher, dc:Date.
UID metadata.
Frontmatter, bookmatter and rearmatter.
Levels and headings
Paragraphs
Table
Lists
Notes and Note References
Emphasis and Strong
Superscript and subscript
Images
Captions
Validity



