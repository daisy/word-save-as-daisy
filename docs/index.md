---
title: Download SaveAsDAISY Word Addin
layout: my-default
---
SaveAsDAISY is available for Windows only. 
We provide an unified installer that detects standard installations of Microsoft Office.
If Office is not found by the installer (like preinstalled or windows store versions of MS Office), it will request which architecture (32bits or 64bits) you want the addin to be installed for.


## Latest stable version: 2.7.2 beta (released on September, 2022)

- [Download universal installer](https://github.com/daisy/word-save-as-daisy/releases/download/v2.7.2-beta/SaveAsDAISYInstaller.exe)
- Or access to the [last release page](https://github.com/daisy/word-save-as-daisy/releases/latest)
- Or test the [latest beta version](https://github.com/daisy/word-save-as-daisy/releases/tag/v2.8.6-beta)

## Report issues

If you have an issue with the installers or with the add-in, please contact the development team by mail to [daisy-pipeline@mail.daisy.org](mailto:daisy-pipeline@mail.daisy.org), or check our [github repository](https://github.com/daisy/word-save-as-daisy/issues).

Any constructive feedbacks are also welcome to help us improve the add-in :-)

## Feedbacks / Known issues

### The accessibility ribbon does not appear after successfull installation

#### Addin is reported unloaded due to an error

The addin assumes you have a recent version of .NET Framework (4+) available on your system.
You can check if you have the required .NET Framework by checking if folders "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"  and/or "C:\Windows\Microsoft.NET\Framework\v4.0.30319" exists.

It is also possible that the wrong bit-version of the addin is installed : 
this might have been the case if your office version has been installed through Microsoft Store, which makes office undetectable from our install process, and you selected the wrong bit-version.
If this is the case, please check your office bit-version in Word in the "Account" page by opening "About Word" dialog.
The bit-version is displayed at the end of the first line of the dialog.

Lastly, it has been reported for other addin that this can occured if you run word on an account with administrator privileges with UAC deactivated.
In this case, you need to do the following actions : 
- uninstall the addin
- Launch the addin installed as administrator
- When accepting the software licence, click on the "Advanced" button
- Select "Install for all users of this machine" and click "Next"
- Change the installation folder if you need and click "Next'
- The addin should then continue the install process as usual

#### Addin is deactived after installation

A user reported the addin to be deactivated after installation : 
In Word, under the **File** / **Options** / **Add-ins** category, **Save As Daisy** appeared under the **Inactive applications** section.

This issue is being investigated, but the following action has been tested and reported working by the user encountering this issue.
- Open the **Manage** panel, select **COM Add-ins** in the combo box and click the **Go** button, 
- Check the **Save as Daisy** add-in and click on the **Add** button.

The add-in should start repairing himself before reopening the word document.

## Speech synthesis issue on Windows 11

On Windows 11 (but not Windows 10), Microsoft seems to have limited or missed a bug that forbid the use of their local text-to-speech engines :
Both legacy SAPI engine and its newer Onecore version are impacted.
Using one of those text-to-speech engine in parallel in one or more applications can often make one or both app crash.
For example: launching an export to DAISY3 with audio synthesis in the DAISY Pipeline 2, and navigating with nvda or jaws with Onecore voices, 
will sometimes make either DAISY pipeline 2 or NVDA/JAWS or both crash.
The crash is also unstoppable on our side when it happens, as it uses special instructions (which is called a FAST_FAIL) that cannot be intercepted when it occurs.

We have mitigated the issue in the DAISY Pipeline 2 for multithreaded audio synthesis by synchronizing the speech synthesis call at the lowest level we are able to.
This reduced the performances of the synthesis but it avoids quasi-systematic unrecoverable crashes in the DAISY Pipeline 2 conversions.
We then discovered that those crashes can still occur during a conversion if a narrator like nvda and jaws with Onecore voices selected is used during conversion.

This has been reported to Microsoft through their Feedback channel, but if anyone has contacts at Microsoft and can help propagate this issue to Microsoft Dev teams, 
we would really appreciate it.

# Changelog

## 2.8.6 beta (June 2024)

This release includes the following changes :
- Non-docx document opened in word are now accepted, as long as word can create a copy of them and upgrade it to docx format.
- Language detection has been reworked to better detect the different languages used in the document and fix paragraph and text language markup.
  - Credits to @ways2read / Richard Orme for providing the code used in WordToEPUB
- The underlying DAISY Pipeline 2 has been updated:
  - The form to export the document to mp3 has been updated to reflect the new parameters of the pipeline 2 script.
- Fixed cases where acronyms and abbreviations were breaking validity of the result.
- Fixed an issue where a title in table was converted as a list (for now converted as paragraph)
- Fixed an issue with pagenum generation where closing paragraph tags would be added in some cases without an opening tag.

## 2.8.5 beta (March 2024)

This minor release includes the following changes :
- Clarify an error message for exception raised by the Microsoft package parsing libraryr when navigating in a document: 
  - The parser can crash in presence of certain malformed hyperlinks, i. e. if an hyperlink host adress contains a coma.
- For Non-docx document, the working copy used for the conversion is now also automatically upgraded to docx format before conversion.
- For new users, the output folder is now set by default to the `%USERPROFILE%\Documents\SaveAsDAISY Results` folder.
  - The output folder is also automatically created if not found. 


## 2.8.4 beta (February 2024)

This minor release includes the following changes :
- For automatic pagination, inline pagenum elements in paragraph (fixes [#40](https://github.com/daisy/word-save-as-daisy/issues/40))
  - Note that automatic pagination is not recommended in the current state of the addin :
    In our tests, Word does not register all automatic pagebreak in the underlying content, only a few one that were rendered while saving the document.
- A new language selector is provided in conversion parameters as to better select the main language of the document (fixes [#39](https://github.com/daisy/word-save-as-daisy/issues/39))
  - This is needed by the DAISY Pipeline 2 to select the correct language for the synthesis of the audio files.
- Fixed a language test for bidirectionnal elements


## 2.8.3 beta (January 2024)

This minor release includes the following changes :
- Fixed an issue with inlined element not being closed correctly before inserting pagenum
- The underlying DAISY Pipeline 2 has been update to version 1.14.17-p1 

## 2.8.2 beta (December 2023)

This major release marks a change in the underlying conversion process 
to dtbook XML and packaged format,
as DAISY Pipeline 1 is completely replaced by DAISY Pipeline 2.
The following features that are either reported unused or that are available 
through standard Microsoft Word actions are removed from the addin :

- The "from multiple documents" exports actions are removed : 
As it was not a batch conversion of documents but a conversion and merging 
feature of multiple word document into a single Dtbook XML or DAISY book, 
testers reported the feature as unused in its current state.
For users needing a similar action, it is recommended to either use 
a master document referencing the documents to merge, or manually 
construct a new Word document from the documents to merge.
- The "Add footnotes" button is removed from the ribbon :
it is recommended to use Microsoft Word notes management 
in the "References" ribbon's tab.
- It is no longer required to select an empty folder as export destination :
For a given export to a selected format, a new folder taking the name of 
the converted file name followed by the selected format and a timestamp suffix
will be created inside the selected destination folder and will contain
the result of the conversion.

Some new features are starting to be integrated in the addin and are still experimental:
- With the updated DAISY Pipeline 2 provided, users can now test the export 
to Megavoice fileset of MP3 files (Issues [#30](https://github.com/daisy/word-save-as-daisy/issues/30)


Various fixes and changes are included in the release :
- Fixed [#25](https://github.com/daisy/word-save-as-daisy/issues/25)
by caching dtbook and mathml dtds and entities in the assembly,
- New conversion form "advanced settings" panel to better display settings
for each export format based on DAISY pipeline 2 settings
- Fixed conversion progress dialog text area behavior
- Fixed inline shapes export that could raise a clipboard issue
- Fixed a bug in acronyms and abbreviations conversions
- New update checking mecanism in the "About" dialog that 
looks at the daisy github repository of the addin.
- Issue [#26](https://github.com/daisy/word-save-as-daisy/issues/26) is fixed by
the removal of pipeline 1 that is vulnerable to attacks through its version of log4j
- Including an improved exception reporting when conversion errors occurs.

## 2.7.2 beta (September 2022)

This minor release adds the following changes to the addin : 
- feature [#19](https://github.com/daisy/word-save-as-daisy/issues/19) : Adding notes positionning options in settings
  - Notes can be positionning at the end of pages, inlined after or at the end of the level of the paragraph holding the noteref
  - An insertion level can be selected for inlining notes in level or at the end of level
- Fixed [#18](https://github.com/daisy/word-save-as-daisy/issues/13) : Consecutive texts in italic, strong or exponent are now merged into one text block
- Fixed [#13](https://github.com/daisy/word-save-as-daisy/issues/13) : hyperlink targets in notes are now correctly retrieved
- Changed the installer to allow install in user-space for windows 10
  - User without administrator privileges will need to contact their IT administrator to update the program the first time and switch to user-space installation if they are allowed to


## 2.7 beta (January 2022)

This major release starts the transition to the DAISY pipeline 2 as conversion engine.
We now include the export to Epub3 (from a single docx file) as experimental functionnality under the SaveAsDAISY menu.

Fixes and updates :
- The installers are now completely unified to ease releases management : 
User with "hidden" office installation (like office installed from the Microsoft Store) are now prompted with a version selector on launch, including a link
to help them find the bit version of their office if needed.
- Progression of the conversion is monitored through a progress bar and conversion messages to better visualize the state of the conversion
- Fixed a mathml handling issue
- Updated the pipeline 1 distribution to use 32bit java instead of 64bit :
Users reported that voices provider are mainly distributing 32bits SAPI voices, thus requiring 32bit java for their launch.
- Conversion controls are now disabled in protected view (that limits word actions including plugin ones on documents)

## 2.6.1 beta - update 3 (September 2021)

Minor behaviour updates including :
- bugfix for [#6](https://github.com/daisy/word-save-as-daisy/issues/6)
- bugfix for an issue regarding a PushLevel error message
- Conversion progress bars are replaced by a Message dialog reporting more precisely the process status

The code base is under a heavy rewrite process to optimize it and prepare the switch to pipeline 2 process,
allowing to provide more outputed format and better support.


## 2.6.1 beta - Minor update 2 (January 28, 2021)

Minor update including
- bugfix for [#2](https://github.com/daisy/word-save-as-daisy/issues/2) in test phase
- bugfix for [#3](https://github.com/daisy/word-save-as-daisy/issues/3) in test phase
- Documentation updated (new manual and authoring guidelines, including integrated help file)
- Added a document file name check pass before launching conversion (to let the user manually rename his file if wanted)
- Fixed an installer issue where previous version of packages could be used instead of the new ones
- Name and logos updated in software and installers




## 2.6.1 beta - Minor update 1 (January 14, 2021)

Minor update including
- bugfix for [#1](https://github.com/daisy/word-save-as-daisy/issues/1) in test phase
- Removed the Accessibility Checker shortcut in the addin ribbon for office 2010 and later
Users are advised to validated using the accessibility checker before launching a conversion in the conversion setting form, with a link to the checker documentation.
- Updated the AdoptOpenJDK java 8 runtime provided with the embedded DAISY pipeline lite

### Feedbacks

- File not found errors with partial file names reported missing:
Some characters like spaces and commas may lead to "File not found" errors in the conversion process. Replace those characters by underscore "_" and launch the conversion again.

## 2.6.1 beta (December 18, 2020)

## Installer update

- Office XP is removed from the list of supported versions of Office,
- The list of officially supported versions of office now goes from Office 2003 to Office 2016/365,
  - For newer version, a warning is raised to let the user decide if the installation must continue or not.
- For Windows x64, a unified installer is provided for Office 32bits and 64Bits
- For undetected Office installations (OEM or Windows store install of Office are preventing version detection):
  - For office 32bits, the unified installer can be used,
  - For office 64bits, a separate installer is provided.
- The default installation directory is now in program files/DAISY/Save-as-DAISY Word Addin,
- The latest version available of the DAISY pipeline 1 is included in the addin installer, and is shipped with an integrated java runtime for Windows x64

## Features

- The Word validation button is being replaced by a shortcut to the Microsoft Accessibility Checker, available in office 2010 and newer.
  - The previous word validation process is only kept for word 2007 and 2003 where the accessibility checker is not available,
- The conversion of word documents to DAISY XML now includes a post-process pass:
  - User can decide to apply cleanups and sentences detection in the conversion process options

## Feedbacks

- An issue has been reported with document containing shapes or images, where the conversion would not launch. The cause has been identified and is being worked on.
- A user reported a "Not able to find directory to install" warning when installing the addin for word 32 bits on x64 system. The install directory was found after going one step back and clicking again on next.
