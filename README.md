# "Save as DAISY" Add-In for Microsoft Word

The "Save as DAISY" Add-In for Microsoft Word is a collaborative standards-based development project initiated in 2007 by the DAISY Consortium and Microsoft.
It was initially developed by [Sonata Software](http://www.sonata-software.com/) (versions 1.x), then further developed by [Intergen](http://intergen.co.nz/) (up to version 2.5.5.1).

The code available in this GitHub project has been initially copied from its original location on SourceForge as the [openxml-daisy project](http://sourceforge.net/projects/openxml-daisy/).

## Building requirement

The project is being updated to build with visual studio 2026.

The MSI packages are build using a wix toolset project, and requires [Wix toolset version 6](https://github.com/wixtoolset/wix/releases/tag/v6.0.2) to be installed.
We recommend to also add the Wix toolset [Heatwave](https://marketplace.visualstudio.com/items?itemName=FireGiant.FireGiantHeatWaveDev17) extension in Visual Studio.

To ease the debugging and building of releases, a `build.ps1` powershell script is provided.
We recommend to have powershell 7+ installed (instead of the default powershell 5 usually provided by Microsoft), that can be installed with the following command:
`winget install --id Microsoft.PowerShell --source winget`

This scripts requires to have the MSBuild tools directory registered in your PATH variable, and we recommend to also register the Wix toolset utilites in it.

For short, when using visual studio 2026 and Wix toolset v6, you can add the following entries in your path
 - `%PROGRAMFILES%\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin`
 - `%PROGRAMFILES%\WiX Toolset v6.0\bin\x64`
 - `%PROGRAMFILES%\WiX Toolset v6.0\bin`


## Build the current release

With all the requirement install, just launch `build.ps1` and it should launch the packaging process.
The final installer should be then found in the `Installer\SaveAsDAISYInstaller\bin\Release` on success

You can update the version of the release package by using the flag `-version` followed by the new version code (in regex format `\d+.\d+.\d+.\d+`)
For example, `build.ps1 -version 2.9.4.2` will replace the existing version code in the whole solution and build a SaveAsDAISY installer with this version code.

## Updating the embedded engine

The embedded daisy pipeline version for SaveAsDAISY can be found here : https://github.com/NPavie/pipeline/tree/save-as-daisy
Just clone the repo somewhere, install the requirements (java8 JDK and maven) and use the following command from the pipeline repository root folder
`.\assembly\make.exe FIXED_BUILD=true dist-zip-sad`

After replacing the `resources\daisy-pipeline` it is recommended to do a quick debug build of the `Common\JNIWrapper` project, to get the updated pipeline properties descriptor file and ensure it works.

You can then use the command `build.ps1 -refreshPipeline -nobuild` to update the wix project (but without triggering the whole build with the `-nobuild` flag to gain some time)

## Debugging the addin

If you need to debug the addin, you need to first install the last beta release available.
After that you can just close Microsoft word if it is open and launch `build.ps1 -debug` to build the addin in debug mode.
The command will then open word, and you can debug it by attaching Visual studio to the word instance.

Note that the debug command will also redeploy the embedded pipeline in the addin install repository.

## Notes

In case of multiple crash of the addin, it might be registered in Word's addin blocklist.
`HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItem`


