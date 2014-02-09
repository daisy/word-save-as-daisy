:: called with the following parameters:
::     %1 - path to project folder
::     %2 - BuiltOutputPath, e.g. .\Debug\DaisyTranslatorWordAddin.msi
::     %3 - Configuration, e.g. "Release", "Release to Client(signed)"

:: add custom install actions
ECHO Adding custom actions
ECHO %1..\Scripts\AddCustomActions.vbs %2 ..\..\Lib\OdfInstallHelper.dll
cscript.exe %1..\Scripts\AddCustomActions.vbs %2 ..\..\Lib\OdfInstallHelper.dll


cscript.exe %1..\Scripts\CorrectConditions.vbs %2 %1..\Scripts\Conditions.xml

::sign the MSI file
if %3 == "Release to Client(signed)" CALL ..\..\signing\sign.bat %2

::build self-extracting installer
CALL %1..\Scripts\MakeSetupExe.bat %1%3 

::sign the self-extracting installer
if %3 == "Debug1" CALL ..\..\signing\sign.bat *.exe
