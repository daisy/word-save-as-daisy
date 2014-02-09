rem %1 - $(TargetDir)
rem %2 - $(TargetFileName) = DaisyAddinForWordSetup.msi
rem %3 - $(ConfigurationName)

echo Building self-extracting installer
echo IExpress /Q /N ..\..\scripts\%3\DaisyAddinForWordSetup.sed
IExpress /N ..\..\scripts\%3\DaisyAddinForWordSetup.sed

rem echo Signing MSI files
rem CALL ..\..\..\signing\sign.bat *.exe
rem call %1..\..\..\signing\sign.bat "%1%2"

