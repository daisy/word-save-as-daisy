@echo off

:: configure project here
set PROJECT_NAME=DAISY Translator for Microsoft Word
set PROJECT_URL=https://github.com/daisy/word-save-as-daisy

:: configure certificate here
:: CERT_FILE must contain the full path to the certificate, 
:: %~p0 expands to the path this script is in
set CERT_FILE=%~p0OdfConverter.pfx
set CERT_PASSWORD=sonata123

:: optional parameter for timestamping
set TIMESTAMP_SERVER=http://timestamp.verisign.com/scripts/timstamp.dll

:: path to signtool.exe
set SIGNTOOL=C:\Program Files (x86)\Windows Kits\10\bin\10.0.18362.0\x86\signtool.exe

IF NOT EXIST "%SIGNTOOL%" goto :error

@echo on 
"%SIGNTOOL%" sign /d "%PROJECT_NAME%" /du "%PROJECT_URL%" /t "%TIMESTAMP_SERVER%" /f "%CERT_FILE%" /p "%CERT_PASSWORD%" /v "%~f1"
@echo off

goto :eof

:error
echo ERROR: signtool.exe not found (Please configure the path correctly in file %~f0).

:eof