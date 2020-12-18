# :: configure project here
param (
    [string]$dll = "J:\_DAISY\word-save-as-daisy\Word2003\DaisyTranslatorWord2003Shim\Debug\DaisyTranslatorWord2003Shim.dll"
);

# Search for an existing "signtool.exe" : 
# https://www.twelve21.io/using-signtool-exe-to-sign-a-dotnet-core-assembly-with-a-digital-certificate/

# Note, the windows kits is used by visual studio, i'm not sure of the path used by other windows SDK install
# it may be more appropriate to use a reg key on windows 10 :
# HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows Kits\Installed Roots
# Or HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Microsoft SDKs\Windows to get the different version installed
# and use the InstalationFolder key concatenated with ProductVersion

$programFiles = "";
if([System.Environment]::Is64BitOperatingSystem){
    $programFiles = ${env:ProgramFiles(x86)};
} else {
    $programFiles = $env:ProgramFiles;
}
$signtool = Get-ChildItem -Path $programFiles"\Windows Kits\10\bin\*\signtool.exe" -Recurse | Where-Object -FilterScript {$_.FullName -match "\\x86\\"} | Select-Object -Last 1;
if($signtool -eq $null){
    Write-Error "No Signing tool (signtool.exe) was found on your system, please install a windows 10 SDK.";
    exit 1;
} 

$makecert = Join-Path $(Split-Path $signtool) "makecert.exe";
$pvk2pfx = Join-Path $(Split-Path $signtool) "pvk2pfx.exe";

$projectName = "DAISY Translator for Microsoft Word";
$projectUrl = "https://github.com/daisy/word-save-as-daisy";
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path;

# expected pfx file from previous bat file
$certFile = Join-Path $ScriptDir "OdfConverter.pfx"
$certPassword = "sonata123"; 

if(-not (Test-Path $certFile)){
    # Create a new self signed certificate
    # https://medium.com/the-new-control-plane/generating-self-signed-certificates-on-windows-7812a600c2d8
    # https://docs.microsoft.com/en-us/powershell/module/pkiclient/new-selfsignedcertificate?view=win10-ps
    # Due to an error while trying to sign the app with the certificate created with power shell$
    # i use an alternate methode using tools provided with the SDK, following this stackoverflow post
    # https://stackoverflow.com/questions/16082333/why-i-get-the-specified-pfx-password-is-not-correct-when-trying-to-sign-applic
    Write-Warning "The expected odfConverter.pfx certificate was not found, checking for self signed certificate...";
    $certFile = Join-Path $ScriptDir "SelfSignedCertificate.pfx";
    if(-not(Test-Path $certFile)){
        if(Test-Path $ScriptDir\SelfSignedCertificate.pvk){
            
        }
        Write-Host "$makecert -sv `"$(Join-Path $ScriptDir "SelfSignedCertificate.pvk")`" -n `"CN=daisy.org`" `"$(Join-Path $ScriptDir "SelfSignedCertificate.cer")`" -r";
        Start-Process -Wait "$makecert" "-sv `"$(Join-Path $ScriptDir "SelfSignedCertificate.pvk")`" -n `"CN=daisy.org`" `"$(Join-Path $ScriptDir "SelfSignedCertificate.cer")`" -r";
        Write-Host "$pvk2pfx -pvk `"$(Join-Path $ScriptDir "SelfSignedCertificate.pvk")`" -spc `"$(Join-Path $ScriptDir "SelfSignedCertificate.cer")`" -pfx `"$certFile`" -po $certPassword";
        Start-Process -Wait "$pvk2pfx" "-pvk `"$(Join-Path $ScriptDir "SelfSignedCertificate.pvk")`" -spc `"$(Join-Path $ScriptDir "SelfSignedCertificate.cer")`" -pfx `"$certFile`" -po $certPassword";
    }
}


if(Test-Path $certFile){
    Write-Host "Signing $dll with $certFile ...";
    Write-Host "$signtool"+" sign /d `"$projectName`" /du $projectUrl /f `"$certFile`" /p $certPassword /v `"$dll`"" ;
    $signing = Start-Process -Wait "$signtool" "sign /d `"$projectName`" /du $projectUrl /f `"$certFile`" /p $certPassword /v `"$dll`"" ;
    if($signing.ExitCode -gt 0){
        Write-Error "An error have occured during the signing of $dll with $certFile";
    }else{
        Write-Host "$dll has been signed with $certFile (password $certPassword )";
    }

} else {
    Write-Error "No certificate were found nor could be created during the process";
}


# Write-Error "${signtool} sign /d $projectName /du $projectUrl /t $timestampServer /f $certFile /p $certPassword /v $dll";



