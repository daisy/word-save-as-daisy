# SaveAsDAISY build script
# Requires msbuild to be accessible in path

param(
	[string]$version = "",
	[switch]$refreshpipeline = $false,
    [switch]$nobuild = $false,
    [switch]$debug = $false,
    [switch]$beta = $false
)

$currentVersion = "2.9.4.2"
$wixProductPath = Join-Path $PSScriptRoot "Installer\DaisyAddinForWordSetup\Product.wxs"

$newWixFragmentPath = Join-Path $PSScriptRoot "Installer\DaisyAddinMSIPackage\EngineComponents.wxs"
function Update-WixTree {
    param (
        [string]$oldroot,
        [string]$newroot,
        [string]$path,
        [string]$mediaId,
        [int]$level = 0,
        [string]$indent="    ",
        [int]$startIndent=0
    )

    $content = ''
    $refs = ''
    $path_end = $path.Substring($oldroot.Length)
    $base_id = $path_end.Replace('\', '_').Replace('-', '_').Replace('$', '_').Replace(' ', '_')
    
    if (Test-Path $path -PathType Leaf) {
        $content = ($indent * $level * $startIndent) + "<File DiskId=`"$mediaId`" Id=`"$base_id`" Name=`"$([System.IO.Path]::GetFileName($path))`" Source=`"$($newroot + $path_end)`" />`n"
        return $content, $refs
    }
    else {
        $content = ($indent * $level * $startIndent) + "<Directory Id=`"$base_id`" Name=`"$([System.IO.Path]::GetFileName($path))`">`n"
        $subpaths = Get-ChildItem -Path $path
        $files = $subpaths | Where-Object { !$_.PSIsContainer }
        $directories = $subpaths | Where-Object { $_.PSIsContainer }

        if ($files.Count -gt 0) {
            $id = "$($base_id)_files"
            $content += ($indent * ($level + 1)) + "<Component Id=`"$id`" Guid=`"$([guid]::NewGuid())`">`n"
            $refs += ($indent * 3 * $startIndent) + "<ComponentRef Id=`"$id`"/>`n"

            foreach ($file in $files) {
                $_cont, $_refs = Update-WixTree -path $file.FullName -mediaId $mediaId -level ($level + 2) -oldroot $oldroot -newroot $newroot -indent $indent -startIndent 1
                $content += $_cont
                $refs += $_refs
            }

            $content += ($indent * ($level + 1)) + "</Component>`n"
        }

        foreach ($directory in $directories) {
            $_cont, $_refs = Update-WixTree -path $directory.FullName -mediaId $mediaId -level ($level + 1) -oldroot $oldroot -newroot $newroot -indent $indent -startIndent 1
            $content += $_cont
            $refs += $_refs
        }

        $content += ($indent * $level) + "</Directory><!--$([System.IO.Path]::GetFileName($path))-->`n"
        return $content, $refs
    }
}

function Update-Version {
    param (
        [string]$path,
        [string]$oldVersion,
        [string]$newVersion
    )

    $excludeDirs = @('bin', 'obj')

    Get-ChildItem -Path $path -Recurse -File | ForEach-Object {
        $exclude = $false
        foreach ($dir in $excludeDirs) {
            if ($_.DirectoryName -match "\\$dir\\") {
                $exclude = $true
                break
            }
        }

        if (!$exclude -and $(Select-String -Path $_.FullName -Pattern $oldVersion -SimpleMatch)) {
            Write-Host "Updating>" $_.FullName
            (Get-Content $_.FullName) |
            Foreach-Object { $_.Replace($oldVersion, $newVersion )} |
            Set-Content $_.FullName -Encoding UTF8
        }
    }
}

# From https://rkeithhill.wordpress.com/2009/08/03/effective-powershell-item-16-dealing-with-errors/
function CheckLastExitCode {
    param ([int[]]$SuccessCodes = @(0), [scriptblock]$CleanupScript=$null)

    if ($SuccessCodes -notcontains $LastExitCode) {
        if ($CleanupScript) {
            "Executing cleanup script: $CleanupScript"
            &$CleanupScript
        }
        $msg = @"
EXE RETURNED EXIT CODE $LastExitCode
CALLSTACK:$(Get-PSCallStack | Out-String)
"@
        throw $msg
    }
}

if($version) {
    if($version -match '\d+\.\d+\.\d+\.\d+'){
        #   - search and replace occurences of previous versions by new version
        # Search in replace in the required parts of the project
        Update-Version -path $(Join-Path $PSScriptRoot "Common") -oldVersion $currentVersion -newVersion $version
        Update-Version -path $(Join-Path $PSScriptRoot "WordAddin") -oldVersion $currentVersion -newVersion $version
        Update-Version -path $(Join-Path $PSScriptRoot "Installer") -oldVersion $currentVersion -newVersion $version
        Update-Version -path $(Join-Path $PSScriptRoot "CustomActionAddin") -oldVersion $currentVersion -newVersion $version
        
        Write-Host "Updating Product and Package GUIDs in" $wixProductPath
        $productUID = $(Select-String -Path $wixProductPath -Pattern "Product Id=`"(?<uid>[^`"]*)`"").Matches[0].Groups['uid'].Value
        $packageUID = $(Select-String -Path $wixProductPath -Pattern "Package Id=`"(?<uid>[^`"]*)`"").Matches[0].Groups['uid'].Value
        (Get-Content $wixProductPath) |
            Foreach-Object { $_.Replace($productUID, [guid]::NewGuid().toString().ToUpper()).Replace($packageUID, [guid]::NewGuid().toString().ToUpper())
        } | Set-Content $wixProductPath -Encoding UTF8
        # self update the scrip current version
        Write-Host "Updating version in build script" $PSCommandPath
        (Get-Content $PSCommandPath) |
            Foreach-Object { $_.Replace($currentVersion, $version ) } |
            Set-Content $PSCommandPath -Encoding UTF8

        $currentVersion = $version
    } else {
        Write-Host "$version is not a major.minor.patch.build version string"
    }

}


if($refreshpipeline) {
    # NP 2026 03 26 : i'm removing the engine submodule from the project
    # - The submodule repository is blocking page generation du to an internal submodule error
    # - The build time of the pipeline repository has extensively increased (between 30min to several hours), so rebuilding it at each build is not viable anymore
    # For now, the recommended process to update the pipeline is to separatly 
    # - clone my fork of the pipeline repository (with improved SimpleAPI class for jni interaction, to be merged in the original repository through a PR) : https://github.com/NPavie/pipeline.git
    # - within this fork, launch the command `.\assembly\make.exe FIXED_BUILD=true dist-zip-sad` to build the pipeline archive for SaveAsDAISY
    # - copy the generated "daisy-pipeline" folder of the archive in the project resources/ folder
    # When the JNI wrapper is built, a post-process action will generate a description file of settable properties in the resources folder (that is required for embedded engine quick access to these properties)
    # NP 2025 08 08 : problem with the new build system introduced by the pipeline ui (cleaning and rebuild raises errors, but initial build works)
    # For now, i redo the build manually when needed
    # Start-Process -WorkingDirectory $PSScriptRoot -FilePath $(Join-Path $PSScriptRoot "engine\assembly\make.exe") -ArgumentList "clean" -Wait
    # Start-Process -WorkingDirectory $PSScriptRoot -FilePath $(Join-Path $PSScriptRoot "engine\assembly\make.exe") -Wait
    # Start-Sleep 1

    # replace the resources\daisy-pipeline\etc\pipeline.properties by the resources\pipeline.properties generated by the pipeline build
    # (allows to use the embedded engine on localfs and bound it to port 49255 when running as a webservice)
    Copy-Item -Path (Join-Path $PSScriptRoot "resources\pipeline.properties") -Destination (Join-Path $PSScriptRoot "resources\daisy-pipeline\etc\pipeline.properties") -Force

    # Regen the settable-properties.xml descriptor by building the JNIWrapper and then running it with the "settable-properties --output "resources\daisy-pipeline\settable-properties.xml" argument
    MSBuild.exe DaisyConverter.sln /t:Common\JNIWrapper /p:Configuration="Release" /p:Platform="x64";
    CheckLastExitCode
    Write-Debug "Generating settable-properties.xml description file for daisy-pipeline embedded engine usage"
    # We don't need to copy the pipeline in the result, a post process action of the JNIWrapper project will copy the pipeline in its release folder for autonomous usage.
    Start-Process -FilePath (Join-Path $PSScriptRoot "Common\JNIWrapper\bin\x64\Release\JNIWrapper.exe") -ArgumentList "settable-properties --output `"$($PSScriptRoot)\resources\daisy-pipeline\settable-properties.xml`"" -Wait
    CheckLastExitCode

    #REM -- Update the pipeline with the preextracted settable properties description file
    #cmd /c $(TargetPath) settable-properties --output "$(SolutionDir)resources\daisy-pipeline\settable-properties.xml"
    Write-Debug "Updating the wix project embedded engine components with the content of the daisy-pipeline folder in resources"
    $_oldroot = Join-Path $PSScriptRoot "resources"
    # regenerate and update the wix project "product.wxs"
    # - compute wix components and references for daisy-pipeline folder in Lib
    $_cont, $_refs = Update-WixTree `
        -oldroot $_oldroot `
        -path $(Join-Path $_oldroot "daisy-pipeline") `
        -newroot "`$(var.SolutionDir)resources" `
        -mediaId 2 `
        -level 3 `
        -indent "    "
    
    # Replace content of the engine wix fragment
    Set-Content -Path $newWixFragmentPath `
        -Value "<Wix xmlns=""http://wixtoolset.org/schemas/v4/wxs"">
  <Fragment>
	  <Media Id=""2"" Cabinet=""pipeline.cab"" EmbedCab=""no"" />
	  <DirectoryRef Id=""APPLICATIONFOLDER"">
		  $_cont
	  </DirectoryRef>
	  <ComponentGroup Id=""EmbeddedEngineFiles"" Directory=""APPLICATIONFOLDER"" >
		  $_refs
	  </ComponentGroup>
  </Fragment>
</Wix>
" `
        -Encoding UTF8
        
    
}

if($nobuild) {
    exit
} elseif ($debug) {
    # Stop Microsoft Word if it is running
    $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    if ($wordProcesses) {
        Write-Host "Stopping Microsoft Word..."
        Stop-Process -Name "WINWORD" -Force
    } else {
        Write-Host "Microsoft Word is not running."
    }
    # build the addin project and its dependencies, and copy the result to %LOCALAPPDATA%\Apps\Save-as-DAISY Word Addin
    MSBuild.exe DaisyConverter.sln /t:Word\DaisyWord2007Addin /p:Configuration="Debug" /p:Platform="x64";
    # copy the result in WordAddin\bin\Debug\x64 to the local app data folder
    $addinPath = Join-Path $PSScriptRoot "WordAddin\bin\x64\Debug"
    $localAppDataPath = Join-Path $env:LOCALAPPDATA "Apps" "Save-as-DAISY Word Addin"
    Copy-Item -Path $addinPath\* -Destination $localAppDataPath -Recurse -Force
    Start-Process WINWORD

} else {
    # build the x86 MSIs
    MSBuild.exe DaisyConverter.sln /t:clean /p:Configuration="Release" /p:Platform="x86";
    CheckLastExitCode
    MSBuild.exe DaisyConverter.sln /t:restore /p:Configuration="Release" /p:Platform="x86";
    CheckLastExitCode
    MSBuild.exe DaisyConverter.sln /t:Installer\DaisyAddinMSIPackage /p:Configuration="Release" /p:Platform="x86";
    CheckLastExitCode

    # build the x64 MSIs
    MSBuild.exe DaisyConverter.sln /t:clean /p:Configuration="Release" /p:Platform="x64";
    CheckLastExitCode
    MSBuild.exe DaisyConverter.sln /t:restore /p:Configuration="Release" /p:Platform="x64";
    CheckLastExitCode
    MSBuild.exe DaisyConverter.sln /t:Installer\DaisyAddinMSIPackage /p:Configuration="Release" /p:Platform="x64";
    CheckLastExitCode

    # build the installer
    MSBuild.exe DaisyConverter.sln /t:Installer\SaveAsDAISYInstaller /p:Configuration="Release" /p:Platform="Any CPU" /p:DefineConstants="UNIFIED";
    CheckLastExitCode

    
    #rename the installer and move it to the Release folder
    $installerName = "SaveAsDAISYInstaller-$currentVersion"
    if ($beta) {
        $installerName += "-beta"
    }
    $installerName += ".exe"
    if (!(Test-Path (Join-Path $PSScriptRoot "Release"))) {
        New-Item -ItemType Directory -Path (Join-Path $PSScriptRoot "Release") | Out-Null
    }
    Copy-Item -Path (Join-Path $PSScriptRoot "Installer\SaveAsDAISYInstaller\bin\Release\SaveAsDAISYInstaller.exe") -Destination (Join-Path $PSScriptRoot "Release\$installerName") -Force

    Write-Host "Build completed. Installer available at:" (Join-Path $PSScriptRoot "Release\$installerName")
}
