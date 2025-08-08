# SaveAsDAISY build script
# Requires msbuild to be accessible in path

param(
	[string]$version = "",
	[switch]$refreshpipeline = $false,
    [switch]$nobuild = $false,
    [switch]$debug = $false
)

$currentVersion = "2.9.2"
$wixProductPath = Join-Path $PSScriptRoot "Installer\DaisyAddinForWordSetup\Product.wxs"

# Create the wix directory tree for a path
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

if($version) {
    if($version -match '\d+\.\d+\.\d+'){
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
    } else {
        Write-Host "$version is not a major.minor.patch version string"
    }

}

if($refreshpipeline) {
    # recompute the pipeline using the engine make tool and a makefile in the root folder
    Start-Process -WorkingDirectory $PSScriptRoot -FilePath $(Join-Path $PSScriptRoot "engine\make.exe") -ArgumentList "clean" -Wait
    Start-Process -WorkingDirectory $PSScriptRoot -FilePath $(Join-Path $PSScriptRoot "engine\make.exe") -Wait
    Start-Sleep 1
    $_oldroot = Join-Path $PSScriptRoot "resources"
    # regenerate and update the wix project "product.wxs"
    # - compute wix components and references for daisy-pipeline folder in Lib
    $_cont, $_refs = Update-WixTree `
        -oldroot $_oldroot `
        -path $(Join-Path $_oldroot "daisy-pipeline") `
        -newroot "`$(var.SolutionDir)resources" `
        -mediaId 2 `
        -level 6 `
        -indent "    "
    # Replace the text in range between markers (also replacing start markers)
    $refMarker = "<!--daisy-pipeline refs-->"
    $dirMarker = "<!--daisy-pipeline-->"


    $wixProductContent = Get-Content -Raw $wixProductPath
    $dirMarkerStart = $wixProductContent.IndexOf("<Directory Id=`"_daisy_pipeline`"")
    $dirMarkerEnd = $wixProductContent.IndexOf($dirMarker) + $dirMarker.Length #because this marker is readded by the print-wix function
    $refMarkerStart = $wixProductContent.IndexOf("<ComponentRef Id=`"_daisy_pipeline_files`"/>")
    $refMarkerEnd = $wixProductContent.IndexOf($refMarker)
    $beforeDirectoryStart = $wixProductContent.Substring(0, $dirMarkerStart)
    $betweenDirectoryEndAndRefStart = $wixProductContent.Substring($dirMarkerEnd + 1, $refMarkerStart - $dirMarkerEnd - 1)
    $afterRefEnd = $wixProductContent.Substring($refMarkerEnd)
    if(($dirMarkerStart -gt -1) -and ($dirMarkerEnd -gt -1) -and ($refMarkerStart -gt -1) -and ($refMarkerEnd -gt -1)){
        Set-Content -Path $wixProductPath `
            -Value "$beforeDirectoryStart$_cont$betweenDirectoryEndAndRefStart$_refs$("    " * 3)$afterRefEnd" `
            -Encoding UTF8
    } else {
        Write-Host "Can't update wix project, could not find every markers in content'"
    }
    
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
    # build the MSIs
    MSBuild.exe DaisyConverter.sln /t:clean /p:Configuration="Release" /p:Platform="x86";
    MSBuild.exe DaisyConverter.sln /t:restore /p:Configuration="Release" /p:Platform="x86";
    MSBuild.exe DaisyConverter.sln /t:Installer\DaisyAddinForWordSetup /p:Configuration="Release" /p:Platform="x86";

    MSBuild.exe DaisyConverter.sln /t:clean /p:Configuration="Release" /p:Platform="x64";
    MSBuild.exe DaisyConverter.sln /t:restore /p:Configuration="Release" /p:Platform="x64";
    MSBuild.exe DaisyConverter.sln /t:Installer\DaisyAddinForWordSetup /p:Configuration="Release" /p:Platform="x64";
    # build the installer
    MSBuild.exe DaisyConverter.sln /t:Installer\SaveAsDAISYInstaller /p:Configuration="Release" /p:Platform="Any CPU" /p:DefineConstants="UNIFIED";
}
