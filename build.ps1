# SaveAsDAISY build script
# Requires msbuild to be accessible in path

param(
	[string]$version = "",
	[switch]$refreshpipeline = $false
)

function Print-WixTree {
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
                $_cont, $_refs = Print-WixTree -path $file.FullName -mediaId $mediaId -level ($level + 2) -oldroot $oldroot -newroot $newroot -indent $indent -startIndent 1
                $content += $_cont
                $refs += $_refs
            }

            $content += ($indent * ($level + 1)) + "</Component>`n"
        }

        foreach ($directory in $directories) {
            $_cont, $_refs = Print-WixTree -path $directory.FullName -mediaId $mediaId -level ($level + 1) -oldroot $oldroot -newroot $newroot -indent $indent -startIndent 1
            $content += $_cont
            $refs += $_refs
        }

        $content += ($indent * $level) + "</Directory><!--$([System.IO.Path]::GetFileName($path))-->`n"
        return $content, $refs
    }
}

function Replace-Text {
    param (
        [string]$path,
        [string]$oldText,
        [string]$newText
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

        if (!$exclude) {
            (Get-Content $_.FullName) |
            Foreach-Object { $_ -replace $oldText, $newText } |
            Set-Content $_.FullName
        }
    }
}

if($version) {
#   - search and replace occurences of previous versions by new version
# Search in replace in all files 
    
}

if($refreshpipeline) {
    # regenerate and update the wix project "product.wxs"
    # - compute wix components and references for daisy-pipeline folder in Lib
    $_oldroot = Join-Path $PSScriptRoot "Lib"
    $_cont, $_refs = Print-WixTree `
        -oldroot $_oldroot `
        -path $(Join-Path $_oldroot "daisy-pipeline") `
        -newroot "`$(var.SolutionDir)Lib" `
        -mediaId 2 `
        -level 6 `
        -indent "    "
    # Replace the text in range between markers (also replacing start markers)
    $wixProductPath = Join-Path $PSScriptRoot "Installer\DaisyAddinForWordSetup\Product.wxs"
    $wixProductContent = Get-Content -Raw $wixProductPath
    $dirMarkerStart = $wixProductContent.IndexOf("<Directory Id=`"_daisy_pipeline`"")
    $dirMarkerEnd = $wixProductContent.IndexOf("<!--daisy-pipeline-->") + "<!--daisy-pipeline-->".Length #this is added by the print-wix function
    $refMarkerStart = $wixProductContent.IndexOf("<ComponentRef Id=`"_daisy_pipeline_files`"/>")
    $refMarkerEnd = $wixProductContent.IndexOf("<!--daisy-pipeline refs-->")
    $beforeDirectoryStart = $wixProductContent.Substring(0, $dirMarkerStart)
    $betweenDirectoryEndAndRefStart = $wixProductContent.Substring($dirMarkerEnd + 1, $refMarkerStart - $dirMarkerEnd - 1)
    $afterRefEnd = $wixProductContent.Substring($refMarkerEnd)
    if(($dirMarkerStart -gt -1) -and ($dirMarkerEnd -gt -1) -and ($refMarkerStart -gt -1) -and ($refMarkerEnd -gt -1)){
        Set-Content -Path $wixProductPath -Value "$beforeDirectoryStart$_cont$betweenDirectoryEndAndRefStart$_refs$("    " * 3)$afterRefEnd"
    } else {
        Write-Host "Can't update wix project, could not find every markers in content'"
    }
    
}

# build the MSIs
MSBuild.exe DaisyConverter.sln /t:Installer\DaisyAddinForWordSetup /p:Configuration="Release" /p:Platform="x86";
MSBuild.exe DaisyConverter.sln /t:Installer\DaisyAddinForWordSetup /p:Configuration="Release" /p:Platform="x64";
# build the installer
MSBuild.exe DaisyConverter.sln /t:Installer\SaveAsDAISYInstaller /p:Configuration="Release" /p:Platform="Any CPU" /p:DefineConstants="UNIFIED";
