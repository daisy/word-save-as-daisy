# Script for building a minimal java runtime (jre) that can run a list of jar files
# usage : powershell -ExecutionPolicy Unrestricted "./build_runtime.ps1" [jar1.jar jar2.jar] -o output/jre/dir
# An installed JDK 9 or higher is required,
# the java version currently available is the path is used to build the runtime

$modules=[System.Collections.ArrayList]@();
$outputdir = "jre";

for ( $i = 0; $i -lt $args.count; $i++ ) {
    if ($args[ $i ] -eq "-o"){ $outputdir=$args[ $i+1 ]; $i = $i + 1;}
    else {
        foreach ($line in (jdeps --list-deps $args[$i])){
            $null = $modules.Add($line.Trim());
        }
        
    }
}

# the default module java.base is required if no jars are provided
if($modules.Length -eq 0){
    $modules.Add("java.base");
}

$addModulesValue = $modules -join ',';
Write-Host "Modules to be included in the runtime : $addModulesValue"

if(Test-Path "$outputdir"){
    Write-Host "Removing $outputdir for new runtime deployement"
    Remove-Item -Recurse -Force "$outputdir"
}

Write-Host "Deploying java runtime in $outputdir"

jlink -G --no-header-files --no-man-pages --compress=0 --add-modules $addModulesValue --output "$outputdir"
