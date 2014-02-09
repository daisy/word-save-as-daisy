mkdir DotNetFX30
move EnableDotNet3.exe DotNetFX30
mkdir PIA2007
move o2007pia.msi PIA2007
mkdir PIA2003
move o2003pia.msi PIA2003
mkdir KB908002
move extensibilityMSM.msi KB908002
move lockbackRegKey.msi KB908002
move office2003-kb907417sfxcab-ENU.exe KB908002
mkdir JAVARUNTIME
move jre-6u10-windows-i586-p-iftw.exe JAVARUNTIME
mkdir CompatibilityPack
move FileFormatConverters.exe CompatibilityPack

start /W setup.exe
