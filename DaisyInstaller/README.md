# DAISY Installer : packages bundle executable

This project build the "bootstrapper" executable of the installation process for the addin.
This bootstrapper first checks for installations of office before launching the bundled MSI

For now, the addin and this installer is made to be working with windows x64, mainly due to the embedded pipeline-lite that itself embed a x64 jre. 
(this may be changed if requested by enough user, but be aware that Microsoft has dropped OEM distribution of Windows 10 32bits except for ARM speci).

## Unified installer

The actual unified installer bundles both the full x86 and x64 MSI packages (including for each a pipeline-lite with x64 jre).
To build this installer, you first need to build the "DaisyAddinForWordSetup" in both versions (x86 and x64).
Ensure that both packages are then declared in the resource manifest of the project `Properties/Resources.resx`.

Build the DaisyInstaller project in x86 configuration with the `UNIFIED` conditional build symbol defined in your project to obtain the unified installer

## Office specific installer

For office specific installer (with a lower bundle size), you will need to edit the resource manifest of the project `Properties/Resources.resx` 
(or the `Properties/Resources` tab) and change the conditionnal build symbols

For Office 32bits only installer : 
- remove all conditional build symbol in the DaisyInstaller `Properties/Build` tab;
- remove the `DaisyAddinForWordSetup_x64` file and keep the `DaisyAddinForWordSetup_x86` one in the `Properties/Resources` tab 

For Office 64bits only installer : 
- Add the `X64INSTALLER` value to the conditional build symbols in the DaisyInstaller `Properties/Build` tab;
- remove the DaisyAddinForWordSetup_x86 file and keep the `DaisyAddinForWordSetup_x64` one in the `Properties/Resources` tab

## Compilation constants

Please set the following constantes
`X64INSTALLER` : create the installer for Office 64 bits only
`UNIFIED` : create the installer for both version on office

# TODO List

- Move the installation directory selection from the MSI Packages to the bootstraper project
- Remove the pipeline from MSI packages et unzip a separate embedded pipeline-lite in the installation directory after the package installation is successfull
- If possible, set the pipeline lite launcher to search for either x86 or x64 JRE, remove the JRE from it, and deploy the JRE using the MSI packages
- if not possible, also export the launchers with specific settings and deploy the launchers with the JRE in the MSI packages