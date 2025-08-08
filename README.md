# "Save as DAISY" Add-In for Microsoft Word

The "Save as DAISY" Add-In for Microsoft Word is a collaborative standards-based development project initiated in 2007 by the DAISY Consortium and Microsoft.
It was initially developed by [Sonata Software](http://www.sonata-software.com/) (versions 1.x), then further developed by [Intergen](http://intergen.co.nz/) (up to version 2.5.5.1).

The code available in this GitHub project has been initially copied from its original location on SourceForge as the [openxml-daisy project](http://sourceforge.net/projects/openxml-daisy/).

## Building requirement

The project is being updated to build with visual studio 2019, using the dotnet4 plugin and the wix toolser in its version 3.11.

## Notes

In case of multiple crash of the addin, it might be registered in Word's addin blocklist.
`HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItem`

