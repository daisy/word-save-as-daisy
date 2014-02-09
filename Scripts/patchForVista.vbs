dim installer, summary
set installer = Wscript.CreateObject("WindowsInstaller.Installer")
set summary = Installer.SummaryInformation(WScript.Arguments(0), 20)
summary.Property(15) = 10
summary.Persist