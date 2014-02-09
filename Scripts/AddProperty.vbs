option explicit

Const msiOpenDatabaseModeReadOnly     = 0
Const msiOpenDatabaseModeTransact     = 1
Const msiOpenDatabaseModeCreate       = 3

Const msiViewModifyInsert         = 1
Const msiViewModifyUpdate         = 2
Const msiViewModifyAssign         = 3
Const msiViewModifyReplace        = 4
Const msiViewModifyDelete         = 6

Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
Const TristateTrue = -1

main

sub main()
    Dim WshShell
    Set WshShell = WScript.CreateObject("WScript.Shell")

    dim databasePath
    
    if WScript.Arguments.Count <> 1 then
		WScript.Echo "usage : AddProperty.vbs MsiFile"
		exit sub
	end if
	
	databasePath = WScript.Arguments(0)
	
	WScript.Echo "AddProperty : Target MSI = " + databasePath

    dim installer
    set installer = Wscript.CreateObject("WindowsInstaller.Installer")

    dim openMode
    openMode = msiOpenDatabaseModeTransact

    dim database
    set database = installer.OpenDatabase(databasePath, openMode)

    AddProperty installer, database, "FolderForm_PrevArgs", "Custom3Buttons"
    AddProperty installer, database, "DISABLEADVTSHORTCUTS", "1"  
    database.Commit
    
    WScript.Echo "AddProperty terminsé"
    
end sub

sub AddProperty(installer, database, propertyName, propertyValue)
    dim selectQuery
    dim view
    dim record

	' find the form just before "FolderForm"
    selectQuery = "SELECT Property, Value FROM Property"
    
    set view = database.OpenView(selectQuery)
	set record = installer.CreateRecord(2)
	record.StringData(1) = propertyName
	record.StringData(2) = propertyValue
	view.Execute record
	view.Modify msiViewModifyAssign, record

end sub
