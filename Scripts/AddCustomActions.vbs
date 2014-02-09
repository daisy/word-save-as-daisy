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
    'WScript.Echo WshShell.CurrentDirectory

    
    
    dim databasePath
    dim dllPath
    if WScript.Arguments.Count <> 2 then
		WScript.Echo "usage : AddCustomAction.vbs MsiFile DllFile"
		exit sub
	end if
	
	databasePath = WScript.Arguments(0)
    	dllPath = WScript.Arguments(1)
	
	'WScript.Echo "Target MSI : " + databasePath
	'WScript.Echo "Source DLL : " + dllPath

    dim installer
    set installer = Wscript.CreateObject("WindowsInstaller.Installer")

    dim openMode
    openMode = msiOpenDatabaseModeTransact

    dim database
    set database = installer.OpenDatabase(databasePath, openMode)
    

    AjouterAction installer, database, "OdfInstallHelper", dllPath, "ForceUniqueInstall", 300
    AjouterAction installer, database, "OdfInstallHelper", "", "DetectPreviousConverters", 301
    AjouterAction installer, database, "OdfInstallHelper", "", "GetWordVersion", 302
    AjouterAction installer, database, "OdfInstallHelper", "", "LaunchReadme", 1400

    'WScript.Echo "Custom actions installed"
end sub


sub AjouterAction(installer, database, dllName, dllPath, actionName, sequence)
    dim selectQuery
    dim view
    dim record
    
    ' declare the custom action
    selectQuery = "SELECT Action,Type,Source,Target FROM CustomAction"
    set view = database.OpenView(selectQuery)
    set record = installer.CreateRecord(4)
    record.StringData(1) = actionName
    record.IntegerData(2) = 1
    record.StringData(3) = dllName
    record.StringData(4) = actionName
    view.Execute record
    view.Modify msiViewModifyAssign, record


    ' register it in the InstallUISequence
    selectQuery = "SELECT Action,Sequence FROM InstallUISequence"
    set view = database.OpenView(selectQuery)
    set record = installer.CreateRecord(2)
    record.StringData(1) = actionName
    record.IntegerData(2) = sequence
    view.Execute record
    view.Modify msiViewModifyAssign, record

    ' add the DLL file in the MSI
    if dllPath <> "" then
        selectQuery = "SELECT Name,Data FROM Binary"
        set view = database.OpenView(selectQuery)
        set record = installer.CreateRecord(2)
        record.StringData(1) = dllName
        record.SetStream 2, dllPath
        view.Execute record
        view.Modify msiViewModifyAssign, record
    end if

    ' save the changed
    database.Commit
end sub
