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
    dim conditionsPath
    if WScript.Arguments.Count <> 2 then
		WScript.Echo "usage : CorrectConditions.vbs MsiFile ConditionsFile"
		exit sub
	end if
	
	databasePath = WScript.Arguments(0)
    WScript.Echo "Database Path = " & databasePath

	conditionsPath = WScript.Arguments(1)
    WScript.Echo "Conditions Path = " & conditionsPath
    
    
    dim installer
    set installer = Wscript.CreateObject("WindowsInstaller.Installer")

    dim openMode
    openMode = msiOpenDatabaseModeTransact

    dim database
    set database = installer.OpenDatabase(databasePath, openMode)

    dim conditions
    set conditions = WScript.CreateObject("Msxml2.DOMDocument.5.0")
    conditions.Async = false
    conditions.Load conditionsPath
    
    dim selectQuery
    dim view
    dim record
    dim continue
    
    dim filenames
    
    ' declare the custom action
    selectQuery = "SELECT Component_,FileName,Version FROM File "
    'selectQuery = "SELECT Action,Sequence FROM InstallUISequence"
    set view = database.OpenView(selectQuery)
    
    view.Execute
    
    continue = true
    while continue
        set record = view.Fetch
        if record is nothing then
            continue = false
        else
            filenames = Split(record.StringData(2), "|")
            'WScript.Echo record.StringData(1) & "," & filenames(1) & "," & record.StringData(3)
            CheckFile database, conditions, UCase(filenames(1)), record.StringData(3), record.StringData(1)
        end if
    wend
    
    Database.Commit
end sub

sub CheckFile(database, conditions, filename, fileversion, componentguid)
    if fileversion <> "" then
        dim xsl, node
        '   <file name= "MICROSOFT.VBE.INTEROP.DLL" version="12.0">NEVER</file>
        'xsl = "//file[@name=""" & filename & """ and starts-with(""" & fileversion & """, @version)]"
        xsl = "//file[@name=""" & filename & """ and starts-with(""" & fileversion & """, @version)]"
        'WScript.Echo xsl
        set node = conditions.SelectSingleNode(xsl)
        if not(node is nothing) then
            WScript.Echo "Fichier = " & filename
            WScript.Echo "Version = " & fileversion
            WScript.Echo "Condition = " & node.Text
            
            dim query2, view2
            query2 = "SELECT Component, Condition FROM Component WHERE Component='" & componentguid & "'"
            
            set view2 = database.OpenView(query2)
            view2.Execute
            
            dim record2
            set record2 = view2.Fetch
            if record2 is nothing then
                WScript.Echo "Component not found"
            else
                WScript.Echo "Component found"
                
                record2.StringData(2) = node.Text
                view2.Modify msiViewModifyUpdate, record2
            end if
        end if
    end if
end sub
