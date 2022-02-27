'
'    ===========================================================================
'    Name:           PowerShell scripts launcher for E3.Series
'
'    Description:
'    -------------------------------------------------------------------------
'     Filename:      launcher.vbs
'     Created by:    Dmytro Chukhran
'     Support:       https://github.com/chukhran/E3.Series-powershell-scripts
'    -------------------------------------------------------------------------
'     Changes:
'        27/02/2022 Initial Version
'
'    ===========================================================================

Option Explicit


'
Dim FSO
Dim launcherFileName, launcherFolderFullPath
Set FSO = CreateObject("Scripting.FileSystemObject")
launcherFileName = FSO.GetFileName(Wscript.ScriptFullName)
launcherFolderFullPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder.Path


' Determine running method
'-------------------------
' There are two ways of scripts running in E3.Series. 
' These are internally and externally. 
' The launcher script must be run internally only.

' If it is internally running 
If(InStr(WScript.FullName, ".series")) Then
    Dim E3
    Set E3 = WScript
Else
    MsgBox "The launcher must be run from E3.Series only!" &_
    vbCrLf & vbCrLf & "Press OK for exit.", vbOKOnly + vbExclamation, _
    "PowerSell scripls launcher for E3.Series"
    Wscript.Quit
End If


' E3.PutInfo 0, "Running of " & launcherFileName & " launcher ..."
E3.PutInfo 0, "Running of PowerShell script launcher ..."


' Determine a file path of the PowerShell script
'-----------------------------------------------
' The filename of the PowerShell script 
' must be passed as a parameter of the launcher

Dim args, argsCnt
args = E3.ScriptArguments
argsCnt = UBound(args)-LBound(args)+1

If(argsCnt < 1) Then
    E3.PutError 1, _
    "The filename of the PowerShell script " &_ 
    "was not passed to the launcher. " & vbCrLf &_
    "Please verify the Arguments field " &_
    "in the Add-on tab of the Customize window."
    Wscript.Quit
End If


' Get a file path of the PowerShell script
Dim psFilePath
psFilePath = args(0)

' Get the full path of the PowerShell script
Dim psFileFullPath
psFileFullPath = FSO.GetAbsolutePathName(psFilePath)

' If a file not found
If Not FSO.FileExists(psFileFullPath) Then

    ' Trying to check the full path built from 
    ' a given relative file path and launcher folder
    psFileFullPath = FSO.BuildPath(launcherFolderFullPath, psFilePath)
    psFileFullPath = FSO.GetAbsolutePathName(psFileFullPath)
    
    If Not FSO.FileExists(psFileFullPath) Then
        E3.PutError 1, psFilePath & " file not found. " &_
        "Please verify the correct file name was given."
        Wscript.Quit
    End If
End If

' Checking file extension of the PowerShell script
Dim usedExt, corExt, checkExt
usedExt = FSO.GetExtensionName(psFileFullPath)
corExt = "ps1"
checkExt = StrComp(corExt, usedExt, vbTextCompare)

If checkExt <> 0 Then
    E3.PutError 1, _
    "A file extension of the PowerShell script " &_
    "must be '" & LCase(corExt) & "' however '" & _
    LCase(usedExt) & "' is used for a passed file. " &_
    "Please verify the correct file extension was given."
    Wscript.Quit
End If 

' Running of launcher is correct
'E3.PutInfo 0, "Ok"


' Running of the PowerShell script
E3.PutInfo 0, "Running of the '" &_
FSO.GetFileName(psFileFullPath) & "' script ..."

Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -sta -file " & Chr(34) & psFileFullPath & Chr(34), 0, True 
Set WshShell = Nothing

E3.PutInfo 0, "Exiting of the launcher ..."