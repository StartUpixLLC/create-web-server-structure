Set FSO = CreateObject("Scripting.FileSystemObject")
Set F = FSO.GetFile(Wscript.ScriptFullName)
CurrentDir = FSO.GetParentFolderName(F) & "\"

CreateFolder CurrentDir & "/bin" & "/apache", 3
CreateFolder CurrentDir & "/bin" & "/mysql", 3
CreateFolder CurrentDir & "/bin" & "/php", 3

CreateFolder CurrentDir & "/data" & "/db" & "/data", 3
CreateFolder CurrentDir & "/data" & "/pma", 3
CreateFolder CurrentDir & "/data" & "/project001", 3
CreateFolder CurrentDir & "/data" & "/project002", 3

' Создание директории, если не существует
Sub CreateFolder(FolderSpec, MaxFoldersCount)
    For i = 0 to MaxFoldersCount
        SmartCreateFolder FolderSpec
    Next
End Sub

Sub SmartCreateFolder(strFolder)
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(strFolder) then
            SmartCreateFolder(.getparentfoldername(strFolder))
            .CreateFolder(strFolder)
        End If
    End With
End Sub
