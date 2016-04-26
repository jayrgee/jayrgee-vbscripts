Option Explicit

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim scriptName: scriptName = fso.GetBaseName(WScript.ScriptName)

Call Main()

'* MAIN ======================================================================

Sub Main

    Dim path

    If WScript.Arguments.Count < 1 Then

        WScript.Echo "Usage:" & vbcrlf & "cscript //nologo " & WScript.ScriptName & " root\file\path"

    Else

        IterateFromRoot WScript.Arguments(0)

    End If

End Sub

'* ROUTINES ==================================================================

Sub IterateFromRoot(pathRoot)

    WScript.Echo "pathRoot", pathRoot

    If fso.FolderExists(pathRoot) Then
        Recurse fso.GetFolder(pathRoot)
    Else
        WScript.Echo "path does not exist", pathRoot
    End If

End Sub

'-----------------------------------------------------------------------------

Sub Recurse(objFolder)
    Dim objFile, objSubFolder
    Dim path

    For Each objFile In objFolder.Files
        path = objfile.Path
        
        Select Case LCase(fso.GetExtensionName(path))
            Case "doc", "dot"
                ProcessWordFile path

            Case Else
                ' do nothing
        End Select
    Next

    For Each objSubFolder In objFolder.SubFolders
        Recurse objSubFolder
    Next
End Sub

'-----------------------------------------------------------------------------

Sub ProcessWordFile(path)

    WScript.Echo path

End Sub
