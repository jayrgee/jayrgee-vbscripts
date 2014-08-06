Option Explicit

Call Main()

'* MAIN ======================================================================

Sub Main

    Dim arrPaths, path

    arrPaths = GetFilePaths

    If UBound(arrPaths) < 0 Then

        MsgBox "Usage:" & vbcrlf & vbcrlf & Wscript.ScriptName & " Filename", vbInformation, Wscript.ScriptName

    Else

        For each path in arrPaths

            BackupFile path

        Next

    End If

End Sub

'* ROUTINES ==================================================================

Sub BackupFile(pathSrc)

    Const ReadOnly = 1

    Dim fso, f, dtm
    Dim nameDestFile, pathDestFolder, pathDest, extn

    Set fso = WScript.CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(pathSrc) Then
        Set f = fso.GetFile(pathSrc)
        dtm = f.DateLastModified
    Else
        ' file does not exist!
        Exit Sub

    End If

    pathDestFolder = fso.GetParentFolderName(pathSrc) & "\" & fso.GetBaseName(WScript.ScriptName)

    If Not fso.FolderExists(pathDestFolder) Then
        fso.CreateFolder pathDestFolder
    End If

    If Not fso.FolderExists(pathDestFolder) Then
        ' backup folder could not be created!
        Exit Sub

    End If

    extn = fso.GetExtensionName(pathSrc)
    If Len(extn) > 0 Then extn = "."  & extn

    nameDestFile = fso.GetBaseName(pathSrc) & "_" & GetDateStamp(dtm) & extn
    pathDest = pathDestFolder & "\" & nameDestFile

    If fso.FileExists(pathDest) Then
        MsgBox "file already exists: " & pathDest
        Exit Sub
    Else
        fso.CopyFile pathSrc, pathDest, false

        Set f = fso.GetFile(pathDest)
        If Not (f.attributes And ReadOnly) Then
            ' set read-only file attribute
            f.attributes = f.attributes + ReadOnly
        End If
    End If

End Sub

'-----------------------------------------------------------------------------

Sub DisplayError(msg)

    MsgBox msg, vbExclamation, Wscript.ScriptName

End Sub

'-----------------------------------------------------------------------------

Function GetFilePaths

    Const cLIST_DELIM = "|"

    Dim objArgs, arg
    Dim strPaths

    Set objArgs = WScript.Arguments

    strPaths = ""
    If objArgs.Count > 0 Then

        For Each arg in objArgs
            strPaths = strPaths & CStr(arg) & cLIST_DELIM
        Next

        If Len(strPaths) > 0 Then
            ' strip trailing delimiter
            strPaths = Left(strPaths, Len(strPaths) - 1)
        End If

        GetFilePaths = Split(strPaths, cLIST_DELIM, -1, vbTextCompare)
    Else
        GetFilePaths = Array()
    End if

End Function

'-----------------------------------------------------------------------------

Function GetDateStamp(dtm)

    On Error Resume Next

    GetDateStamp = CStr(Year(dtm)) & _
      LZ(Month(dtm)) _
    & LZ(Day(dtm)) _
    & "_" _
    & LZ(Hour(dtm)) _
    & LZ(Minute(dtm)) _
    & LZ(Second(dtm))

End Function

'-----------------------------------------------------------------------------

Function LZ(ByVal number)

    If number < 10 Then
        LZ = "0" & CStr(number)
    Else
        LZ = CStr(number)
    End If

End Function
