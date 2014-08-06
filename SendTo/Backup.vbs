Option Explicit

Call Main()

'* MAIN ======================================================================

Sub Main

    Dim arrPaths, path

    arrPaths = GetFilePaths

    For each path in arrPaths

        MsgBox path, vbInformation, GetDateStamp(path)

    Next

End Sub

'* ROUTINES ==================================================================

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

Function GetDateStamp(path)
    Dim fso, f, dtm, dateStamp
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path) Then
        Set f = fso.GetFile(path)
        dtm = f.DateLastModified
    End If

    dateStamp = Year(dtm) & _
      LZ(Month(dtm)) _
    & LZ(Day(dtm)) _
    & "_" _
    & LZ(Hour(dtm)) _
    & LZ(Minute(dtm)) _
    & LZ(Second(dtm))

    GetDateStamp = dateStamp

End Function

Function LZ(ByVal Number)
  If Number < 10 Then
    LZ = "0" & CStr(Number)
  Else
    LZ = CStr(Number)
  End If
End Function

