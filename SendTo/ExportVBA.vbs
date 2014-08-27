Option Explicit

Const cPROCESS = "ExportVBA"

Call Main()

'* MAIN ======================================================================

Sub Main

    Dim arrPaths, strPath

    arrPaths = GetFilePaths

    For each strPath in arrPaths

        If Not CheckVBAPath(strPath) Then
            MsgBox "Sorry, cannot export VBA from:" & vbCrLf & strPath, vbExclamation, cPROCESS
        End If

    Next

End Sub

'* ROUTINES ==================================================================

Function CheckVBAPath(path)

    Dim fso, extn
    Set fso = CreateObject("Scripting.FileSystemObject")

    CheckVBAPath = True

    'Check file Exists
    If Not fso.FileExists(path) Then
        CheckVBAPath = False
        Exit Function
    End If

    'Check file Extension
    extn = LCase(fso.GetExtensionName(path))
    Select Case extn
        Case "doc", "docm"
            'ExportVBAFromDoc path
        Case "xls", "xlsm"
            ExportVBAFromXLS path
        Case Else
            CheckVBAPath = False
    End Select

End Function

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


Function CreateNewFolder(path)
    Const cPROCESS = "CreateNewFolder"

    Dim fso

    On Error Resume Next

    CreateNewFolder = False

    Set fso = WScript.CreateObject("Scripting.FileSystemObject")

    If Err.Number <> 0 Then Exit Function

    If fso.FolderExists(path) Then
        Err.Raise vbObjectError + 1050, cProcess, "Folder already exists:" & vbCrLf & path
        Exit Function
    End If

    If Err.Number = 0 Then
        fso.CreateFolder(path)
        If Err.Number = 0 Then CreateNewFolder = True
    End If

End Function


Sub CopyFileToFolder(pathFile, pathFolder)

    Dim fso
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(pathFile) And fso.FolderExists(pathFolder) Then

        fso.CopyFile pathFile, pathFolder & "\"

    End if

End Sub

Function ExportVBComponents(ByRef colVbc, ByVal pathRoot)

    On Error Resume Next

    Dim fso
    Dim oVBC   'As VBIDE.VBComponent
    Dim iX     'As Integer

    iX = 0
    For Each oVBC In colVBC
        oVBC.Export pathRoot & "\" & oVBC.Name & GetVBExtension(oVBC.Type)
        iX = iX + 1
    Next 'oVBC
    
    ExportVBComponents = iX

End Function

Sub ExportVBAFromXLS(ByVal strPathFile)
    Dim oApp   'As Excel.Application
    Dim oWbk   'As Excel.Workbook
    Dim oVBP   'As VBIDE.VBProject
    Dim colVBC 'As VBIDE.VBComponents
    Dim oVBC   'As VBIDE.VBComponent
    Dim iCount     'As Integer
    Dim dateStamp
    Dim strExportRoot 'As String

    dateStamp = GetDateStamp(strPathFile)

    Set oApp = CreateObject("Excel.Application")
    ' disable events in case of any sneaky
    oApp.EnableEvents = False

    Set oWbk = oApp.Workbooks.Open(strPathFile, , True)

    strExportRoot = oWbk.Path & "\" & Replace(oWbk.Name, ".", "_") & "_" & dateStamp

    Set colVBC = oWbk.VBProject.VBComponents

    If colVBC.Count > 0 Then
        If CreateNewFolder(strExportRoot) Then
            iCount = ExportVBComponents(colVBC, strExportRoot)
        Else
            MsgBox "Could not create sub-folder: " & vbCrLf & strExportRoot, vbCritical, cPROCESS
        End If
    Else
        MsgBox "No VBA to export from this file: " & vbCrLf & strPathFile, vbExclamation, cPROCESS
    End If

    Set colVBC = Nothing
    
    oWbk.Close
    Set oWbk = Nothing

    oApp.Quit
    Set oApp = Nothing

    If iX > 0 Then
        CopyFileToFolder strPathFile, strExportRoot
        MsgBox CStr(iCount + 1) & " file(s) exported to " & strExportRoot, vbInformation, cPROCESS
    End If

End Sub

Sub ExportVBAFromDoc(ByVal strPathFile)
    Dim oApp   'As Word.Application
    Dim oDoc   'As Word.Document
    Dim oVBP   'As VBIDE.VBProject
    Dim colVBC 'As VBIDE.VBComponents
    Dim oVBC   'As VBIDE.VBComponent
    Dim iX     'As Integer

    Set oApp = CreateObject("Word.Application")

    Set oDoc = oApp.Documents.Open(strPathFile, , True, , , , , , , , , False)
    'MsgBox oDoc.Path

    On Error Resume Next

    Dim strExportRoot 'As String
    strExportRoot = oDoc.Path & "\" & Replace(oDoc.Name, ".", "_") & "_" & LongStamp()
    CreateNewFolder strExportRoot
    On Error GoTo 0

    Set colVBC = oDoc.VBProject.VBComponents

    iX = 0
    For Each oVBC In colVBC

        oVBC.Export strExportRoot & "\" & oVBC.Name & GetVBExtension(oVBC.Type)
        iX = iX + 1

    Next 'oVBC

    Set colVBC = Nothing

    oDoc.Close False

    Set oDoc = Nothing

    oApp.Quit False

    Set oApp = Nothing

    If iX > 0 Then CopyFileToFolder strPathFile, strExportRoot

    MsgBox CStr(iX) & " file(s) exported to " & strExportRoot, vbInformation, cPROCESS

End Sub

Function GetVBExtension(ByVal vbType)

    Const vbext_ct_ActiveXDesigner = 11
    Const vbext_ct_ClassModule = 2
    Const vbext_ct_Document = 100
    Const vbext_ct_MSForm = 3
    Const vbext_ct_StdModule = 1

    Dim strResult 'As String

    Select Case vbType
    Case vbext_ct_ActiveXDesigner
        strResult = ".dsr"
    Case vbext_ct_ClassModule
        strResult = ".cls"
    Case vbext_ct_Document
        strResult = ".cls"
    Case vbext_ct_MSForm
        strResult = ".frm"
    Case vbext_ct_StdModule
        strResult = ".bas"
    Case Else
        strResult = ".txt"
    End Select

    GetVBExtension = strResult

End Function

Function LZ(ByVal Number)
  If Number < 10 Then
    LZ = "0" & CStr(Number)
  Else
    LZ = CStr(Number)
  End If
End Function

Function LongStamp()
  Dim CurrTime
  CurrTime = Now()

  LongStamp = Year(CurrTime) & _
      LZ(Month(CurrTime)) _
    & LZ(Day(CurrTime)) _
    & LZ(Hour(CurrTime)) _
    & LZ(Minute(CurrTime)) _
    & LZ(Second(CurrTime))
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

Sub ReportError()

    Dim errType
    Dim arrMsg
    
    Select Case (Err.Number - vbObjectError)
    Case 1050 ErrType = vbCritical
    Case Else ErrType = vbExclamation
    End Select
    
    arrMsg = Array( _
        "Error: " & CStr(Err.Number), _
        "Description: " & Err.Description, _
        "Source: " & Err.Source)

    MsgBox Join(arrMsg, vbCrLf), errType, WScript.ScriptName
    Err.Clear

End Sub