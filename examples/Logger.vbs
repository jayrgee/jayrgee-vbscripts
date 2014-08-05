'*  ==========================================================================
'*  Script name : Logger.vbs
'*  Created on  : 2014-08-05
'*  Author      : John Gantner
'*  Purpose     : Implementation of reusable logger class
'*  --------------------------------------------------------------------------
Option Explicit

'*  Globals
Dim oLog

'* call main sub
Call Main

'*  ==========================================================================
'*  Subroutine  : Main
'*  Description : Main process
'*  --------------------------------------------------------------------------
Sub Main()
    ' Any errors handled by next statement
    'On Error Resume Next

    Const cProcess = "Main"

    Set oLog = New Logger

    oLog.WriteLog "blah"

End Sub


'*  ==========================================================================
'*  Class       : Logger
'*  Description : It logs. with a timestamp.
'*  --------------------------------------------------------------------------
Class Logger

    Private fso
    Private logPath
    Private logFile

    Public Sub WriteLog(ByVal sMsg)

        logFile.WriteLine(GetTimeStamp & " : " & sMsg)

    End Sub

    Private Sub Class_Initialize
        'WScript.Echo "Class_Initialize: Logger"

        Set fso = WScript.CreateObject("Scripting.FileSystemObject")

        logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\" _
            & fso.GetBaseName(WScript.ScriptFullName) & "_" _
            & GetTimeStamp() & ".log"

        'WScript.Echo "logPath: " & logPath

        Set logFile = fso.CreateTextFile(logPath, True)

    End Sub

    Private Sub Class_Terminate
        'WScript.Echo "Class_Terminate: Logger"

        logFile.Close

    End Sub

    Private Function GetTimeStamp()
        Dim dtm
        dtm = Now()

        GetTimeStamp = Year(dtm) & _
          LZ(Month(dtm)) _
        & LZ(Day(dtm)) _
        & LZ(Hour(dtm)) _
        & LZ(Minute(dtm)) _
        & LZ(Second(dtm))

    End Function

    Private Function LZ(ByVal number)
        If number < 10 Then
            LZ = "0" & CStr(number)
        Else
            LZ = CStr(number)
        End If
    End Function

End Class