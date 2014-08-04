'*  ==========================================================================
'*  Script name: RaiseErr.vbs
'*  Created on:  2014-08-04
'*  Author:      John Gantner
'*  Purpose:     Demonstrate use of On Error and Err.Raise
'*               with custom errors
'*  --------------------------------------------------------------------------
Option Explicit

'*  Global custom errors
Const ERR_BAAH    = 10000
Const ERR_OH_NO   = 10001
Const ERR_OUCH    = 10002

'* call main sub
Call Main

'*  ==========================================================================
'*  Subroutine  : Main
'*  Description : Main process
'*  --------------------------------------------------------------------------
Sub Main()
    ' Any errors handled by next statement
    On Error Resume Next

    Const cProcess = "Main"

    RaiseACriticalError
    If Err.Number <> 0 Then DisplayClearError

    Ouch
    If Err.Number <> 0 Then DisplayClearError

    DoSomethingStupid
    If Err.Number <> 0 Then DisplayError : Err.Clear

    RaiseErr ERR_OH_NO, cProcess
    If VBError(Err.Number) = ERR_OH_NO Then
        DisplayClearError
        Exit Sub
    End If

End Sub

Sub RaiseACriticalError()
    ' Any errors handled by calling process
    On Error Goto 0

    Const cProcess = "RaiseACriticalError"

    RaiseErr ERR_BAAH, cProcess

End Sub

Sub Ouch()
    ' Any errors handled by calling process
    On Error Goto 0

    Const cProcess = "Ouch"

    RaiseErr ERR_OUCH, cProcess

End Sub

Sub DoSomethingStupid()
    ' Any errors handled by calling process
    On Error Goto 0

    Const cProcess = "DoSomethingStupid"

    Dim divByZero

    divByZero = 1 / 0

End Sub

'*  ==========================================================================
'*  Function    : VBError
'*  Description : Returns custom error number
'*                http://vb.mvps.org/hardcore/html/howtoraiseerrors.htm
'*  --------------------------------------------------------------------------
Function VBError(ByVal e)
    VBError = CLng(e) And &HFFFF&
End Function

'*  ==========================================================================
'*  Function    : COMError
'*  Description : Returns COM error number
'*                http://vb.mvps.org/hardcore/html/howtoraiseerrors.htm
'*  --------------------------------------------------------------------------
Function COMError(ByVal e)
    COMError = CLng(e) Or vbObjectError
End Function

'*  ==========================================================================
'*  Subroutine  : RaiseErr
'*  Description : Raise a custom error
'*  --------------------------------------------------------------------------
Sub RaiseErr(ByVal e, ByVal strSource)

    Dim strMsg

    Select Case e
    Case ERR_BAAH  strMsg = "baahh"
    Case ERR_OH_NO strMsg = "oh no"
    Case ERR_OUCH  strMsg = "ouch!"
    Case Else      strMsg = "x"
    End Select

    Err.Raise COMError(e), strSource, strMsg

End Sub

'*  ==========================================================================
'*  Subroutine  : DisplayError
'*  Description : Display an error
'*  --------------------------------------------------------------------------
Sub DisplayError()

    Dim arrMsg

    If Err.Number = 0 Then Exit Sub

    arrMsg = Array( _
        "Error: " & CStr(Err.Number), _
        "Description: " & Err.Description, _
        "Source: " & Err.Source)

    MsgBox Join(arrMsg, vbCrLf), vbExclamation, WScript.ScriptName

End Sub

'*  ==========================================================================
'*  Subroutine  : DisplayClearError
'*  Description : Display an error before clearing
'*  --------------------------------------------------------------------------
Sub DisplayClearError()
    DisplayError
    Err.Clear
End Sub

