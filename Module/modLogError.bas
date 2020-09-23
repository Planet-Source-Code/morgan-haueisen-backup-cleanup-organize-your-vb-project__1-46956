Attribute VB_Name = "modLogError"
Option Explicit

Public Sub Err_Handler(Optional ByVal vblnDisplayError As Boolean = True, _
                       Optional ByVal vstrErrNumber As String = vbNullString, _
                       Optional ByVal vstrErrDescription As String = vbNullString, _
                       Optional ByVal vstrModuleName As String = vbNullString, _
                       Optional ByVal vstrProcName As String = vbNullString)

  Dim strTemp     As String
  Dim lngFN       As Long
  Dim frmErrorMsg As New frmMsgBox

    '/* Purpose: Error handling - On Error

    '/* Show Error Message
    If vblnDisplayError Then
        strTemp = "Error occured: "
        If Len(vstrErrNumber) > 0 Then strTemp = strTemp & vstrErrNumber & vbNewLine Else strTemp = strTemp & vbNewLine
        If Len(vstrErrDescription) > 0 Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
        If Len(vstrModuleName) > 0 Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
        If Len(vstrProcName) > 0 Then strTemp = strTemp & "Function: " & vstrProcName
        frmErrorMsg.SMessageModal strTemp, vbCritical, App.Title & " - ERROR"
    End If

    '/* Write error log
    lngFN = FreeFile
    Open App.Path & "\ErrorLog.txt" For Append As #lngFN
    Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, Environ$("username"), Environ$("computername")
    Close #lngFN

End Sub


