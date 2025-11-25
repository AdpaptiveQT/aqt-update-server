Option Explicit
' AQT_Debug_Logger.bas
' Lightweight logging utilities for Immediate Window and optional log worksheet.

Public Sub AQT_Log(msg As String)
    On Error Resume Next
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " | INFO  | " & msg
    If SheetExists("AQT_Log") Then
        With ThisWorkbook.Sheets("AQT_Log")
            .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Now
            .Cells(.Rows.Count, 1).End(xlUp).Offset(0, 1).Value = msg
        End With
    End If
End Sub

Public Sub AQT_LogError(msg As String)
    On Error Resume Next
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " | ERROR | " & msg
    If SheetExists("AQT_Log") Then
        With ThisWorkbook.Sheets("AQT_Log")
            .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Now
            .Cells(.Rows.Count, 1).End(xlUp).Offset(0, 1).Value = "ERROR: " & msg
        End With
    End If
End Sub

Public Sub AQT_LogFatal(msg As String)
    On Error Resume Next
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " | FATAL | " & msg
    MsgBox msg, vbCritical, "AQT Fatal Error"
    AQT_LogError msg
End Sub

Private Function SheetExists(shtName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shtName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
