Option Explicit
Private Const SETUP_SHEET As String = "AQT Setup Tracker"
Private Const XP_SHEET As String = "AQT XP & Gamification System"
Sub AQT_LogTrade()
    On Error GoTo ErrHandler
    Dim wsSetup As Worksheet: Set wsSetup = ThisWorkbook.Sheets(SETUP_SHEET)
    Dim wsXP As Worksheet: Set wsXP = ThisWorkbook.Sheets(XP_SHEET)
    Dim lastRowSetup As Long: lastRowSetup = wsSetup.Cells(wsSetup.Rows.Count, "A").End(xlUp).Row
    Dim nextRowSetup As Long: nextRowSetup = lastRowSetup + 1
    If wsSetup.Cells(nextRowSetup, "A").Value = "" Then MsgBox "No trade data to log.", vbExclamation: Exit Sub
    wsSetup.Cells(lastRowSetup, "K").AutoFill Destination:=wsSetup.Range(wsSetup.Cells(lastRowSetup, "K"), wsSetup.Cells(nextRowSetup, "K"))
    wsSetup.Cells(lastRowSetup, "I").AutoFill Destination:=wsSetup.Range(wsSetup.Cells(lastRowSetup, "I"), wsSetup.Cells(nextRowSetup, "I"))
    wsSetup.Cells(lastRowSetup, "P").AutoFill Destination:=wsSetup.Range(wsSetup.Cells(lastRowSetup, "P"), wsSetup.Cells(nextRowSetup, "P"))
    Dim lastRowXP As Long: lastRowXP = wsXP.Cells(wsXP.Rows.Count, "A").End(xlUp).Row
    Dim nextRowXP As Long: nextRowXP = lastRowXP + 1
    wsXP.Cells(nextRowXP, "A").Value = wsSetup.Cells(nextRowSetup, "A").Value
    wsXP.Cells(nextRowXP, "B").Value = nextRowSetup - 1
    wsXP.Cells(nextRowXP, "C").Value = wsSetup.Cells(nextRowSetup, "K").Value
    wsXP.Range(wsXP.Cells(lastRowXP, "D"), wsXP.Cells(lastRowXP, "J")).AutoFill Destination:=wsXP.Range(wsXP.Cells(lastRowXP, "D"), wsXP.Cells(nextRowXP, "J"))
    Application.Calculate
    AQT_Log "Trade logged to XP sheet."
    Exit Sub
ErrHandler:
    AQT_LogError "AQT_LogTrade error: " & Err.Description
    MsgBox "Error logging trade: " & Err.Description, vbCritical
End Sub
