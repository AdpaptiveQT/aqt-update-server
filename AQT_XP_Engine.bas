Option Explicit
Sub AQT_RecalculateAllXP()
    On Error GoTo ErrHandler
    Dim wsXP As Worksheet: Set wsXP = ThisWorkbook.Sheets("AQT XP & Gamification System")
    Dim lastRow As Long: lastRow = wsXP.Cells(wsXP.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim result As String: result = UCase(Trim(wsXP.Cells(r, "P").Value))
        If result = "WIN" Then wsXP.Cells(r, "D").Value = 100 ElseIf result = "LOSS" Then wsXP.Cells(r, "D").Value = -25 Else wsXP.Cells(r, "D").Value = 0
        Dim qs As Long: qs = CLng(wsXP.Cells(r, "C").Value)
        Select Case qs
            Case 5: wsXP.Cells(r, "E").Value = 50
            Case 4: wsXP.Cells(r, "E").Value = 25
            Case Is <= 2: wsXP.Cells(r, "E").Value = -25
            Case Else: wsXP.Cells(r, "E").Value = 0
        End Select
        If Not IsNumeric(wsXP.Cells(r, "F").Value) Then wsXP.Cells(r, "F").Value = 0
        wsXP.Cells(r, "G").Value = wsXP.Cells(r, "D").Value + wsXP.Cells(r, "E").Value + wsXP.Cells(r, "F").Value
        If r = 2 Then wsXP.Cells(r, "H").Value = wsXP.Cells(r, "G").Value Else wsXP.Cells(r, "H").Value = wsXP.Cells(r - 1, "H").Value + wsXP.Cells(r, "G").Value
        wsXP.Cells(r, "I").Value = AQT_GetRankText(CLng(wsXP.Cells(r, "H").Value))
    Next r
    MsgBox "XP recalculated for " & (lastRow - 1) & " trades.", vbInformation
    Exit Sub
ErrHandler:
    AQT_LogError "AQT_RecalculateAllXP error: " & Err.Description
    MsgBox "Error recalculating XP: " & Err.Description, vbCritical
End Sub

Function AQT_GetRankText(totalXP As Long) As String
    If totalXP < 1000 Then AQT_GetRankText = "Novice" _
    ElseIf totalXP < 2500 Then AQT_GetRankText = "Developing Trader" _
    ElseIf totalXP < 5000 Then AQT_GetRankText = "Consistent Operator" _
    ElseIf totalXP < 8000 Then AQT_GetRankText = "Institutional Mindset" _
    Else AQT_GetRankText = "Adaptive Quantum Trader"
End Function
