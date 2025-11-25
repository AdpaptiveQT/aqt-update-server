Option Explicit
Sub AQT_ShowMenu()
    Dim msg As String
    msg = "AQT Menu:" & vbCrLf & _
          "1) Log Trade (AQT_LogTrade)" & vbCrLf & _
          "2) Recalculate XP (AQT_RecalculateAllXP)" & vbCrLf & _
          "3) Update System (AQT_DownloadAndInstallUpdate)" & vbCrLf & _
          "4) Activate License (AQT_ActivateSoftware)"
    MsgBox msg, vbInformation, "AQT Menu"
End Sub

Sub AQT_OpenSettings()
    MsgBox "Open AQT Settings - placeholder.", vbInformation
End Sub
