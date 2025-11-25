Option Explicit
Public Const LICENSE_SERVER_URL As String = "https://your-license-server.example/api/validate"

Function AQT_IsLicenseValidLocal(key As String) As Boolean
    On Error GoTo ErrHandler
    If Len(Trim(key)) = 0 Then AQT_IsLicenseValidLocal = False: Exit Function
    If InStr(1, key, "AQT-", vbTextCompare) <> 1 Then AQT_IsLicenseValidLocal = False: Exit Function
    AQT_IsLicenseValidLocal = True
    Exit Function
ErrHandler:
    AQT_LogError "AQT_IsLicenseValidLocal error: " & Err.Description
    AQT_IsLicenseValidLocal = False
End Function

Function AQT_IsKeyValidatedRemotely(licenseKey As String) As Boolean
    On Error GoTo ErrHandler
    Dim httpReq As Object: Set httpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim payload As String
    payload = "{" & Chr(34) & "license_key" & Chr(34) & ":" & Chr(34) & licenseKey & Chr(34) & "}"
    AQT_Log "Remote validation POST to: " & LICENSE_SERVER_URL
    httpReq.Open "POST", LICENSE_SERVER_URL, False
    httpReq.SetRequestHeader "Content-Type", "application/json"
    httpReq.Send payload
    AQT_Log "Remote validation HTTP status: " & httpReq.Status
    If httpReq.Status = 200 Then
        AQT_IsKeyValidatedRemotely = (InStr(1, httpReq.responseText, """valid"": true", vbTextCompare) > 0)
    Else
        AQT_LogError "Remote validation returned status: " & httpReq.Status
        AQT_IsKeyValidatedRemotely = False
    End If
    Exit Function
ErrHandler:
    AQT_LogFatal "Network error validating license: " & Err.Description
    AQT_IsKeyValidatedRemotely = False
End Function

Sub AQT_ActivateSoftware()
    On Error GoTo ErrHandler
    Dim newKey As String
    newKey = InputBox("Please enter your AQT License Key:", "AQT License Activation")
    If newKey = "" Then Exit Sub
    AQT_Log "Activation started for key: " & Left(newKey, 20)
    If Not AQT_IsLicenseValidLocal(newKey) Then
        MsgBox "Invalid license key format.", vbCritical: Exit Sub
    End If
    If Not AQT_IsKeyValidatedRemotely(newKey) Then
        MsgBox "License key validation failed (server).", vbCritical: Exit Sub
    End If
    On Error Resume Next: ThisWorkbook.Names("AQT_LICENSE_KEY").Delete: On Error GoTo ErrHandler
    ThisWorkbook.Names.Add Name:="AQT_LICENSE_KEY", RefersTo:="=""" & newKey & """"
    AQT_Log "License activated and stored."
    AQT_DownloadAndInstallUpdate
    Exit Sub
ErrHandler:
    AQT_LogError "AQT_ActivateSoftware error: " & Err.Description
    MsgBox "Activation error: " & Err.Description, vbCritical
End Sub
