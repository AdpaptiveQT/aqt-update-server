Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As LongPtr
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If
Sub AQT_DownloadAndInstallUpdate()
    On Error GoTo ErrHandler
    AQT_Log "AQT_DownloadAndInstallUpdate started."
    Dim manifestUrl As String: manifestUrl = InputBox("Enter the version.json manifest URL:", "AQT Updater")
    If manifestUrl = "" Then Exit Sub
    Dim manifestLocal As String: manifestLocal = ThisWorkbook.Path & Application.PathSeparator & "AQT_manifest.json"
    If URLDownloadToFile(0, manifestUrl, manifestLocal, 0, 0) <> 0 Then AQT_LogFatal "Failed to download manifest"
    AQT_Log "Manifest downloaded."
    Dim json As String: json = AQT_ReadFile(manifestLocal)
    Dim downloadUrl As String: downloadUrl = AQT_JSONValue(json, "download_url")
    Dim expectedHash As String: expectedHash = AQT_JSONValue(json, "checksum_sha256")
    If downloadUrl = "" Then AQT_LogFatal "Manifest missing download_url"
    Dim localZip As String: localZip = ThisWorkbook.Path & Application.PathSeparator & "AQT_Payload.zip"
    If URLDownloadToFile(0, downloadUrl, localZip, 0, 0) <> 0 Then AQT_LogFatal "Failed to download payload"
    AQT_Log "Payload downloaded to: " & localZip
    Dim actualHash As String: actualHash = AQT_SHA256_File(localZip)
    AQT_Log "Calculated SHA256: " & actualHash
    If UCase(actualHash) <> UCase(expectedHash) Then AQT_LogFatal "SHA mismatch: expected " & expectedHash & " got " & actualHash
    AQT_Log "Integrity check passed."
    Dim unpackFolder As String: unpackFolder = ThisWorkbook.Path & Application.PathSeparator & "AQT_Unpacked"
    AQT_Unzip localZip, unpackFolder
    AQT_ImportModules unpackFolder
    AQT_Log "Update installation complete."
    MsgBox "AQT modules installed. Please save workbook as .xlsm to persist macros.", vbInformation
    Exit Sub
ErrHandler:
    AQT_LogError "AQT_DownloadAndInstallUpdate error: " & Err.Description
    MsgBox "Updater error: " & Err.Description, vbCritical
End Sub
