Option Explicit
' AQT_Utils.bas - Helpers: file read, JSON value extraction, unzip, VBProject import utilities.
Public Function AQT_ReadFile(path As String) As String
    On Error GoTo ErrHandler
    Dim fnum As Integer: fnum = FreeFile
    Open path For Binary As #fnum
    Dim txt As String
    If LOF(fnum) > 0 Then
        txt = Space(LOF(fnum))
        Get #fnum, , txt
    Else
        txt = ""
    End If
    Close #fnum
    AQT_ReadFile = txt
    Exit Function
ErrHandler:
    AQT_LogError "AQT_ReadFile error: " & Err.Description
    AQT_ReadFile = ""
End Function

Public Function AQT_JSONValue(json As String, key As String) As String
    On Error GoTo ErrHandler
    Dim pattern As String: pattern = " & key & " & ":"
    Dim p As Long: p = InStr(1, json, """" & key & """" & ":", vbTextCompare)
    If p = 0 Then AQT_JSONValue = "": Exit Function
    Dim subS As String: subS = Mid$(json, p + Len("""" & key & """" & ":""))
    Dim q1 As Long: q1 = InStr(1, subS, """" )
    If q1 = 0 Then
        Dim m As Long: m = InStr(1, subS, ",")
        If m = 0 Then m = InStr(1, subS, "}")
        If m = 0 Then AQT_JSONValue = Trim(subS) Else AQT_JSONValue = Trim(Left$(subS, m - 1))
        Exit Function
    End If
    Dim q2 As Long: q2 = InStr(q1 + 1, subS, """" )
    If q2 = 0 Then AQT_JSONValue = "" Else AQT_JSONValue = Mid$(subS, q1 + 1, q2 - q1 - 1)
    Exit Function
ErrHandler:
    AQT_LogError "AQT_JSONValue error: " & Err.Description
    AQT_JSONValue = ""
End Function

Public Sub AQT_Unzip(zipPath As String, outFolder As String)
    On Error GoTo ErrHandler
    Dim fso As Object, sh As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outFolder) Then fso.CreateFolder outFolder
    Set sh = CreateObject("Shell.Application")
    sh.Namespace(outFolder).CopyHere sh.Namespace(zipPath).Items, 16
    AQT_Log "Unzip requested: " & zipPath & " -> " & outFolder
    Exit Sub
ErrHandler:
    AQT_LogError "AQT_Unzip error: " & Err.Description
End Sub

Public Sub AQT_ImportModules(folderPath As String)
    On Error GoTo ErrHandler
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fld As Object: Set fld = fso.GetFolder(folderPath)
    Dim file As Object
    For Each file In fld.Files
        If LCase(Right(file.Name, 4)) = ".bas" Then AQT_ImportModule file.Path
    Next file
    Exit Sub
ErrHandler:
    AQT_LogError "AQT_ImportModules error: " & Err.Description
End Sub

Public Sub AQT_ImportModule(modulePath As String)
    On Error GoTo ErrHandler
    Dim vbProj As Object, vbComp As Object
    Set vbProj = ThisWorkbook.VBProject
    Dim moduleName As String: moduleName = Replace(Dir(modulePath), ".bas", "")
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
    On Error GoTo ErrHandler
    vbProj.VBComponents.Import modulePath
    AQT_Log "Imported module: " & moduleName
    Exit Sub
ErrHandler:
    AQT_LogError "AQT_ImportModule error: " & Err.Description & " (path:" & modulePath & ")"
End Sub
