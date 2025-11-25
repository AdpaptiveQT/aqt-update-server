Option Explicit
' AQT_SHA256.bas - Compact SHA-256 implementation (suitable for integrity checks)

Function AQT_FileToBytes(path As String) As Byte()
    On Error GoTo ErrHandler
    Dim f As Integer: f = FreeFile
    Open path For Binary As #f
    Dim b() As Byte
    If LOF(f) > 0 Then
        ReDim b(0 To LOF(f) - 1)
        Get #f, , b
    Else
        ReDim b(0 To -1)
    End If
    Close #f
    AQT_FileToBytes = b
    Exit Function
ErrHandler:
    AQT_LogError "AQT_FileToBytes failed: " & Err.Description
    Dim empty() As Byte: ReDim empty(0 To -1)
    AQT_FileToBytes = empty
End Function

Function AQT_SHA256_File(path As String) As String
    Dim b() As Byte: b = AQT_FileToBytes(path)
    AQT_SHA256_File = AQT_SHA256_Bytes(b)
End Function

Private Function ROTR(x As Long, n As Long) As Long
    ROTR = ((x And &HFFFFFFFF) \ (2 ^ n)) Or ((x And &HFFFFFFFF) * (2 ^ (32 - n)) And &HFFFFFFFF)
End Function
Private Function SHR(x As Long, n As Long) As Long
    SHR = (x And &HFFFFFFFF) \ (2 ^ n)
End Function
Private Function Sigma0(x As Long) As Long
    Sigma0 = (ROTR(x, 2) Xor ROTR(x, 13) Xor ROTR(x, 22)) And &HFFFFFFFF
End Function
Private Function Sigma1(x As Long) As Long
    Sigma1 = (ROTR(x, 6) Xor ROTR(x, 11) Xor ROTR(x, 25)) And &HFFFFFFFF
End Function
Private Function sigma0(x As Long) As Long
    sigma0 = (ROTR(x, 7) Xor ROTR(x, 18) Xor SHR(x, 3)) And &HFFFFFFFF
End Function
Private Function sigma1(x As Long) As Long
    sigma1 = (ROTR(x, 17) Xor ROTR(x, 19) Xor SHR(x, 10)) And &HFFFFFFFF
End Function
Private Function Ch(x As Long, y As Long, z As Long) As Long
    Ch = ((x And y) Xor ((Not x) And z)) And &HFFFFFFFF
End Function
Private Function Maj(x As Long, y As Long, z As Long) As Long
    Maj = ((x And y) Xor (x And z) Xor (y And z)) And &HFFFFFFFF
End Function

Private Function SHA256_K() As Variant
    SHA256_K = Array(&H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DF, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEEA9215, &HC67178F2)
End Function

Function AQT_SHA256_Bytes(bytes() As Byte) As String
    On Error GoTo ErrHandler
    Dim K As Variant: K = SHA256_K()
    Dim H(0 To 7) As Long
    H(0) = &H6A09E667: H(1) = &HBB67AE85: H(2) = &H3C6EF372: H(3) = &HA54FF53A
    H(4) = &H510E527F: H(5) = &H9B05688C: H(6) = &H1F83D9AB: H(7) = &H5BE0CD19
    Dim ml As Long: ml = -1
    If Not (UBound(bytes) < 0) Then ml = UBound(bytes) + 1
    If ml = -1 Then AQT_SHA256_Bytes = "": Exit Function
    Dim totalBits As Double: totalBits = ml * 8#
    Dim padLen As Long: padLen = 64 - ((ml + 9) Mod 64)
    If padLen < 0 Then padLen = padLen + 64
    Dim paddedLen As Long: paddedLen = ml + 1 + padLen + 8
    Dim p() As Byte: ReDim p(0 To paddedLen - 1)
    Dim i As Long
    For i = 0 To ml - 1: p(i) = bytes(i): Next i
    p(ml) = &H80
    Dim highBits As Long: highBits = Int((totalBits) / (2 ^ 32))
    Dim lowBits As Long: lowBits = totalBits Mod (2 ^ 32)
    p(paddedLen - 8) = (highBits \ (2 ^ 24)) And &HFF
    p(paddedLen - 7) = (highBits \ (2 ^ 16)) And &HFF
    p(paddedLen - 6) = (highBits \ (2 ^ 8)) And &HFF
    p(paddedLen - 5) = highBits And &HFF
    p(paddedLen - 4) = (lowBits \ (2 ^ 24)) And &HFF
    p(paddedLen - 3) = (lowBits \ (2 ^ 16)) And &HFF
    p(paddedLen - 2) = (lowBits \ (2 ^ 8)) And &HFF
    p(paddedLen - 1) = lowBits And &HFF
    Dim w(0 To 63) As Long, a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As Long
    Dim off As Long, j As Long
    For off = 0 To paddedLen - 1 Step 64
        For j = 0 To 15
            w(j) = (CLng(p(off + j * 4)) * (2 ^ 24)) Or (CLng(p(off + j * 4 + 1)) * (2 ^ 16)) Or (CLng(p(off + j * 4 + 2)) * (2 ^ 8)) Or CLng(p(off + j * 4 + 3))
        Next j
        For j = 16 To 63
            w(j) = (sigma1(w(j - 2)) + w(j - 7) + sigma0(w(j - 15)) + w(j - 16)) And &HFFFFFFFF
        Next j
        a = H(0): b = H(1): c = H(2): d = H(3): e = H(4): f = H(5): g = H(6): h = H(7)
        For j = 0 To 63
            Dim T1 As Long, T2 As Long
            T1 = (h + Sigma1(e) + Ch(e, f, g) + CLng(K(j)) + w(j)) And &HFFFFFFFF
            T2 = (Sigma0(a) + Maj(a, b, c)) And &HFFFFFFFF
            h = g: g = f: f = e: e = (d + T1) And &HFFFFFFFF: d = c: c = b: b = a: a = (T1 + T2) And &HFFFFFFFF
        Next j
        H(0) = (H(0) + a) And &HFFFFFFFF: H(1) = (H(1) + b) And &HFFFFFFFF: H(2) = (H(2) + c) And &HFFFFFFFF: H(3) = (H(3) + d) And &HFFFFFFFF
        H(4) = (H(4) + e) And &HFFFFFFFF: H(5) = (H(5) + f) And &HFFFFFFFF: H(6) = (H(6) + g) And &HFFFFFFFF: H(7) = (H(7) + h) And &HFFFFFFFF
    Next off
    Dim hexOut As String: hexOut = ""
    For i = 0 To 7: hexOut = hexOut & Right$("00000000" & Hex$(H(i) And &HFFFFFFFF), 8): Next i
    AQT_SHA256_Bytes = LCase(hexOut)
    Exit Function
End Function

Private Function ROTR32(x As Long, n As Long) As Long: ROTR32 = ((x And &HFFFFFFFF) \ (2 ^ n)) Or ((x And &HFFFFFFFF) * (2 ^ (32 - n)) And &HFFFFFFFF): End Function
Private Function SHR32(x As Long, n As Long) As Long: SHR32 = (x And &HFFFFFFFF) \ (2 ^ n): End Function
Private Function Sigma1(x As Long) As Long: Sigma1 = (ROTR32(x, 6) Xor ROTR32(x, 11) Xor ROTR32(x, 25)) And &HFFFFFFFF: End Function
Private Function Sigma0(x As Long) As Long: Sigma0 = (ROTR32(x, 2) Xor ROTR32(x, 13) Xor ROTR32(x, 22)) And &HFFFFFFFF: End Function
Private Function sigma1(x As Long) As Long: sigma1 = (ROTR32(x, 17) Xor ROTR32(x, 19) Xor SHR32(x, 10)) And &HFFFFFFFF: End Function
Private Function sigma0(x As Long) As Long: sigma0 = (ROTR32(x, 7) Xor ROTR32(x, 18) Xor SHR32(x, 3)) And &HFFFFFFFF: End Function
Private Function Ch(x As Long, y As Long, z As Long) As Long: Ch = ((x And y) Xor ((Not x) And z)) And &HFFFFFFFF: End Function
Private Function Maj(x As Long, y As Long, z As Long) As Long: Maj = ((x And y) Xor (x And z) Xor (y And z)) And &HFFFFFFFF: End Function
