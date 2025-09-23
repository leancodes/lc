Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFileSizeEx Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpFileSize As Currency) As Long
Private Declare PtrSafe Function CreateFileMappingW Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpAttributes As LongPtr, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As LongPtr) As LongPtr
Private Declare PtrSafe Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As LongPtr) As LongPtr
Private Declare PtrSafe Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

Private Const GENERIC_READ As Long = &H80000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_DELETE As Long = &H4
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const PAGE_READONLY As Long = &H2
Private Const FILE_MAP_READ As Long = &H4

Private Type U64
    Lo As Long
    Hi As Long
End Type

Private Function CToLL(ByVal c As Currency) As LongLong
    CToLL = c * 10000@
End Function

Private Function ReadU64(ByVal p As LongPtr) As LongLong
    RtlMoveMemory ReadU64, p, 8
End Function

Private Function ReadU32(ByVal p As LongPtr) As Long
    RtlMoveMemory ReadU32, p, 4
End Function

Private Function ReadByte(ByVal p As LongPtr) As Byte
    RtlMoveMemory ReadByte, p, 1
End Function

Private Function CLngLng(ByVal x As Long) As LongLong
    Dim t As Currency
    t = x
    CLngLng = t
End Function

Private Function Lsr32(ByVal x As Long, ByVal n As Long) As Long
    Dim u As Double
    u = (x And &H7FFFFFFF) + IIf(x < 0, 2147483648#, 0#)
    Lsr32 = Fix(u / (2# ^ n))
End Function

Private Function RotL64(ByVal x As LongLong, ByVal r As Long) As LongLong
    Dim s As Long, lo As Long, hi As Long, nlo As Long, nhi As Long
    s = r And 63
    If s = 0 Then RotL64 = x: Exit Function
    lo = x And &HFFFFFFFF
    hi = (x And &HFFFFFFFF00000000) \ &H100000000
    If s < 32 Then
        nlo = ((lo * (2 ^ s)) And &HFFFFFFFF) Or Lsr32(hi, 32 - s)
        nhi = ((hi * (2 ^ s)) And &HFFFFFFFF) Or Lsr32(lo, 32 - s)
    ElseIf s = 32 Then
        nlo = hi
        nhi = lo
    Else
        s = s - 32
        nlo = ((hi * (2 ^ s)) And &HFFFFFFFF) Or Lsr32(lo, 32 - s)
        nhi = ((lo * (2 ^ s)) And &HFFFFFFFF) Or Lsr32(hi, 32 - s)
    End If
    RotL64 = (CLngLng(nhi) * &H100000000) Or (nlo And &HFFFFFFFF)
End Function

Private Function Mul64(ByVal a As LongLong, ByVal b As LongLong) As LongLong
    Dim al As Long, ah As Long, bl As Long, bh As Long
    Dim p0 As Double, p1 As Double, p2 As Double, p3 As Double
    Dim carry As Double, lo As Double, mid As Double, hi As Double
    al = a And &HFFFFFFFF
    ah = (a And &HFFFFFFFF00000000) \ &H100000000
    bl = b And &HFFFFFFFF
    bh = (b And &HFFFFFFFF00000000) \ &H100000000
    p0 = CDbl(al And &HFFFFFFFF) * CDbl(bl And &HFFFFFFFF)
    p1 = CDbl(al And &HFFFFFFFF) * CDbl(bh And &HFFFFFFFF)
    p2 = CDbl(bl And &HFFFFFFFF) * CDbl(ah And &HFFFFFFFF)
    p3 = CDbl(ah And &HFFFFFFFF) * CDbl(bh And &HFFFFFFFF)
    carry = Fix(p0 / 4294967296#)
    mid = p1 + p2 + carry
    lo = (p0 - Fix(p0 / 4294967296#) * 4294967296#) + ((mid - Fix(mid / 4294967296#) * 4294967296#) * 4294967296#)
    hi = p3 + Fix(mid / 4294967296#)
    Mul64 = ((Fix(hi) * 4294967296#) + (lo - Fix(lo / 1#) * 1#))
End Function

Private Function Xor64(ByVal a As LongLong, ByVal b As LongLong) As LongLong
    Dim aa As Long, ab As Long, ba As Long, bb As Long, rl As Long, rh As Long
    aa = a And &HFFFFFFFF
    ab = (a And &HFFFFFFFF00000000) \ &H100000000
    ba = b And &HFFFFFFFF
    bb = (b And &HFFFFFFFF00000000) \ &H100000000
    rl = (aa Xor ba)
    rh = (ab Xor bb)
    Xor64 = rl Or (CLngLng(rh) * &H100000000)
End Function

Private Function ToHex64(ByVal v As LongLong) As String
    Dim u As U64, s As String
    RtlMoveMemory u, VarPtr(v), 8
    s = Right$("00000000" & Hex$(u.Hi And &HFFFFFFFF), 8) & Right$("00000000" & Hex$(u.Lo And &HFFFFFFFF), 8)
    ToHex64 = s
End Function

Private Function xxh64_block(ByVal p As LongPtr, ByVal len As LongLong, ByVal seed As LongLong) As LongLong
    Const P1 As LongLong = &H9E3779B185EBCA87
    Const P2 As LongLong = &HC2B2AE3D27D4EB4F
    Const P3 As LongLong = &H165667B19E3779F9
    Const P4 As LongLong = &H85EBCA77C2B2AE63
    Const P5 As LongLong = &H27D4EB2F165667C5
    Dim h As LongLong, v1 As LongLong, v2 As LongLong, v3 As LongLong, v4 As LongLong
    Dim q As LongPtr, endp As LongPtr, limit As LongPtr
    q = p
    If len >= 32 Then
        v1 = seed + P1 + P2
        v2 = seed + P2
        v3 = seed
        v4 = seed - P1
        limit = p + (len - 32)
        Do
            v1 = Mul64(RotL64(v1 + ReadU64(q) * P2, 31), P1): q = q + 8
            v2 = Mul64(RotL64(v2 + ReadU64(q) * P2, 31), P1): q = q + 8
            v3 = Mul64(RotL64(v3 + ReadU64(q) * P2, 31), P1): q = q + 8
            v4 = Mul64(RotL64(v4 + ReadU64(q) * P2, 31), P1): q = q + 8
        Loop While q <= limit
        h = RotL64(v1, 1) + RotL64(v2, 7) + RotL64(v3, 12) + RotL64(v4, 18)
        v1 = Mul64(Xor64(v1, ReadU64(p) * P2), P1): h = Xor64(h, v1): h = h * P1 + P4: p = p + 8
        v2 = Mul64(Xor64(v2, ReadU64(p) * P2), P1): h = Xor64(h, v2): h = h * P1 + P4: p = p + 8
        v3 = Mul64(Xor64(v3, ReadU64(p) * P2), P1): h = Xor64(h, v3): h = h * P1 + P4: p = p + 8
        v4 = Mul64(Xor64(v4, ReadU64(p) * P2), P1): h = Xor64(h, v4): h = h * P1 + P4: p = p + 8
    Else
        h = seed + P5
    End If
    h = Xor64(h, len)
    endp = p + len
    Do While (endp - p) >= 8
        h = Xor64(h, Mul64(RotL64(ReadU64(p) * P2, 31), P1)): h = RotL64(h, 27) * P1 + P4: p = p + 8
    Loop
    If (endp - p) >= 4 Then
        h = Xor64(h, CLngLng(ReadU32(p)) * P1): h = RotL64(h, 23) * P2 + P3: p = p + 4
    End If
    Do While p < endp
        h = Xor64(h, CLngLng(ReadByte(p)) * P5): h = RotL64(h, 11) * P1: p = p + 1
    Loop
    h = Xor64(h, h \ &H100000000000000) * P2
    h = Xor64(h, h \ &H100000000000000) * P3
    h = Xor64(h, h \ &H100000000000000)
    xxh64_block = h
End Function

Public Function XXH64_FileHex(ByVal filePath As String, Optional ByVal seed As LongLong = 0) As String
    Dim t0 As Currency, t1 As Currency, f As LongPtr, sz As Currency, map As LongPtr, hmap As LongPtr, h As LongLong, freq As Currency
    Dim okFreq As Long, errCode As Long, timing As String
    If LenB(filePath) = 0 Then XXH64_FileHex = "ERR E_PATH": Exit Function
    okFreq = QueryPerformanceFrequency(freq)
    QueryPerformanceCounter t0
    f = CreateFileW(StrPtr(filePath), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_DELETE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If f = 0 Or f = -1 Then errCode = Err.LastDllError: GoTo Q
    If GetFileSizeEx(f, sz) = 0 Then errCode = Err.LastDllError: GoTo Q
    hmap = CreateFileMappingW(f, 0, PAGE_READONLY, 0, 0, 0)
    If hmap = 0 Then errCode = Err.LastDllError: GoTo Q
    map = MapViewOfFile(hmap, FILE_MAP_READ, 0, 0, 0)
    If map = 0 Then errCode = Err.LastDllError: GoTo Q
    h = xxh64_block(map, CToLL(sz), seed)
Q:
    If map <> 0 Then UnmapViewOfFile map
    If hmap <> 0 Then CloseHandle hmap
    If f <> 0 And f <> -1 Then CloseHandle f
    QueryPerformanceCounter t1
    If okFreq <> 0 Then
        timing = " " & Format$((CToLL(t1 - t0)) / (CToLL(freq)), "0.000000") & "s"
    Else
        timing = ""
    End If
    If errCode <> 0 And h = 0 Then
        XXH64_FileHex = "ERR " & Hex$(errCode) & timing
    Else
        XXH64_FileHex = ToHex64(h) & timing
    End If
End Function

Public Sub Demo()
    Dim r As String
    r = XXH64_FileHex("C:\Windows\explorer.exe")
    Debug.Print r
End Sub
