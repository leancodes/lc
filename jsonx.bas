Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const CF_UNICODETEXT As Long = 13

Private mPrevCalc As XlCalculation
Private mPrevScr As Boolean
Private mPrevEvents As Boolean
Private mPrevStatus As Variant
Private mStateCaptured As Boolean

Private Const MaxChars As Long = 5000000
Private Const MaxRows As Long = 100000
Private Const MaxColumns As Long = 1000

Public Sub PasteJsonTable()
Dim ws As Worksheet
Set ws = Application.ActiveSheet

Dim jsonText As String
jsonText = GetJsonFromClipboardText()
If Len(jsonText) = 0 Then
    If Not Application.Selection Is Nothing Then
        If Application.Selection.Cells.Count = 1 Then
            jsonText = CStr(Application.Selection.Value)
        End If
    End If
End If
If Len(jsonText) = 0 Then
    MsgBox "No JSON found.", vbExclamation
    Exit Sub
End If
If Len(jsonText) > MaxChars Then
    MsgBox "JSON too large.", vbCritical
    Exit Sub
End If

SafeAppOff("Parsing JSON...")
On Error GoTo CleanFail

Dim root As Variant
root = Json_Parse(jsonText)

Dim rows As Collection
Set rows = Json_To_Rows(root)
If rows.Count = 0 Then
    MsgBox "No rows parsed.", vbInformation
    GoTo CleanExit
End If
If rows.Count > MaxRows Then Err.Raise vbObjectError + 2001, , "Row limit exceeded."

Dim headers As Collection
Set headers = CollectHeaders(rows)
If headers.Count = 0 Then Err.Raise vbObjectError + 2002, , "No headers."
If headers.Count > MaxColumns Then Err.Raise vbObjectError + 2003, , "Column limit exceeded."

Dim dataArr As Variant
dataArr = RowsTo2D(rows, headers)

Dim outRange As Range
Set outRange = Application.ActiveCell.Resize(UBound(dataArr, 1), UBound(dataArr, 2))

outRange.Value = dataArr
outRange.Rows(1).Font.Bold = True
outRange.EntireColumn.AutoFit

CleanExit:
SafeAppOn
Exit Sub
CleanFail:
SafeAppOn
MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub SafeAppOff(ByVal statusText As String)
On Error Resume Next
If Not mStateCaptured Then
mPrevScr = Application.ScreenUpdating
mPrevEvents = Application.EnableEvents
mPrevCalc = Application.Calculation
mPrevStatus = Application.StatusBar
mStateCaptured = True
End If
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.StatusBar = statusText
End Sub

Private Sub SafeAppOn()
On Error Resume Next
If mStateCaptured Then
Application.ScreenUpdating = mPrevScr
Application.EnableEvents = mPrevEvents
Application.Calculation = mPrevCalc
Application.StatusBar = mPrevStatus
mStateCaptured = False
End If
End Sub

Private Function GetJsonFromClipboardText() As String
Dim ret As String
#If VBA7 Then
Dim hData As LongPtr, pData As LongPtr
#Else
Dim hData As Long, pData As Long
#End If
Dim attempts As Long
If IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 Then Exit Function
For attempts = 1 To 5
If OpenClipboard(0) <> 0 Then Exit For
Sleep 30
Next attempts
If attempts = 6 Then Exit Function
On Error GoTo CleanClip
hData = GetClipboardData(CF_UNICODETEXT)
If hData <> 0 Then
pData = GlobalLock(hData)
If pData <> 0 Then
Dim cch As Long
cch = lstrlenW(pData)
If cch > 0 Then
Dim bytes() As Byte
ReDim bytes(0 To cch * 2 - 1)
CopyMemory bytes(0), pData, cch * 2
ret = StrConv(bytes, vbUnicode)
End If
GlobalUnlock hData
End If
End If
CleanClip:
CloseClipboard
GetJsonFromClipboardText = ret
End Function

Private Function Json_To_Rows(ByVal root As Variant) As Collection
Dim rows As New Collection
Dim t As String: t = TypeName(root)
If t = "Collection" Then
Dim i As Long
For i = 1 To root.Count
Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
Flatten root(i), row, ""
rows.Add row
Next
ElseIf IsObject(root) Then
Dim row2 As Object: Set row2 = CreateObject("Scripting.Dictionary")
Flatten root, row2, ""
rows.Add row2
Else
Dim row3 As Object: Set row3 = CreateObject("Scripting.Dictionary")
row3.Add "value", root
rows.Add row3
End If
Set Json_To_Rows = rows
End Function

Private Sub Flatten(ByVal v As Variant, ByVal bag As Object, ByVal prefix As String, Optional ByVal depth As Long = 0)
Const MAX_DEPTH As Long = 128
If depth > MAX_DEPTH Then Err.Raise vbObjectError + 1001, , "JSON nesting too deep."
If IsObject(v) Then
Dim tn As String: tn = TypeName(v)
If tn = "Collection" Then
Dim i As Long
For i = 1 To v.Count
Flatten v(i), bag, IIf(prefix = "", "[" & (i - 1) & "]", prefix & "[" & (i - 1) & "]"), depth + 1
Next
Else
Dim k As Variant
For Each k In v.Keys
Flatten v(k), bag, IIf(prefix = "", CStr(k), prefix & "." & CStr(k)), depth + 1
Next k
End If
Else
bag(prefix) = v
End If
End Sub

Private Function CollectHeaders(ByVal rows As Collection) As Collection
Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
Dim out As New Collection
Dim i As Long, k As Variant
For i = 1 To rows.Count
For Each k In rows(i).Keys
If Not seen.Exists(k) Then
seen.Add k, True
out.Add k
If out.Count >= MaxColumns Then Exit For
End If
Next k
If out.Count >= MaxColumns Then Exit For
Next i
Set CollectHeaders = out
End Function

Private Function RowsTo2D(ByVal rows As Collection, ByVal headers As Collection) As Variant
Dim r As Long, c As Long, nR As Long, nC As Long
nR = rows.Count + 1
nC = headers.Count
Dim arr As Variant
ReDim arr(1 To nR, 1 To nC)
For c = 1 To nC
arr(1, c) = headers(c)
Next c
For r = 2 To nR
Dim d As Object: Set d = rows(r - 1)
For c = 1 To nC
Dim h As String: h = headers(c)
If d.Exists(h) Then arr(r, c) = d(h)
Next c
Next r
RowsTo2D = arr
End Function

Private Type JP
s As String
i As Long
n As Long
End Type

Private Function Json_Parse(ByVal s As String) As Variant
Dim p As JP
p.s = s
p.i = 1
p.n = Len(s)
SkipWS p
Dim v As Variant
v = ParseValue(p)
SkipWS p
If p.i <= p.n Then Err.Raise vbObjectError + 1002, , "Trailing characters after JSON value at pos " & p.i
Json_Parse = v
End Function

Private Function ParseValue(ByRef p As JP) As Variant
SkipWS p
If p.i > p.n Then Err.Raise vbObjectError + 1003, , "Unexpected end of JSON."
Select Case Mid$(p.s, p.i, 1)
Case """": ParseValue = ParseString(p)
Case "{": ParseValue = ParseObject(p)
Case "[": ParseValue = ParseArray(p)
Case "t": ExpectLiteral p, "true": ParseValue = True
Case "f": ExpectLiteral p, "false": ParseValue = False
Case "n": ExpectLiteral p, "null": ParseValue = Empty
Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9": ParseValue = ParseNumber(p)
Case Else: Err.Raise vbObjectError + 1004, , "Unexpected character at pos " & p.i
End Select
End Function

Private Function ParseObject(ByRef p As JP) As Variant
p.i = p.i + 1: SkipWS p
Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
If Peek(p) = "}" Then p.i = p.i + 1: Set ParseObject = d: Exit Function
Do
SkipWS p
If Peek(p) <> """" Then Err.Raise vbObjectError + 1005, , "Expected string key at pos " & p.i
Dim key As String: key = ParseString(p)
SkipWS p
If Peek(p) <> ":" Then Err.Raise vbObjectError + 1006, , "Expected ':' after key at pos " & p.i
p.i = p.i + 1
Dim val As Variant: val = ParseValue(p)
If d.Exists(key) Then d(key) = val Else d.Add key, val
SkipWS p
Select Case Peek(p)
Case "}": p.i = p.i + 1: Exit Do
Case ",": p.i = p.i + 1
Case Else: Err.Raise vbObjectError + 1007, , "Expected ',' or '}' at pos " & p.i
End Select
Loop
Set ParseObject = d
End Function

Private Function ParseArray(ByRef p As JP) As Variant
p.i = p.i + 1: SkipWS p
Dim a As New Collection
If Peek(p) = "]" Then p.i = p.i + 1: Set ParseArray = a: Exit Function
Do
a.Add ParseValue(p)
SkipWS p
Select Case Peek(p)
Case "]": p.i = p.i + 1: Exit Do
Case ",": p.i = p.i + 1
Case Else: Err.Raise vbObjectError + 1008, , "Expected ',' or ']' at pos " & p.i
End Select
Loop
Set ParseArray = a
End Function

Private Function ParseString(ByRef p As JP) As String
p.i = p.i + 1
Dim parts() As String
Dim count As Long: count = 0
Dim cap As Long: cap = 8
ReDim parts(1 To cap)
Dim startPos As Long: startPos = p.i
Do While p.i <= p.n
Dim ch As String: ch = Mid$(p.s, p.i, 1)
If ch = """" Then
If p.i > startPos Then
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = Mid$(p.s, startPos, p.i - startPos)
End If
p.i = p.i + 1
If count = 0 Then
ParseString = ""
ElseIf count = 1 Then
ParseString = parts(1)
Else
ReDim Preserve parts(1 To count)
ParseString = Join(parts, "")
End If
Exit Function
ElseIf ch <> "" Then
p.i = p.i + 1
Else
If p.i > startPos Then
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = Mid$(p.s, startPos, p.i - startPos)
End If
p.i = p.i + 1
If p.i > p.n Then Err.Raise vbObjectError + 1010, , "Bad escape at end of string."
Dim e As String: e = Mid$(p.s, p.i, 1)
Select Case e
Case """", "", "/"
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = e
p.i = p.i + 1
Case "b"
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = vbBack
p.i = p.i + 1
Case "f"
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = vbFormFeed
p.i = p.i + 1
Case "n"
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = vbLf
p.i = p.i + 1
Case "r"
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = vbCr
p.i = p.i + 1
Case "t"
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = vbTab
p.i = p.i + 1
Case "u"
Dim hi As Long, lo As Long
Dim hex As String: hex = Mid$(p.s, p.i + 1, 4)
If Len(hex) < 4 Or Not IsHex4(hex) Then Err.Raise vbObjectError + 1011, , "Invalid \u escape at pos " & p.i
hi = CLng("&H" & hex)
p.i = p.i + 5
If hi >= &HD800 And hi <= &HDBFF Then
If p.i + 5 - 1 > p.n Or Mid$(p.s, p.i, 2) <> "\u" Then Err.Raise vbObjectError + 1012, , "High surrogate not followed by low surrogate at pos " & p.i
hex = Mid$(p.s, p.i + 2, 4)
If Len(hex) < 4 Or Not IsHex4(hex) Then Err.Raise vbObjectError + 1012, , "Invalid low surrogate at pos " & p.i
lo = CLng("&H" & hex)
If lo < &HDC00 Or lo > &HDFFF Then Err.Raise vbObjectError + 1012, , "Expected low surrogate at pos " & p.i
p.i = p.i + 6
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = ChrW$(hi) & ChrW$(lo)
Else
count = count + 1
If count > cap Then cap = cap * 2: ReDim Preserve parts(1 To cap)
parts(count) = ChrW$(hi)
End If
Case Else
Err.Raise vbObjectError + 1012, , "Invalid escape '\" & e & "' at pos " & p.i
End Select
startPos = p.i
End If
Loop
Err.Raise vbObjectError + 1013, , "Unterminated string."
End Function

Private Function ParseNumber(ByRef p As JP) As Variant
Dim startI As Long: startI = p.i
Dim ch As String
If Peek(p) = "-" Then p.i = p.i + 1
ch = Peek(p)
If ch = "0" Then
p.i = p.i + 1
ElseIf ch >= "1" And ch <= "9" Then
Do While p.i <= p.n
ch = Mid$(p.s, p.i, 1)
If ch < "0" Or ch > "9" Then Exit Do
p.i = p.i + 1
Loop
Else
Err.Raise vbObjectError + 1014, , "Invalid number at pos " & p.i
End If
If Peek(p) = "." Then
p.i = p.i + 1
If Mid$(p.s, p.i, 1) < "0" Or Mid$(p.s, p.i, 1) > "9" Then Err.Raise vbObjectError + 1015, , "Invalid fraction at pos " & p.i
Do While p.i <= p.n
ch = Mid$(p.s, p.i, 1)
If ch < "0" Or ch > "9" Then Exit Do
p.i = p.i + 1
Loop
End If
ch = Peek(p)

If ch = "e" Or ch = "E" Then
p.i = p.i + 1
If Peek(p) = "+" Or Peek(p) = "-" Then p.i = p.i + 1
If Mid$(p.s, p.i, 1) < "0" Or Mid$(p.s, p.i, 1) > "9" Then Err.Raise vbObjectError + 1016, , "Invalid exponent at pos " & p.i
Do While p.i <= p.n
ch = Mid$(p.s, p.i, 1)
If ch < "0" Or ch > "9" Then Exit Do
p.i = p.i + 1
Loop
End If
Dim numStr As String: numStr = Mid$(p.s, startI, p.i - startI)
If InStr(1, numStr, ".", vbBinaryCompare) > 0 Or InStr(1, numStr, "e", vbTextCompare) > 0 Then
ParseNumber = Val(numStr)
Else
If Len(numStr) > 28 Then
ParseNumber = numStr
Else
On Error GoTo asString
ParseNumber = CDec(numStr)
Exit Function
asString:
ParseNumber = numStr
End If
End If
End Function

Private Sub ExpectLiteral(ByRef p As JP, ByVal lit As String)
If Mid$(p.s, p.i, Len(lit)) <> lit Then Err.Raise vbObjectError + 1017, , "Expected '" & lit & "' at pos " & p.i
p.i = p.i + Len(lit)
End Sub

Private Sub SkipWS(ByRef p As JP)
Do While p.i <= p.n
Select Case Mid$(p.s, p.i, 1)
Case " ", vbTab, vbCr, vbLf: p.i = p.i + 1
Case Else: Exit Do
End Select
Loop
End Sub

Private Function Peek(ByRef p As JP) As String
If p.i > p.n Then
Peek = ""
Else
Peek = Mid$(p.s, p.i, 1)
End If
End Function

Private Function IsHex4(ByVal s As String) As Boolean
Dim i As Long, ch As String
If Len(s) <> 4 Then Exit Function
For i = 1 To 4
ch = Mid$(s, i, 1)
If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then Exit Function
Next i
IsHex4 = True
End Function
