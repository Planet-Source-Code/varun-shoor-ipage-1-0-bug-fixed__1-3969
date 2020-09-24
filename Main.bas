Attribute VB_Name = "Main"
'***************************'

Type Com
Reply As String
BackCommand As String
End Type

'***************************'

Function fnC8(ByVal p0012 As String) As Variant
Dim l0014 As String 'this is a function that changes text into Alternative caps
Dim l0016 As Integer
Dim l0018 As Integer
Dim l001A As String
Dim l001C As String
Let l0014$ = p0012
Let l0016% = Len(l0014$)
Do While l0018% <= l0016%
Let l0018% = l0018% + 1
Let l001A$ = Mid$(l0014$, l0018%, 1)
Let l001A$ = UCase$(l001A$)
Let l001C$ = l001C$ + l001A$
Let l0018% = l0018% + 1
Let l001A$ = Mid$(l0014$, l0018%, 1)
Let l001A$ = LCase$(l001A$)
Let l001C$ = l001C$ + l001A$
Loop
fnC8 = l001C$
End Function

'***************************'

Function fn100(ByVal p0020 As String) As String
Dim l0022 As String 'this is a function that changes the text to backwards
Dim l0024 As Integer
l0022$ = ""
For l0024% = Len(p0020$) To 1 Step -1
l0022$ = l0022$ + Mid$(p0020$, l0024%, 1)
Next l0024%
fn100$ = l0022$
End Function

'***************************'

Function fn138(ByVal p002A As String) As Variant
Dim l002C As String 'this is just an ascci function that changes the text into elie form
Dim l002E As Integer
Dim l0030 As Integer
Dim l0032 As String
Dim l0034 As String
Dim l0036 As String
Dim l0038 As Integer
Let l002C$ = p002A
Let l002E% = Len(l002C$)
Do While l0030% <= l002E%
DoEvents
Let l0030% = l0030% + 1
Let l0032$ = Mid$(l002C$, l0030%, 1)
Let l0034$ = Mid$(l002C$, l0030%, 2)
If l0034$ = "ae" Then Let l0034$ = "æ": Let l0036$ = l0036$ + l0034$: Let l0038% = 2: GoTo L1382
If l0034$ = "AE" Then Let l0034$ = "Æ": Let l0036$ = l0036$ + l0034$: Let l0038% = 2: GoTo L1382
If l0034$ = "oe" Then Let l0034$ = "œ": Let l0036$ = l0036$ + l0034$: Let l0038% = 2: GoTo L1382
If l0034$ = "OE" Then Let l0034$ = "Œ": Let l0036$ = l0036$ + l0034$: Let l0038% = 2: GoTo L1382
If l0038% > 0 Then GoTo L1382
If l0032$ = "A" Then Let l0032$ = "/\"
If l0032$ = "a" Then Let l0032$ = "å"
If l0032$ = "B" Then Let l0032$ = "ß"
If l0032$ = "C" Then Let l0032$ = "Ç"
If l0032$ = "c" Then Let l0032$ = "¢"
If l0032$ = "D" Then Let l0032$ = "Ð"
If l0032$ = "d" Then Let l0032$ = "ð"
If l0032$ = "E" Then Let l0032$ = "Ê"
If l0032$ = "e" Then Let l0032$ = "è"
If l0032$ = "f" Then Let l0032$ = "ƒ"
If l0032$ = "H" Then Let l0032$ = "]-["
If l0032$ = "I" Then Let l0032$ = "¡"
If l0032$ = "i" Then Let l0032$ = "î"
If l0032$ = "k" Then Let l0032$ = "|‹"
If l0032$ = "L" Then Let l0032$ = "|_"
If l0032$ = "l" Then Let l0032$ = "£"
If l0032$ = "M" Then Let l0032$ = "|V|"
If l0032$ = "m" Then Let l0032$ = "^^"
If l0032$ = "N" Then Let l0032$ = "]\["
If l0032$ = "n" Then Let l0032$ = "ñ"
If l0032$ = "O" Then Let l0032$ = "Ø"
If l0032$ = "o" Then Let l0032$ = "º"
If l0032$ = "P" Then Let l0032$ = "¶"
If l0032$ = "p" Then Let l0032$ = "Þ"
If l0032$ = "R" Then Let l0032$ = "|2"
If l0032$ = "r" Then Let l0032$ = "®"
If l0032$ = "S" Then Let l0032$ = "§"
If l0032$ = "s" Then Let l0032$ = "$"
If l0032$ = "T" Then Let l0032$ = "¯|¯"
If l0032$ = "t" Then Let l0032$ = "†"
If l0032$ = "U" Then Let l0032$ = "Ú"
If l0032$ = "u" Then Let l0032$ = "µ"
If l0032$ = "V" Then Let l0032$ = "\/"
If l0032$ = "W" Then Let l0032$ = "VV"
If l0032$ = "w" Then Let l0032$ = "vv"
If l0032$ = "X" Then Let l0032$ = "X"
If l0032$ = "x" Then Let l0032$ = "×"
If l0032$ = "Y" Then Let l0032$ = "¥"
If l0032$ = "y" Then Let l0032$ = "ý"
If l0032$ = "!" Then Let l0032$ = "¡"
If l0032$ = "?" Then Let l0032$ = "¿"
If l0032$ = "." Then Let l0032$ = "…"
If l0032$ = "," Then Let l0032$ = "‚"
If l0032$ = "1" Then Let l0032$ = "¹"
If l0032$ = "%" Then Let l0032$ = "‰"
If l0032$ = "2" Then Let l0032$ = "²"
If l0032$ = "3" Then Let l0032$ = "³"
If l0032$ = "_" Then Let l0032$ = "¯"
If l0032$ = "-" Then Let l0032$ = "—"
If l0032$ = " " Then Let l0032$ = " "
Let l0036$ = l0036$ + l0032$
L1382:
If l0038% > 0 Then Let l0038% = l0038% - 1
DoEvents
Loop
fn138 = l0036$
End Function

'***************************'

Function fn170(ByVal p003E As String) As Variant
Dim l0040 As String 'this is a function that changes text like 'hello' to H3LL0 (hacking or number text)
Dim l0042 As Integer
Dim l0044 As Integer
Dim l0046 As String
Dim l0048 As Integer
Dim l004A As String
Let l0040$ = p003E
Let l0042% = Len(l0040$)
Do While l0044% <= l0042%
DoEvents
Let l0044% = l0044% + 1
Let l0046$ = Mid$(l0040$, l0044%, 1)
If l0048% > 0 Then GoTo L1B62
If l0046$ = "A" Then Let l0046$ = "4"
If l0046$ = "a" Then Let l0046$ = "4"
If l0046$ = "B" Then Let l0046$ = "8"
If l0046$ = "b" Then Let l0046$ = "6"
If l0046$ = "C" Then Let l0046$ = "C"
If l0046$ = "c" Then Let l0046$ = "C"
If l0046$ = "D" Then Let l0046$ = "D"
If l0046$ = "d" Then Let l0046$ = "D"
If l0046$ = "E" Then Let l0046$ = "3"
If l0046$ = "e" Then Let l0046$ = "3"
If l0046$ = "F" Then Let l0046$ = "F"
If l0046$ = "f" Then Let l0046$ = "F"
If l0046$ = "H" Then Let l0046$ = "H"
If l0046$ = "h" Then Let l0046$ = "H"
If l0046$ = "I" Then Let l0046$ = "1"
If l0046$ = "i" Then Let l0046$ = "1"
If l0046$ = "J" Then Let l0046$ = "J"
If l0046$ = "j" Then Let l0046$ = "J"
If l0046$ = "K" Then Let l0046$ = "K"
If l0046$ = "k" Then Let l0046$ = "K"
If l0046$ = "L" Then Let l0046$ = "L"
If l0046$ = "l" Then Let l0046$ = "L"
If l0046$ = "M" Then Let l0046$ = "M"
If l0046$ = "m" Then Let l0046$ = "M"
If l0046$ = "N" Then Let l0046$ = "N"
If l0046$ = "n" Then Let l0046$ = "N"
If l0046$ = "O" Then Let l0046$ = "0"
If l0046$ = "o" Then Let l0046$ = "0"
If l0046$ = "P" Then Let l0046$ = "P"
If l0046$ = "p" Then Let l0046$ = "P"
If l0046$ = "Q" Then Let l0046$ = "Q"
If l0046$ = "q" Then Let l0046$ = "Q"
If l0046$ = "R" Then Let l0046$ = "2"
If l0046$ = "r" Then Let l0046$ = "2"
If l0046$ = "S" Then Let l0046$ = "5"
If l0046$ = "s" Then Let l0046$ = "5"
If l0046$ = "T" Then Let l0046$ = "7"
If l0046$ = "t" Then Let l0046$ = "7"
If l0046$ = "U" Then Let l0046$ = "U"
If l0046$ = "u" Then Let l0046$ = "U"
If l0046$ = "V" Then Let l0046$ = "V"
If l0046$ = "v" Then Let l0046$ = "V"
If l0046$ = "W" Then Let l0046$ = "W"
If l0046$ = "w" Then Let l0046$ = "W"
If l0046$ = "X" Then Let l0046$ = "X"
If l0046$ = "x" Then Let l0046$ = "X"
If l0046$ = "Y" Then Let l0046$ = "Y"
If l0046$ = "y" Then Let l0046$ = "Y"
Let l004A$ = l004A$ + l0046$
L1B62:
If l0048% > 0 Then Let l0048% = l0048% - 1
DoEvents
Loop
fn170 = l004A$
End Function

'***************************'

Function EchoText(Text As String, Reverse As Boolean)
'This will "echo" the text, like this:  Cool ool ol l
On Error GoTo error
Dim i As Long
Dim endstr As String
For i = 1 To Len(Text$)
If Reverse = True Then
endstr$ = Mid$(Text$, i, Len(Text$) - (i - 1)) & " " & endstr$
Else
endstr$ = endstr$ & Mid$(Text$, i, Len(Text$) - (i - 1)) & " "
End If
Next i
endstr$ = Mid$(endstr$, 1, Len(endstr$) - 1)
EchoText = endstr$
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

'***************************'

Function ReplaceChars(Chars As String, Optional ReplaceChr As String, Optional ReplaceWith As String) As String
Dim ChrCnt As Long
If ReplaceChr = "" Then ReplaceChr = " "
ChrCnt = 1
Do
ChrCnt = InStr(ChrCnt, Chars, ReplaceChr)
If ChrCnt = 0 Then Exit Do
Chars = Left$(Chars, ChrCnt - 1) & ReplaceWith & Right(Chars, Len(Chars) + 1 - Len(ReplaceChr) - ChrCnt)
ChrCnt = ChrCnt + Len(ReplaceWith)
Loop
ReplaceChars = Chars
End Function

'***************************'

Public Function GetInBetween(Text As String, Textstring1 As String, Textstring2 As String) As String
If InStr(Text, Textstring1) = 0 Or InStr(Text, Textstring2) = 0 Then
GetInBetween = "STRING NOT FOUND"
Exit Function
End If
GetInBetween = Mid$(Text, InStr(Text, Textstring1) + Len(Textstring1), InStr(Text, Textstring2) - (Len(Textstring1) + InStr(Text, Textstring1)))
End Function

'***************************'

Public Function GetInBetween1(Text As String, Textstring1 As String, Textstring2 As String) As Integer
If InStr(Text, Textstring1) = 0 Or InStr(Text, Textstring2) = 0 Then
GetInBetween1 = "0"
Exit Function
End If
GetInBetween1 = Mid$(Text, InStr(Text, Textstring1) + Len(Textstring1), InStr(Text, Textstring2) - (Len(Textstring1) + InStr(Text, Textstring1)))
End Function
