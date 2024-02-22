Attribute VB_Name = "BlowFish"
Option Explicit

Private pKey() As Byte

Public blf_P(17) As Long
Public blf_S(3, 255) As Long

Private Const ncROUNDS  As Integer = 16
Private Const ncMAXKEYLEN As Integer = 56

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal lpDestination As Any, ByVal lpSource As Any, ByVal Length As Long)

Private aDecTab(255) As Integer
Private aEncTab(63) As Byte

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647

Sub SetKey(pass As String)
 On Error GoTo fallo
    pKey = StrConv(pass, vbFromUnicode)
    Call blf_KeyInit(pKey)
Exit Sub
fallo:
Call LogError("SETKEY " & Err.number & " D: " & Err.Description)

End Sub

Function EncriptaString(str As String) As String
   On Error GoTo fallo
    Dim AbCip() As Byte, AbStr() As Byte
    AbStr = StrConv(str, vbFromUnicode)
    AbCip = blf_BytesEnc(AbStr)
    EncriptaString = EncodeBytes64(AbCip)
Exit Function
fallo:
Call LogError("ENCRIPTASTRING " & Err.number & " D: " & Err.Description)

End Function

Function DesencriptaString(str As String) As String
    On Error GoTo fallo
    Dim AbPlain() As Byte, AbDes() As Byte
    AbPlain = DecodeBytes64(str)
    AbDes = blf_BytesDec(AbPlain)
    DesencriptaString = StrConv(AbDes, vbUnicode)
Exit Function
fallo:
'Call LogError("DESENCRIPTASTRING " & Err.Number & " D: " & Err.Description)

End Function


'CODIGO ORIGINAL

Public Function uwJoin(a As Byte, b As Byte, C As Byte, d As Byte) As Long
On Error GoTo fallo
' Added Version 5: replacement for uw_WordJoin
' Join 4 x 8-bit bytes into one 32-bit word a.b.c.d
    uwJoin = ((a And &H7F) * &H1000000) Or (b * &H10000) Or (CLng(C) * &H100) Or d
    If a And &H80 Then
        uwJoin = uwJoin Or &H80000000
    End If
Exit Function
fallo:
Call LogError("UWJOIN " & Err.number & " D: " & Err.Description)


End Function

Public Sub uwSplit(ByVal w As Long, a As Byte, b As Byte, C As Byte, d As Byte)
On Error GoTo fallo
' Added Version 5: replacement for uw_WordSplit
' Split 32-bit word w into 4 x 8-bit bytes
    a = CByte(((w And &HFF000000) \ &H1000000) And &HFF)
    b = CByte(((w And &HFF0000) \ &H10000) And &HFF)
    C = CByte(((w And &HFF00) \ &H100) And &HFF)
    d = CByte((w And &HFF) And &HFF)
Exit Sub
fallo:
Call LogError("UWSPLIT" & Err.number & " D: " & Err.Description)


End Sub

' Function re-written 11 May 2001.
Public Function uw_ShiftLeftBy8(wordX As Long) As Long
  On Error GoTo fallo
    ' Shift 32-bit long value to left by 8 bits
    ' i.e. VB equivalent of "wordX << 8" in C
    ' Avoiding problem with sign bit
    uw_ShiftLeftBy8 = (wordX And &H7FFFFF) * &H100
    If (wordX And &H800000) <> 0 Then
        uw_ShiftLeftBy8 = uw_ShiftLeftBy8 Or &H80000000
    End If
Exit Function
fallo:
Call LogError("UWSHIFTLEFTBY8 " & Err.number & " D: " & Err.Description)


End Function

Public Function uw_WordAdd(wordA As Long, wordB As Long) As Long
On Error GoTo fallo
' Adds words A and B avoiding overflow
    Dim myUnsigned As Double
    
    myUnsigned = LongToUnsigned(wordA) + LongToUnsigned(wordB)
    ' Cope with overflow
    If myUnsigned > OFFSET_4 Then
        myUnsigned = myUnsigned - OFFSET_4
    End If
    uw_WordAdd = UnsignedToLong(myUnsigned)
  Exit Function
fallo:
Call LogError("UW WORDADD " & Err.number & " D: " & Err.Description)
  
End Function

Public Function uw_WordSub(wordA As Long, wordB As Long) As Long
  On Error GoTo fallo
    Dim myUnsigned As Double
    myUnsigned = LongToUnsigned(wordA) - LongToUnsigned(wordB)
    If myUnsigned < 0 Then
        myUnsigned = myUnsigned + OFFSET_4
    End If
    uw_WordSub = UnsignedToLong(myUnsigned)
  Exit Function
fallo:
Call LogError("UW WORDSUB " & Err.number & " D: " & Err.Description)

End Function

Function UnsignedToLong(value As Double) As Long
    On Error GoTo fallo
    If value < 0 Or value >= OFFSET_4 Then Error 6 ' Overflow
    If value <= MAXINT_4 Then
        UnsignedToLong = value
    Else
        UnsignedToLong = value - OFFSET_4
    End If
  Exit Function
fallo:
Call LogError("USINGNEDTOLONG " & Err.number & " D: " & Err.Description)

End Function

Public Function LongToUnsigned(value As Long) As Double
   On Error GoTo fallo
    If value < 0 Then
        LongToUnsigned = value + OFFSET_4
    Else
        LongToUnsigned = value
    End If
  Exit Function
fallo:
Call LogError("LONGTOUSIGNED " & Err.number & " D: " & Err.Description)


End Function

Public Function EncodeBytes64(abBytes() As Byte) As String
  On Error GoTo fallo
    Dim sOutput As String
    Dim abOutput() As Byte
    Dim sLast As String
    Dim b(3) As Byte
    Dim j As Integer
    Dim i As Long, nLen As Long, nQuants As Long
    Dim iIndex As Long
    
    'Set up error handler to catch empty array
    
    nLen = UBound(abBytes) - LBound(abBytes) + 1
    nQuants = nLen \ 3
    iIndex = 0
    Call MakeEncTab
    If (nQuants > 0) Then
        ReDim abOutput(nQuants * 4 - 1)
        ' Now start reading in 3 bytes at a time
        For i = 0 To nQuants - 1
            For j = 0 To 2
               b(j) = abBytes((i * 3) + j)
            Next
            Call EncodeQuantumB(b)
            abOutput(iIndex) = b(0)
            abOutput(iIndex + 1) = b(1)
            abOutput(iIndex + 2) = b(2)
            abOutput(iIndex + 3) = b(3)
            iIndex = iIndex + 4
        Next
        sOutput = StrConv(abOutput, vbUnicode)
    End If
    
    ' Cope with odd bytes
    ' (no real performance hit by using strings here)
    Select Case nLen Mod 3
    Case 0
        sLast = ""
    Case 1
        b(0) = abBytes(nLen - 1)
        b(1) = 0
        b(2) = 0
        Call EncodeQuantumB(b)
        sLast = StrConv(b(), vbUnicode)
        ' Replace last 2 with =
        sLast = Left(sLast, 2) & "=="
    Case 2
        b(0) = abBytes(nLen - 2)
        b(1) = abBytes(nLen - 1)
        b(2) = 0
        Call EncodeQuantumB(b)
        sLast = StrConv(b(), vbUnicode)
        ' Replace last with =
        sLast = Left(sLast, 3) & "="
    End Select
    
    EncodeBytes64 = sOutput & sLast
    

  Exit Function
fallo:
Call LogError("ENCODEBYTES64 " & Err.number & " D: " & Err.Description)

End Function

Public Function DecodeBytes64(sEncoded As String) As Variant
On Error GoTo fallo
' Return Byte array of decoded binary values given base64 string
' Ignores any chars not in the 64-char subset
    Dim abDecoded() As Byte
    Dim d(3) As Byte
    Dim C As Integer        ' NB Integer to catch -1 value
    Dim di As Integer
    Dim i As Long
    Dim nLen As Long
    Dim iIndex As Long
    
    nLen = Len(sEncoded)
    If nLen < 4 Then
        ' Return an empty array
        DecodeBytes64 = abDecoded
        Exit Function
    End If
    ReDim abDecoded(((nLen \ 4) * 3) - 1)
    
    iIndex = 0
    di = 0
    Call MakeDecTab
    ' Read in each char in turn
    For i = 1 To Len(sEncoded)
        C = CByte(Asc(Mid(sEncoded, i, 1)))
        C = aDecTab(C)
        If C >= 0 Then
            d(di) = CByte(C)
            di = di + 1
            If di = 4 Then
                abDecoded(iIndex) = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
                iIndex = iIndex + 1
                abDecoded(iIndex) = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
                iIndex = iIndex + 1
                abDecoded(iIndex) = SHL6(d(2) And &H3) Or d(3)
                iIndex = iIndex + 1
                If d(3) = 64 Then
                    iIndex = iIndex - 1
                    abDecoded(iIndex) = 0
                End If
                If d(2) = 64 Then
                    iIndex = iIndex - 1
                    abDecoded(iIndex) = 0
                End If
                di = 0
            End If
        End If
    Next i
    ' Trim to correct length
    ReDim Preserve abDecoded(iIndex - 1)
    DecodeBytes64 = abDecoded

  Exit Function
fallo:
Call LogError("DECODEBYTES64 " & Err.number & " D: " & Err.Description)

End Function

Public Function EncodeStr64(sInput As String) As String
On Error GoTo fallo
' Return radix64 encoding of string of binary values
' Does not insert CRLFs. Just returns one long string,
' so it's up to the user to add line breaks or other formatting.
' Version 4: Use Byte array and StrConv - much faster
    Dim abOutput() As Byte  ' Version 4: now a Byte array
    Dim sLast As String
    Dim b(3) As Byte    ' Version 4: Now 3 not 2
    Dim j As Integer
    Dim i As Long, nLen As Long, nQuants As Long
    Dim iIndex As Long
    
    EncodeStr64 = ""
    nLen = Len(sInput)
    nQuants = nLen \ 3
    iIndex = 0
    Call MakeEncTab
    If (nQuants > 0) Then
        ReDim abOutput(nQuants * 4 - 1)
        ' Now start reading in 3 bytes at a time
        For i = 0 To nQuants - 1
            For j = 0 To 2
               b(j) = Asc(Mid(sInput, (i * 3) + j + 1, 1))
            Next
            Call EncodeQuantumB(b)
            abOutput(iIndex) = b(0)
            abOutput(iIndex + 1) = b(1)
            abOutput(iIndex + 2) = b(2)
            abOutput(iIndex + 3) = b(3)
            iIndex = iIndex + 4
        Next
        EncodeStr64 = StrConv(abOutput, vbUnicode)
    End If
    
    ' Cope with odd bytes
    ' (no real performance hit by using strings here)
    Select Case nLen Mod 3
    Case 0
        sLast = ""
    Case 1
        b(0) = Asc(Mid(sInput, nLen, 1))
        b(1) = 0
        b(2) = 0
        Call EncodeQuantumB(b)
        sLast = StrConv(b(), vbUnicode)
        ' Replace last 2 with =
        sLast = Left(sLast, 2) & "=="
    Case 2
        b(0) = Asc(Mid(sInput, nLen - 1, 1))
        b(1) = Asc(Mid(sInput, nLen, 1))
        b(2) = 0
        Call EncodeQuantumB(b)
        sLast = StrConv(b(), vbUnicode)
        ' Replace last with =
        sLast = Left(sLast, 3) & "="
    End Select
    
    EncodeStr64 = EncodeStr64 & sLast

  Exit Function
fallo:
Call LogError("ENCODESTR64 " & Err.number & " D: " & Err.Description)

End Function

Public Function DecodeStr64(sEncoded As String) As String
On Error GoTo fallo
' Return string of decoded binary values given radix64 string
' Ignores any chars not in the 64-char subset
' Version 4: Use Byte array and StrConv - much faster
    Dim abDecoded() As Byte 'Version 4: Now a Byte array
    Dim d(3) As Byte
    Dim C As Integer        ' NB Integer to catch -1 value
    Dim di As Integer
    Dim i As Long
    Dim nLen As Long
    Dim iIndex As Long
    
    nLen = Len(sEncoded)
    If nLen < 4 Then
        Exit Function
    End If
    ReDim abDecoded(((nLen \ 4) * 3) - 1) 'Version 4: Now base zero
    
    iIndex = 0  ' Version 4: Changed to base 0
    di = 0
    Call MakeDecTab
    ' Read in each char in turn
    For i = 1 To Len(sEncoded)
        C = CByte(Asc(Mid(sEncoded, i, 1)))
        C = aDecTab(C)
        If C >= 0 Then
            d(di) = CByte(C)    ' Version 3.1: add CByte()
            di = di + 1
            If di = 4 Then
                abDecoded(iIndex) = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
                iIndex = iIndex + 1
                abDecoded(iIndex) = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
                iIndex = iIndex + 1
                abDecoded(iIndex) = SHL6(d(2) And &H3) Or d(3)
                iIndex = iIndex + 1
                If d(3) = 64 Then
                    iIndex = iIndex - 1
                    abDecoded(iIndex) = 0
                End If
                If d(2) = 64 Then
                    iIndex = iIndex - 1
                    abDecoded(iIndex) = 0
                End If
                di = 0
            End If
        End If
    Next i
    ' Convert to a string
    DecodeStr64 = StrConv(abDecoded(), vbUnicode)
    ' Remove any unwanted trailing chars
    DecodeStr64 = Left(DecodeStr64, iIndex)
  Exit Function
fallo:
Call LogError("DECODESTR64 " & Err.number & " D: " & Err.Description)


End Function

Private Sub EncodeQuantumB(b() As Byte)
' Expects at least 4 bytes in b, i.e. Dim b(3) As Byte
 On Error GoTo fallo
    Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
     
    b0 = SHR2(b(0)) And &H3F
    b1 = SHL4(b(0) And &H3) Or (SHR4(b(1)) And &HF)
    b2 = SHL2(b(1) And &HF) Or (SHR6(b(2)) And &H3)
    b3 = b(2) And &H3F
    
    b(0) = aEncTab(b0)
    b(1) = aEncTab(b1)
    b(2) = aEncTab(b2)
    b(3) = aEncTab(b3)
      Exit Sub
fallo:
Call LogError("ENCODEQUANTUMB " & Err.number & " D: " & Err.Description)

End Sub


Private Function MakeDecTab()
On Error GoTo fallo
' Set up Radix 64 decoding table
    Dim t As Integer
    Dim C As Integer

    For C = 0 To 255
        aDecTab(C) = -1
    Next
  
    t = 0
    For C = Asc("A") To Asc("Z")
        aDecTab(C) = t
        t = t + 1
    Next
  
    For C = Asc("a") To Asc("z")
        aDecTab(C) = t
        t = t + 1
    Next
    
    For C = Asc("0") To Asc("9")
        aDecTab(C) = t
        t = t + 1
    Next
    
    C = Asc("+")
    aDecTab(C) = t
    t = t + 1
    
    C = Asc("/")
    aDecTab(C) = t
    t = t + 1
    
    C = Asc("=")    ' flag for the byte-deleting char
    aDecTab(C) = t  ' should be 64
  Exit Function
fallo:
Call LogError("MAKEDECTAB " & Err.number & " D: " & Err.Description)

End Function

Private Function MakeEncTab()
On Error GoTo fallo
' Set up Radix 64 encoding table in bytes
    Dim i As Integer
    Dim C As Integer

    i = 0
    For C = Asc("A") To Asc("Z")
        aEncTab(i) = C
        i = i + 1
    Next
  
    For C = Asc("a") To Asc("z")
        aEncTab(i) = C
        i = i + 1
    Next
    
    For C = Asc("0") To Asc("9")
        aEncTab(i) = C
        i = i + 1
    Next
    
    C = Asc("+")
    aEncTab(i) = C
    i = i + 1
    
    C = Asc("/")
    aEncTab(i) = C
    i = i + 1
   Exit Function
fallo:
Call LogError("MAKEENCTAB " & Err.number & " D: " & Err.Description)
   
End Function

' Version 3: ShiftLeft and ShiftRight functions improved.
Private Function SHL2(ByVal bytValue As Byte) As Byte
On Error GoTo fallo
' Shift 8-bit value to left by 2 bits
' i.e. VB equivalent of "bytValue << 2" in C
    SHL2 = (bytValue * &H4) And &HFF

Exit Function
fallo:
Call LogError("SHL2 " & Err.number & " D: " & Err.Description)


End Function

Private Function SHL4(ByVal bytValue As Byte) As Byte
On Error GoTo fallo
' Shift 8-bit value to left by 4 bits
' i.e. VB equivalent of "bytValue << 4" in C
    SHL4 = (bytValue * &H10) And &HFF
  Exit Function
fallo:
Call LogError("SHL4 " & Err.number & " D: " & Err.Description)

End Function

Private Function SHL6(ByVal bytValue As Byte) As Byte
On Error GoTo fallo
' Shift 8-bit value to left by 6 bits
' i.e. VB equivalent of "bytValue << 6" in C
    SHL6 = (bytValue * &H40) And &HFF
  Exit Function
fallo:
Call LogError("SHL6" & Err.number & " D: " & Err.Description)

End Function

Private Function SHR2(ByVal bytValue As Byte) As Byte
On Error GoTo fallo
' Shift 8-bit value to right by 2 bits
' i.e. VB equivalent of "bytValue >> 2" in C
    SHR2 = bytValue \ &H4
  Exit Function
fallo:
Call LogError("SHR2" & Err.number & " D: " & Err.Description)

End Function

Private Function SHR4(ByVal bytValue As Byte) As Byte
On Error GoTo fallo
' Shift 8-bit value to right by 4 bits
' i.e. VB equivalent of "bytValue >> 4" in C
    SHR4 = bytValue \ &H10
  Exit Function
fallo:
Call LogError("SHR4 " & Err.number & " D: " & Err.Description)

End Function

Private Function SHR6(ByVal bytValue As Byte) As Byte
On Error GoTo fallo
' Shift 8-bit value to right by 6 bits
' i.e. VB equivalent of "bytValue >> 6" in C
    SHR6 = bytValue \ &H40
  Exit Function
fallo:
Call LogError("SHR6 " & Err.number & " D: " & Err.Description)

End Function

Public Function cv_BytesFromHex(ByVal sInputHex As String) As Variant
On Error GoTo fallo
' Returns array of bytes from hex string in big-endian order
' E.g. sHex="FEDC80" will return array {&HFE, &HDC, &H80}
    Dim i As Long
    Dim M As Long
    Dim aBytes() As Byte
    If Len(sInputHex) Mod 2 <> 0 Then
        sInputHex = "0" & sInputHex
    End If
    
    M = Len(sInputHex) \ 2
    If M <= 0 Then
        ' Version 2: Returns empty array
        cv_BytesFromHex = aBytes
        Exit Function
    End If
    
    ReDim aBytes(M - 1)
    
    For i = 0 To M - 1
        aBytes(i) = val("&H" & Mid$(sInputHex, i * 2 + 1, 2))
    Next
    
    cv_BytesFromHex = aBytes
  Exit Function
fallo:
Call LogError("CV BYTESFROMHEX " & Err.number & " D: " & Err.Description)

End Function

Public Function cv_WordsFromHex(ByVal sHex As String) As Variant
' Converts string <sHex> with hex values into array of words (long ints)
' E.g. "fedcba9876543210" will be converted into {&HFEDCBA98, &H76543210}
  On Error GoTo fallo
    Const ncLEN As Integer = 8
    Dim i As Long
    Dim nWords As Long
    Dim aWords() As Long
    
    nWords = Len(sHex) \ ncLEN
    If nWords <= 0 Then
        ' Version 2: Returns empty array
        cv_WordsFromHex = aWords
        Exit Function
    End If
    
    ReDim aWords(nWords - 1)
    For i = 0 To nWords - 1
        aWords(i) = val("&H" & Mid(sHex, i * ncLEN + 1, ncLEN))
    Next
    
    cv_WordsFromHex = aWords
    Exit Function
fallo:
Call LogError("CV WORDSFROMHEX" & Err.number & " D: " & Err.Description)
  
End Function

Public Function cv_HexFromWords(aWords) As String
' Converts array of words (Longs) into a hex string
' E.g. {&HFEDCBA98, &H76543210} will be converted to "FEDCBA9876543210"
    On Error GoTo fallo
    Const ncLEN As Integer = 8
    Dim i As Long
    Dim nWords As Long
    Dim sHex As String * ncLEN
    Dim iIndex As Long
    
    'Set up error handler to catch empty array
  
    If Not IsArray(aWords) Then
        Exit Function
    End If
    
    nWords = UBound(aWords) - LBound(aWords) + 1
    cv_HexFromWords = String(nWords * ncLEN, " ")
    iIndex = 0
    For i = 0 To nWords - 1
        sHex = Hex(aWords(i))
        sHex = String(ncLEN - Len(sHex), "0") & sHex
        Mid$(cv_HexFromWords, iIndex + 1, ncLEN) = sHex
        iIndex = iIndex + ncLEN
    Next
    

  Exit Function
fallo:
Call LogError("CV HEXFROMWORDS " & Err.number & " D: " & Err.Description)

End Function

Public Function cv_HexFromBytes(aBytes() As Byte) As String
On Error GoTo fallo
' Returns hex string from array of bytes
' E.g. aBytes() = {&HFE, &HDC, &H80} will return "FEDC80"
    Dim i As Long
    Dim iIndex As Long
    Dim nLen As Long
    
    'Set up error handler to catch empty array
  

    nLen = UBound(aBytes) - LBound(aBytes) + 1

    cv_HexFromBytes = String(nLen * 2, " ")
    iIndex = 0
    For i = LBound(aBytes) To UBound(aBytes)
        Mid$(cv_HexFromBytes, iIndex + 1, 2) = HexFromByte(aBytes(i))
        iIndex = iIndex + 2
    Next
    

Exit Function
fallo:
Call LogError("CV HEXFROMBYTES " & Err.number & " D: " & Err.Description)

End Function

Public Function cv_HexFromString(str As String) As String
On Error GoTo fallo
' Converts string <str> of ascii chars to string in hex format
' str may contain chars of any value between 0 and 255.
' E.g. "abc." will be converted to "6162632E"
    Dim byt As Byte
    Dim i As Long
    Dim n As Long
    Dim iIndex As Long
    Dim sHex As String
    
    n = Len(str)
    sHex = String(n * 2, " ")
    iIndex = 0
    For i = 1 To n
        byt = CByte(Asc(Mid$(str, i, 1)) And &HFF)
        Mid$(sHex, iIndex + 1, 2) = HexFromByte(byt)
        iIndex = iIndex + 2
    Next
    cv_HexFromString = sHex
 Exit Function
fallo:
Call LogError("CV HEXFROMSTRING" & Err.number & " D: " & Err.Description)
   
End Function

Public Function cv_StringFromHex(strHex As String) As String
On Error GoTo fallo
' Converts string <strHex> in hex format to string of ascii chars
' with value between 0 and 255.
' E.g. "6162632E" will be converted to "abc."
    Dim i As Integer
    Dim nBytes As Integer
    
    nBytes = Len(strHex) \ 2
    cv_StringFromHex = String(nBytes, " ")
    For i = 0 To nBytes - 1
        Mid$(cv_StringFromHex, i + 1, 1) = Chr$(val("&H" & Mid$(strHex, i * 2 + 1, 2)))
    Next
 Exit Function
fallo:
Call LogError("CV STRINGFROMHEX " & Err.number & " D: " & Err.Description)
   
End Function

Public Function cv_GetHexByte(ByVal sInputHex As String, iIndex As Long) As Byte
On Error GoTo fallo
' Extracts iIndex'th byte from hex string (starting at 1)
' E.g. cv_GetHexByte("fecdba98", 3) will return &HBA
    Dim i As Long
    i = 2 * iIndex
    If i > Len(sInputHex) Or i <= 0 Then
        cv_GetHexByte = 0
    Else
        cv_GetHexByte = val("&H" & Mid$(sInputHex, i - 1, 2))
    End If
 Exit Function
fallo:
Call LogError("CV GETHEXBYTE " & Err.number & " D: " & Err.Description)
   
End Function

Public Function RandHexByte() As String
On Error GoTo fallo
'   Returns a random byte as a 2-digit hex string
    Static stbInit As Boolean
    If Not stbInit Then
        Randomize
        stbInit = True
    End If
    
    RandHexByte = HexFromByte(CByte((Rnd * 256) And &HFF))
Exit Function
fallo:
Call LogError("RANDHEXBYTE " & Err.number & " D: " & Err.Description)


End Function

Public Function HexFromByte(ByVal X) As String
On Error GoTo fallo
' Returns a 2-digit hex string for byte x
    X = X And &HFF
    If X < 16 Then
        HexFromByte = "0" & Hex(X)
    Else
        HexFromByte = Hex(X)
    End If
Exit Function
fallo:
Call LogError("HEXFROMBYTE " & Err.number & " D: " & Err.Description)

End Function


Public Function testWordsHex()
  On Error GoTo fallo
    Dim aWords
    
    aWords = cv_WordsFromHex("FEDCBA9876543210")
    Debug.Print cv_HexFromWords(aWords)
 Exit Function
fallo:
Call LogError("TESTWORDSHEX " & Err.number & " D: " & Err.Description)
   
End Function

Public Function blf_BytesRaw(abData() As Byte, bEncrypt As Boolean) As Variant
    On Error GoTo fallo
    Dim nLen As Long
    Dim nBlocks As Long
    Dim iBlock As Long
    Dim j As Long
    Dim abOutput() As Byte
    Dim abBlock(7) As Byte
    Dim iIndex As Long
    
    ' Calc number of 8-byte blocks (ignore odd trailing bytes)
    nLen = UBound(abData) - LBound(abData) + 1
    nBlocks = nLen \ 8
    ReDim abOutput(nBlocks * 8 - 1)
    
    ' Work through in blocks of 8 bytes
    iIndex = 0
    For iBlock = 1 To nBlocks
        ' Get the next block of 8 bytes
        CopyMemory VarPtr(abBlock(0)), VarPtr(abData(iIndex)), 8&

        ' En/Decrypt the block according to flag
        If bEncrypt Then
            Call blf_EncryptBytes(abBlock())
        Else
            Call blf_DecryptBytes(abBlock())
        End If
        
        ' Copy to output string
        CopyMemory VarPtr(abOutput(iIndex)), VarPtr(abBlock(0)), 8&

        iIndex = iIndex + 8
    Next
    
    blf_BytesRaw = abOutput
 Exit Function
fallo:
'Call LogError("BLF BYTESRAW " & Err.Number & " D: " & Err.Description)
   
    
End Function

Public Function blf_BytesEnc(abData() As Byte) As Variant
 On Error GoTo fallo
    Dim abOutput() As Byte
    
    abOutput = PadBytes(abData)
    abOutput = blf_BytesRaw(abOutput, True)
    
    blf_BytesEnc = abOutput
Exit Function
fallo:
Call LogError("BLF BYTESENC" & Err.number & " D: " & Err.Description)

End Function

Public Function blf_BytesDec(abData() As Byte) As Variant
    On Error GoTo fallo
    Dim abOutput() As Byte
    
    abOutput = blf_BytesRaw(abData, False)
    abOutput = UnpadBytes(abOutput)
    
    blf_BytesDec = abOutput
Exit Function
fallo:
'Call LogError("BLF BYTESDEC" & Err.Number & " D: " & Err.Description)
    
End Function

Public Function PadBytes(abData() As Byte) As Variant

' Pad data bytes to next multiple of 8 bytes as per PKCS#5/RFC2630/RFC3370
    Dim nLen As Long
    Dim nPad As Integer
    Dim abPadded() As Byte
    Dim i As Long
    
    'Set up error handler for empty array
 On Error GoTo ArrayIsEmpty

    nLen = UBound(abData) - LBound(abData) + 1
    nPad = ((nLen \ 8) + 1) * 8 - nLen
    
    ReDim abPadded(nLen + nPad - 1)  ' Pad with # of pads (1-8)
    If nLen > 0 Then
        CopyMemory VarPtr(abPadded(0)), VarPtr(abData(0)), nLen
    End If
    For i = nLen To nLen + nPad - 1
        abPadded(i) = CByte(nPad)
    Next
    
ArrayIsEmpty:
    PadBytes = abPadded

End Function

Public Function UnpadBytes(abData() As Byte) As Variant
' Strip PKCS#5/RFC2630/RFC3370-style padding
 
    Dim nLen As Long
    Dim nPad As Long
    Dim abUnpadded() As Byte
    Dim i As Long
    
    'Set up error handler for empty array
On Error GoTo ArrayIsEmpty
    
    nLen = UBound(abData) - LBound(abData) + 1
    If nLen = 0 Then GoTo ArrayIsEmpty
    ' Get # of padding bytes from last char
    nPad = abData(nLen - 1)
    If nPad > 8 Then nPad = 0   ' In case invalid
    If nLen - nPad > 0 Then
        ReDim abUnpadded(nLen - nPad - 1)
        CopyMemory VarPtr(abUnpadded(0)), VarPtr(abData(0)), nLen - nPad
    End If

ArrayIsEmpty:
    UnpadBytes = abUnpadded
    
End Function

Public Function TestPadBytes()
    Dim abData() As Byte
    
    abData = StrConv("abc", vbFromUnicode)
    abData = PadBytes(abData)
    Stop
    abData = UnpadBytes(abData)
    Stop
    
End Function

Private Sub bXorBytes(aByt1() As Byte, aByt2() As Byte, nBytes As Long)
    Dim i As Long
    For i = 0 To nBytes - 1
        aByt1(i) = aByt1(i) Xor aByt2(i)
    Next
End Sub

Public Function blf_BytesEncRawCBC(abData() As Byte, abInitV() As Byte) As Variant
    Dim nLen As Long
    Dim nBlocks As Long
    Dim iBlock As Long
    Dim abBlock(7) As Byte
    Dim iIndex As Long
    Dim abReg(7) As Byte    ' Feedback register
    Dim abOutput() As Byte
    
    ' Initialisation vector should be a 8-byte array
    ' so ReDim just to make sure
    ' This will add zero bytes if too short or chop off any extra
    ReDim Preserve abInitV(7)
    
    ' Calc number of 8-byte blocks
    nLen = UBound(abData) - LBound(abData) + 1
    nBlocks = nLen \ 8
    
    ' Dimension output
    ReDim abOutput(nBlocks * 8 - 1)
    
    ' C_0 = IV
    CopyMemory VarPtr(abReg(0)), VarPtr(abInitV(0)), 8&
    
    ' Work through string in blocks of 8 bytes
    iIndex = 0
    For iBlock = 1 To nBlocks
        ' Fetch next block from input
        CopyMemory VarPtr(abBlock(0)), VarPtr(abData(iIndex)), 8&
        
        ' XOR with feedback register = Pi XOR C_i-1
        Call bXorBytes(abBlock, abReg, 8)
        
        ' Encrypt the block Ci = Ek(Pi XOR C_i-1)
        Call blf_EncryptBytes(abBlock())
        
        ' Store in feedback register Reg = Ci
        CopyMemory VarPtr(abReg(0)), VarPtr(abBlock(0)), 8&
        
        ' Copy to output string
        CopyMemory VarPtr(abOutput(iIndex)), VarPtr(abBlock(0)), 8&

        iIndex = iIndex + 8
    Next
    
    blf_BytesEncRawCBC = abOutput
    
End Function

Public Function blf_BytesDecRawCBC(abData() As Byte, abInitV() As Byte) As Variant
    Dim strIn As String
    Dim strOut As String
    Dim nLen As Long
    Dim nBlocks As Long
    Dim iBlock As Long
    Dim abBlock(7) As Byte
    Dim iIndex As Long
    Dim abReg(7) As Byte    ' Feedback register
    Dim abStore(7) As Byte
    Dim abOutput() As Byte
    
    ' Initialisation vector should be a 8-byte array
    ' so ReDim just to make sure
    ' This will add zero bytes if too short or chop off any extra
    ReDim Preserve abInitV(7)
    
    ' Calc number of 8-byte blocks
    nLen = UBound(abData) - LBound(abData) + 1
    nBlocks = nLen \ 8
    
    ' Dimension output
    ReDim abOutput(nBlocks * 8 - 1)
    
    ' C_0 = IV
    CopyMemory VarPtr(abReg(0)), VarPtr(abInitV(0)), 8&
    
    ' Work through string in blocks of 8 bytes
    iIndex = 0
    For iBlock = 1 To nBlocks
        ' Fetch next block from input
        CopyMemory VarPtr(abBlock(0)), VarPtr(abData(iIndex)), 8&
        
        ' Save C_i-1
        CopyMemory VarPtr(abStore(0)), VarPtr(abBlock(0)), 8&
        
        ' Decrypt the block Dk(Ci)
        Call blf_DecryptBytes(abBlock())
        
        ' XOR with feedback register = C_i-1 XOR Dk(Ci)
        Call bXorBytes(abBlock, abReg, 8)
        
        ' Store in feedback register Reg = C_i-1
        CopyMemory VarPtr(abReg(0)), VarPtr(abStore(0)), 8&
        
        ' Copy to output string
        CopyMemory VarPtr(abOutput(iIndex)), VarPtr(abBlock(0)), 8&

        iIndex = iIndex + 8
    Next
    
    blf_BytesDecRawCBC = abOutput
    
End Function

Public Function blf_BytesEncCBC(abData() As Byte, abInitV() As Byte) As Variant
    Dim abOutput() As Byte
    
    abOutput = PadBytes(abData)
    abOutput = blf_BytesEncRawCBC(abOutput, abInitV)
    
    blf_BytesEncCBC = abOutput
End Function

Public Function blf_BytesDecCBC(abData() As Byte, abInitV() As Byte) As Variant
    Dim abOutput() As Byte
    
    abOutput = blf_BytesDecRawCBC(abData, abInitV)
    abOutput = UnpadBytes(abOutput)
    
    blf_BytesDecCBC = abOutput
End Function

Private Function blf_F(X As Long) As Long
    Dim a As Byte, b As Byte, C As Byte, d As Byte
    Dim Y As Long
    
    Call uwSplit(X, a, b, C, d)
    
    Y = uw_WordAdd(blf_S(0, a), blf_S(1, b))
    Y = Y Xor blf_S(2, C)
    Y = uw_WordAdd(Y, blf_S(3, d))
    blf_F = Y
    
End Function

Public Function blf_EncipherBlock(xL As Long, xR As Long)
    Dim i As Integer
    Dim temp As Long
    
    For i = 0 To ncROUNDS - 1
        xL = xL Xor blf_P(i)
        xR = blf_F(xL) Xor xR
        temp = xL
        xL = xR
        xR = temp
    Next
    
    temp = xL
    xL = xR
    xR = temp
    
    xR = xR Xor blf_P(ncROUNDS)
    xL = xL Xor blf_P(ncROUNDS + 1)
        
End Function

Public Function blf_DecipherBlock(xL As Long, xR As Long)
    Dim i As Integer
    Dim temp As Long
    
    For i = ncROUNDS + 1 To 2 Step -1
        xL = xL Xor blf_P(i)
        xR = blf_F(xL) Xor xR
        temp = xL
        xL = xR
        xR = temp
    Next
    
    temp = xL
    xL = xR
    xR = temp
    
    xR = xR Xor blf_P(1)
    xL = xL Xor blf_P(0)
        
End Function

Public Function blf_Initialise(aKey() As Byte, nKeyBytes As Integer)
    Dim i As Integer, j As Integer, k As Integer
    Dim wData As Long, wDataL As Long, wDataR As Long
    
    Call blf_LoadArrays     ' Initialise P and S arrays

    j = 0
    For i = 0 To (ncROUNDS + 2 - 1)
        wData = &H0
        For k = 0 To 3
            wData = uw_ShiftLeftBy8(wData) Or aKey(j)
            j = j + 1
            If j >= nKeyBytes Then j = 0
        Next k
        blf_P(i) = blf_P(i) Xor wData
    Next i
    
    wDataL = &H0
    wDataR = &H0
    
    For i = 0 To (ncROUNDS + 2 - 1) Step 2
        Call blf_EncipherBlock(wDataL, wDataR)
        
        blf_P(i) = wDataL
        blf_P(i + 1) = wDataR
    Next i
    
    For i = 0 To 3
        For j = 0 To 255 Step 2
            Call blf_EncipherBlock(wDataL, wDataR)
    
            blf_S(i, j) = wDataL
            blf_S(i, j + 1) = wDataR
        Next j
    Next i

End Function

Public Function blf_Key(aKey() As Byte, nKeyLen As Integer) As Boolean
    blf_Key = False
    If nKeyLen < 0 Or nKeyLen > ncMAXKEYLEN Then
        Exit Function
    End If
    
    Call blf_Initialise(aKey, nKeyLen)
    
    blf_Key = True
End Function

Public Function blf_KeyInit(aKey() As Byte) As Boolean
' Added Version 5: Replacement for blf_Key to avoid specifying keylen
' Version 6: Added error checking for input
    Dim nKeyLen As Integer
    
    blf_KeyInit = False
    
    'Set up error handler to catch empty array
    On Error GoTo ArrayIsEmpty

    nKeyLen = UBound(aKey) - LBound(aKey) + 1
    If nKeyLen < 0 Or nKeyLen > ncMAXKEYLEN Then
        Exit Function
    End If
    
    Call blf_Initialise(aKey, nKeyLen)
    
    blf_KeyInit = True
    
ArrayIsEmpty:

End Function

Public Function blf_EncryptBytes(aBytes() As Byte)
' aBytes() must be 8 bytes long
' Revised Version 5: January 2002. To use faster uwJoin and uwSplit fns.
    Dim wordL As Long, wordR As Long
    
    ' Convert to 2 x words
    wordL = uwJoin(aBytes(0), aBytes(1), aBytes(2), aBytes(3))
    wordR = uwJoin(aBytes(4), aBytes(5), aBytes(6), aBytes(7))
    ' Encrypt it
    Call blf_EncipherBlock(wordL, wordR)
    ' Put back into bytes
    Call uwSplit(wordL, aBytes(0), aBytes(1), aBytes(2), aBytes(3))
    Call uwSplit(wordR, aBytes(4), aBytes(5), aBytes(6), aBytes(7))
End Function

Public Function blf_DecryptBytes(aBytes() As Byte)
' aBytes() must be 8 bytes long
' Revised Version 5:: January 2002. To use faster uwJoin and uwSplit fns.
    Dim wordL As Long, wordR As Long
    
    ' Convert to 2 x words
    wordL = uwJoin(aBytes(0), aBytes(1), aBytes(2), aBytes(3))
    wordR = uwJoin(aBytes(4), aBytes(5), aBytes(6), aBytes(7))
    ' Decrypt it
    Call blf_DecipherBlock(wordL, wordR)
    ' Put back into bytes
    Call uwSplit(wordL, aBytes(0), aBytes(1), aBytes(2), aBytes(3))
    Call uwSplit(wordR, aBytes(4), aBytes(5), aBytes(6), aBytes(7))

End Function

Public Function blf_LoadArrays()
' Use Array fn and a temp variant array to load data into arrays
    Dim vntA As Variant
    Dim i As Integer

    ' P-array
    vntA = Array( _
        &H243F6A88, &H85A308D3, &H13198A2E, &H3707344, _
        &HA4093822, &H299F31D0, &H82EFA98, &HEC4E6C89, _
        &H452821E6, &H38D01377, &HBE5466CF, &H34E90C6C, _
        &HC0AC29B7, &HC97C50DD, &H3F84D5B5, &HB5470917, _
        &H9216D5D9, &H8979FB1B)
    
    For i = 0 To 17
        blf_P(i) = vntA(i)
    Next
    
    ' Load S-boxes - 16 x 4 at a time
    
    ' S-box[0]
    vntA = Array( _
        &HD1310BA6, &H98DFB5AC, &H2FFD72DB, &HD01ADFB7, _
        &HB8E1AFED, &H6A267E96, &HBA7C9045, &HF12C7F99, _
        &H24A19947, &HB3916CF7, &H801F2E2, &H858EFC16, _
        &H636920D8, &H71574E69, &HA458FEA3, &HF4933D7E, _
        &HD95748F, &H728EB658, &H718BCD58, &H82154AEE, _
        &H7B54A41D, &HC25A59B5, &H9C30D539, &H2AF26013, _
        &HC5D1B023, &H286085F0, &HCA417918, &HB8DB38EF, _
        &H8E79DCB0, &H603A180E, &H6C9E0E8B, &HB01E8A3E, _
        &HD71577C1, &HBD314B27, &H78AF2FDA, &H55605C60, _
        &HE65525F3, &HAA55AB94, &H57489862, &H63E81440, _
        &H55CA396A, &H2AAB10B6, &HB4CC5C34, &H1141E8CE, _
        &HA15486AF, &H7C72E993, &HB3EE1411, &H636FBC2A, _
        &H2BA9C55D, &H741831F6, &HCE5C3E16, &H9B87931E, _
        &HAFD6BA33, &H6C24CF5C, &H7A325381, &H28958677, _
        &H3B8F4898, &H6B4BB9AF, &HC4BFE81B, &H66282193, _
        &H61D809CC, &HFB21A991, &H487CAC60, &H5DEC8032)
        
    For i = 0 To 63
        blf_S(0, i) = vntA(i)
    Next
    
    vntA = Array( _
        &HEF845D5D, &HE98575B1, &HDC262302, &HEB651B88, _
        &H23893E81, &HD396ACC5, &HF6D6FF3, &H83F44239, _
        &H2E0B4482, &HA4842004, &H69C8F04A, &H9E1F9B5E, _
        &H21C66842, &HF6E96C9A, &H670C9C61, &HABD388F0, _
        &H6A51A0D2, &HD8542F68, &H960FA728, &HAB5133A3, _
        &H6EEF0B6C, &H137A3BE4, &HBA3BF050, &H7EFB2A98, _
        &HA1F1651D, &H39AF0176, &H66CA593E, &H82430E88, _
        &H8CEE8619, &H456F9FB4, &H7D84A5C3, &H3B8B5EBE, _
        &HE06F75D8, &H85C12073, &H401A449F, &H56C16AA6, _
        &H4ED3AA62, &H363F7706, &H1BFEDF72, &H429B023D, _
        &H37D0D724, &HD00A1248, &HDB0FEAD3, &H49F1C09B, _
        &H75372C9, &H80991B7B, &H25D479D8, &HF6E8DEF7, _
        &HE3FE501A, &HB6794C3B, &H976CE0BD, &H4C006BA, _
        &HC1A94FB6, &H409F60C4, &H5E5C9EC2, &H196A2463, _
        &H68FB6FAF, &H3E6C53B5, &H1339B2EB, &H3B52EC6F, _
        &H6DFC511F, &H9B30952C, &HCC814544, &HAF5EBD09)
        
    For i = 0 To 63     '64 To 127
        blf_S(0, i + 64) = vntA(i)
    Next
        
    vntA = Array( _
        &HBEE3D004, &HDE334AFD, &H660F2807, &H192E4BB3, _
        &HC0CBA857, &H45C8740F, &HD20B5F39, &HB9D3FBDB, _
        &H5579C0BD, &H1A60320A, &HD6A100C6, &H402C7279, _
        &H679F25FE, &HFB1FA3CC, &H8EA5E9F8, &HDB3222F8, _
        &H3C7516DF, &HFD616B15, &H2F501EC8, &HAD0552AB, _
        &H323DB5FA, &HFD238760, &H53317B48, &H3E00DF82, _
        &H9E5C57BB, &HCA6F8CA0, &H1A87562E, &HDF1769DB, _
        &HD542A8F6, &H287EFFC3, &HAC6732C6, &H8C4F5573, _
        &H695B27B0, &HBBCA58C8, &HE1FFA35D, &HB8F011A0, _
        &H10FA3D98, &HFD2183B8, &H4AFCB56C, &H2DD1D35B, _
        &H9A53E479, &HB6F84565, &HD28E49BC, &H4BFB9790, _
        &HE1DDF2DA, &HA4CB7E33, &H62FB1341, &HCEE4C6E8, _
        &HEF20CADA, &H36774C01, &HD07E9EFE, &H2BF11FB4, _
        &H95DBDA4D, &HAE909198, &HEAAD8E71, &H6B93D5A0, _
        &HD08ED1D0, &HAFC725E0, &H8E3C5B2F, &H8E7594B7, _
        &H8FF6E2FB, &HF2122B64, &H8888B812, &H900DF01C)
        
    For i = 0 To 63     ' 128 To 191
        blf_S(0, i + 128) = vntA(i)
    Next
        
    vntA = Array( _
        &H4FAD5EA0, &H688FC31C, &HD1CFF191, &HB3A8C1AD, _
        &H2F2F2218, &HBE0E1777, &HEA752DFE, &H8B021FA1, _
        &HE5A0CC0F, &HB56F74E8, &H18ACF3D6, &HCE89E299, _
        &HB4A84FE0, &HFD13E0B7, &H7CC43B81, &HD2ADA8D9, _
        &H165FA266, &H80957705, &H93CC7314, &H211A1477, _
        &HE6AD2065, &H77B5FA86, &HC75442F5, &HFB9D35CF, _
        &HEBCDAF0C, &H7B3E89A0, &HD6411BD3, &HAE1E7E49, _
        &H250E2D, &H2071B35E, &H226800BB, &H57B8E0AF, _
        &H2464369B, &HF009B91E, &H5563911D, &H59DFA6AA, _
        &H78C14389, &HD95A537F, &H207D5BA2, &H2E5B9C5, _
        &H83260376, &H6295CFA9, &H11C81968, &H4E734A41, _
        &HB3472DCA, &H7B14A94A, &H1B510052, &H9A532915, _
        &HD60F573F, &HBC9BC6E4, &H2B60A476, &H81E67400, _
        &H8BA6FB5, &H571BE91F, &HF296EC6B, &H2A0DD915, _
        &HB6636521, &HE7B9F9B6, &HFF34052E, &HC5855664, _
        &H53B02D5D, &HA99F8FA1, &H8BA4799, &H6E85076A)

    For i = 0 To 63     ' 192 To 255
        blf_S(0, i + 192) = vntA(i)
    Next
   
    ' S-box[1]
    vntA = Array( _
        &H4B7A70E9, &HB5B32944, &HDB75092E, &HC4192623, _
        &HAD6EA6B0, &H49A7DF7D, &H9CEE60B8, &H8FEDB266, _
        &HECAA8C71, &H699A17FF, &H5664526C, &HC2B19EE1, _
        &H193602A5, &H75094C29, &HA0591340, &HE4183A3E, _
        &H3F54989A, &H5B429D65, &H6B8FE4D6, &H99F73FD6, _
        &HA1D29C07, &HEFE830F5, &H4D2D38E6, &HF0255DC1, _
        &H4CDD2086, &H8470EB26, &H6382E9C6, &H21ECC5E, _
        &H9686B3F, &H3EBAEFC9, &H3C971814, &H6B6A70A1, _
        &H687F3584, &H52A0E286, &HB79C5305, &HAA500737, _
        &H3E07841C, &H7FDEAE5C, &H8E7D44EC, &H5716F2B8, _
        &HB03ADA37, &HF0500C0D, &HF01C1F04, &H200B3FF, _
        &HAE0CF51A, &H3CB574B2, &H25837A58, &HDC0921BD, _
        &HD19113F9, &H7CA92FF6, &H94324773, &H22F54701, _
        &H3AE5E581, &H37C2DADC, &HC8B57634, &H9AF3DDA7, _
        &HA9446146, &HFD0030E, &HECC8C73E, &HA4751E41, _
        &HE238CD99, &H3BEA0E2F, &H3280BBA1, &H183EB331)
        
    For i = 0 To 63
        blf_S(1, i) = vntA(i)
    Next
        
    vntA = Array( _
        &H4E548B38, &H4F6DB908, &H6F420D03, &HF60A04BF, _
        &H2CB81290, &H24977C79, &H5679B072, &HBCAF89AF, _
        &HDE9A771F, &HD9930810, &HB38BAE12, &HDCCF3F2E, _
        &H5512721F, &H2E6B7124, &H501ADDE6, &H9F84CD87, _
        &H7A584718, &H7408DA17, &HBC9F9ABC, &HE94B7D8C, _
        &HEC7AEC3A, &HDB851DFA, &H63094366, &HC464C3D2, _
        &HEF1C1847, &H3215D908, &HDD433B37, &H24C2BA16, _
        &H12A14D43, &H2A65C451, &H50940002, &H133AE4DD, _
        &H71DFF89E, &H10314E55, &H81AC77D6, &H5F11199B, _
        &H43556F1, &HD7A3C76B, &H3C11183B, &H5924A509, _
        &HF28FE6ED, &H97F1FBFA, &H9EBABF2C, &H1E153C6E, _
        &H86E34570, &HEAE96FB1, &H860E5E0A, &H5A3E2AB3, _
        &H771FE71C, &H4E3D06FA, &H2965DCB9, &H99E71D0F, _
        &H803E89D6, &H5266C825, &H2E4CC978, &H9C10B36A, _
        &HC6150EBA, &H94E2EA78, &HA5FC3C53, &H1E0A2DF4, _
        &HF2F74EA7, &H361D2B3D, &H1939260F, &H19C27960)
        
    For i = 0 To 63     '64 To 127
        blf_S(1, i + 64) = vntA(i)
    Next
        
    vntA = Array( _
        &H5223A708, &HF71312B6, &HEBADFE6E, &HEAC31F66, _
        &HE3BC4595, &HA67BC883, &HB17F37D1, &H18CFF28, _
        &HC332DDEF, &HBE6C5AA5, &H65582185, &H68AB9802, _
        &HEECEA50F, &HDB2F953B, &H2AEF7DAD, &H5B6E2F84, _
        &H1521B628, &H29076170, &HECDD4775, &H619F1510, _
        &H13CCA830, &HEB61BD96, &H334FE1E, &HAA0363CF, _
        &HB5735C90, &H4C70A239, &HD59E9E0B, &HCBAADE14, _
        &HEECC86BC, &H60622CA7, &H9CAB5CAB, &HB2F3846E, _
        &H648B1EAF, &H19BDF0CA, &HA02369B9, &H655ABB50, _
        &H40685A32, &H3C2AB4B3, &H319EE9D5, &HC021B8F7, _
        &H9B540B19, &H875FA099, &H95F7997E, &H623D7DA8, _
        &HF837889A, &H97E32D77, &H11ED935F, &H16681281, _
        &HE358829, &HC7E61FD6, &H96DEDFA1, &H7858BA99, _
        &H57F584A5, &H1B227263, &H9B83C3FF, &H1AC24696, _
        &HCDB30AEB, &H532E3054, &H8FD948E4, &H6DBC3128, _
        &H58EBF2EF, &H34C6FFEA, &HFE28ED61, &HEE7C3C73)
        
    For i = 0 To 63     ' 128 To 191
        blf_S(1, i + 128) = vntA(i)
    Next
        
    vntA = Array( _
        &H5D4A14D9, &HE864B7E3, &H42105D14, &H203E13E0, _
        &H45EEE2B6, &HA3AAABEA, &HDB6C4F15, &HFACB4FD0, _
        &HC742F442, &HEF6ABBB5, &H654F3B1D, &H41CD2105, _
        &HD81E799E, &H86854DC7, &HE44B476A, &H3D816250, _
        &HCF62A1F2, &H5B8D2646, &HFC8883A0, &HC1C7B6A3, _
        &H7F1524C3, &H69CB7492, &H47848A0B, &H5692B285, _
        &H95BBF00, &HAD19489D, &H1462B174, &H23820E00, _
        &H58428D2A, &HC55F5EA, &H1DADF43E, &H233F7061, _
        &H3372F092, &H8D937E41, &HD65FECF1, &H6C223BDB, _
        &H7CDE3759, &HCBEE7460, &H4085F2A7, &HCE77326E, _
        &HA6078084, &H19F8509E, &HE8EFD855, &H61D99735, _
        &HA969A7AA, &HC50C06C2, &H5A04ABFC, &H800BCADC, _
        &H9E447A2E, &HC3453484, &HFDD56705, &HE1E9EC9, _
        &HDB73DBD3, &H105588CD, &H675FDA79, &HE3674340, _
        &HC5C43465, &H713E38D8, &H3D28F89E, &HF16DFF20, _
        &H153E21E7, &H8FB03D4A, &HE6E39F2B, &HDB83ADF7)
        
    For i = 0 To 63     ' 192 To 255
        blf_S(1, i + 192) = vntA(i)
    Next
    
    ' S-box[2]
    vntA = Array( _
        &HE93D5A68, &H948140F7, &HF64C261C, &H94692934, _
        &H411520F7, &H7602D4F7, &HBCF46B2E, &HD4A20068, _
        &HD4082471, &H3320F46A, &H43B7D4B7, &H500061AF, _
        &H1E39F62E, &H97244546, &H14214F74, &HBF8B8840, _
        &H4D95FC1D, &H96B591AF, &H70F4DDD3, &H66A02F45, _
        &HBFBC09EC, &H3BD9785, &H7FAC6DD0, &H31CB8504, _
        &H96EB27B3, &H55FD3941, &HDA2547E6, &HABCA0A9A, _
        &H28507825, &H530429F4, &HA2C86DA, &HE9B66DFB, _
        &H68DC1462, &HD7486900, &H680EC0A4, &H27A18DEE, _
        &H4F3FFEA2, &HE887AD8C, &HB58CE006, &H7AF4D6B6, _
        &HAACE1E7C, &HD3375FEC, &HCE78A399, &H406B2A42, _
        &H20FE9E35, &HD9F385B9, &HEE39D7AB, &H3B124E8B, _
        &H1DC9FAF7, &H4B6D1856, &H26A36631, &HEAE397B2, _
        &H3A6EFA74, &HDD5B4332, &H6841E7F7, &HCA7820FB, _
        &HFB0AF54E, &HD8FEB397, &H454056AC, &HBA489527, _
        &H55533A3A, &H20838D87, &HFE6BA9B7, &HD096954B)
        
    For i = 0 To 63
        blf_S(2, i) = vntA(i)
    Next
        
    vntA = Array( _
        &H55A867BC, &HA1159A58, &HCCA92963, &H99E1DB33, _
        &HA62A4A56, &H3F3125F9, &H5EF47E1C, &H9029317C, _
        &HFDF8E802, &H4272F70, &H80BB155C, &H5282CE3, _
        &H95C11548, &HE4C66D22, &H48C1133F, &HC70F86DC, _
        &H7F9C9EE, &H41041F0F, &H404779A4, &H5D886E17, _
        &H325F51EB, &HD59BC0D1, &HF2BCC18F, &H41113564, _
        &H257B7834, &H602A9C60, &HDFF8E8A3, &H1F636C1B, _
        &HE12B4C2, &H2E1329E, &HAF664FD1, &HCAD18115, _
        &H6B2395E0, &H333E92E1, &H3B240B62, &HEEBEB922, _
        &H85B2A20E, &HE6BA0D99, &HDE720C8C, &H2DA2F728, _
        &HD0127845, &H95B794FD, &H647D0862, &HE7CCF5F0, _
        &H5449A36F, &H877D48FA, &HC39DFD27, &HF33E8D1E, _
        &HA476341, &H992EFF74, &H3A6F6EAB, &HF4F8FD37, _
        &HA812DC60, &HA1EBDDF8, &H991BE14C, &HDB6E6B0D, _
        &HC67B5510, &H6D672C37, &H2765D43B, &HDCD0E804, _
        &HF1290DC7, &HCC00FFA3, &HB5390F92, &H690FED0B)
        
    For i = 0 To 63     '64 To 127
        blf_S(2, i + 64) = vntA(i)
    Next
        
    vntA = Array( _
        &H667B9FFB, &HCEDB7D9C, &HA091CF0B, &HD9155EA3, _
        &HBB132F88, &H515BAD24, &H7B9479BF, &H763BD6EB, _
        &H37392EB3, &HCC115979, &H8026E297, &HF42E312D, _
        &H6842ADA7, &HC66A2B3B, &H12754CCC, &H782EF11C, _
        &H6A124237, &HB79251E7, &H6A1BBE6, &H4BFB6350, _
        &H1A6B1018, &H11CAEDFA, &H3D25BDD8, &HE2E1C3C9, _
        &H44421659, &HA121386, &HD90CEC6E, &HD5ABEA2A, _
        &H64AF674E, &HDA86A85F, &HBEBFE988, &H64E4C3FE, _
        &H9DBC8057, &HF0F7C086, &H60787BF8, &H6003604D, _
        &HD1FD8346, &HF6381FB0, &H7745AE04, &HD736FCCC, _
        &H83426B33, &HF01EAB71, &HB0804187, &H3C005E5F, _
        &H77A057BE, &HBDE8AE24, &H55464299, &HBF582E61, _
        &H4E58F48F, &HF2DDFDA2, &HF474EF38, &H8789BDC2, _
        &H5366F9C3, &HC8B38E74, &HB475F255, &H46FCD9B9, _
        &H7AEB2661, &H8B1DDF84, &H846A0E79, &H915F95E2, _
        &H466E598E, &H20B45770, &H8CD55591, &HC902DE4C)
        
    For i = 0 To 63     ' 128 To 191
        blf_S(2, i + 128) = vntA(i)
    Next
        
    vntA = Array( _
        &HB90BACE1, &HBB8205D0, &H11A86248, &H7574A99E, _
        &HB77F19B6, &HE0A9DC09, &H662D09A1, &HC4324633, _
        &HE85A1F02, &H9F0BE8C, &H4A99A025, &H1D6EFE10, _
        &H1AB93D1D, &HBA5A4DF, &HA186F20F, &H2868F169, _
        &HDCB7DA83, &H573906FE, &HA1E2CE9B, &H4FCD7F52, _
        &H50115E01, &HA70683FA, &HA002B5C4, &HDE6D027, _
        &H9AF88C27, &H773F8641, &HC3604C06, &H61A806B5, _
        &HF0177A28, &HC0F586E0, &H6058AA, &H30DC7D62, _
        &H11E69ED7, &H2338EA63, &H53C2DD94, &HC2C21634, _
        &HBBCBEE56, &H90BCB6DE, &HEBFC7DA1, &HCE591D76, _
        &H6F05E409, &H4B7C0188, &H39720A3D, &H7C927C24, _
        &H86E3725F, &H724D9DB9, &H1AC15BB4, &HD39EB8FC, _
        &HED545578, &H8FCA5B5, &HD83D7CD3, &H4DAD0FC4, _
        &H1E50EF5E, &HB161E6F8, &HA28514D9, &H6C51133C, _
        &H6FD5C7E7, &H56E14EC4, &H362ABFCE, &HDDC6C837, _
        &HD79A3234, &H92638212, &H670EFA8E, &H406000E0)
        
    For i = 0 To 63     ' 192 To 255
        blf_S(2, i + 192) = vntA(i)
    Next
    
    ' S-box[3]
    vntA = Array( _
        &H3A39CE37, &HD3FAF5CF, &HABC27737, &H5AC52D1B, _
        &H5CB0679E, &H4FA33742, &HD3822740, &H99BC9BBE, _
        &HD5118E9D, &HBF0F7315, &HD62D1C7E, &HC700C47B, _
        &HB78C1B6B, &H21A19045, &HB26EB1BE, &H6A366EB4, _
        &H5748AB2F, &HBC946E79, &HC6A376D2, &H6549C2C8, _
        &H530FF8EE, &H468DDE7D, &HD5730A1D, &H4CD04DC6, _
        &H2939BBDB, &HA9BA4650, &HAC9526E8, &HBE5EE304, _
        &HA1FAD5F0, &H6A2D519A, &H63EF8CE2, &H9A86EE22, _
        &HC089C2B8, &H43242EF6, &HA51E03AA, &H9CF2D0A4, _
        &H83C061BA, &H9BE96A4D, &H8FE51550, &HBA645BD6, _
        &H2826A2F9, &HA73A3AE1, &H4BA99586, &HEF5562E9, _
        &HC72FEFD3, &HF752F7DA, &H3F046F69, &H77FA0A59, _
        &H80E4A915, &H87B08601, &H9B09E6AD, &H3B3EE593, _
        &HE990FD5A, &H9E34D797, &H2CF0B7D9, &H22B8B51, _
        &H96D5AC3A, &H17DA67D, &HD1CF3ED6, &H7C7D2D28, _
        &H1F9F25CF, &HADF2B89B, &H5AD6B472, &H5A88F54C)
        
    For i = 0 To 63
        blf_S(3, i) = vntA(i)
    Next
        
    vntA = Array( _
        &HE029AC71, &HE019A5E6, &H47B0ACFD, &HED93FA9B, _
        &HE8D3C48D, &H283B57CC, &HF8D56629, &H79132E28, _
        &H785F0191, &HED756055, &HF7960E44, &HE3D35E8C, _
        &H15056DD4, &H88F46DBA, &H3A16125, &H564F0BD, _
        &HC3EB9E15, &H3C9057A2, &H97271AEC, &HA93A072A, _
        &H1B3F6D9B, &H1E6321F5, &HF59C66FB, &H26DCF319, _
        &H7533D928, &HB155FDF5, &H3563482, &H8ABA3CBB, _
        &H28517711, &HC20AD9F8, &HABCC5167, &HCCAD925F, _
        &H4DE81751, &H3830DC8E, &H379D5862, &H9320F991, _
        &HEA7A90C2, &HFB3E7BCE, &H5121CE64, &H774FBE32, _
        &HA8B6E37E, &HC3293D46, &H48DE5369, &H6413E680, _
        &HA2AE0810, &HDD6DB224, &H69852DFD, &H9072166, _
        &HB39A460A, &H6445C0DD, &H586CDECF, &H1C20C8AE, _
        &H5BBEF7DD, &H1B588D40, &HCCD2017F, &H6BB4E3BB, _
        &HDDA26A7E, &H3A59FF45, &H3E350A44, &HBCB4CDD5, _
        &H72EACEA8, &HFA6484BB, &H8D6612AE, &HBF3C6F47)
        
    For i = 0 To 63     '64 To 127
        blf_S(3, i + 64) = vntA(i)
    Next
        
    vntA = Array( _
        &HD29BE463, &H542F5D9E, &HAEC2771B, &HF64E6370, _
        &H740E0D8D, &HE75B1357, &HF8721671, &HAF537D5D, _
        &H4040CB08, &H4EB4E2CC, &H34D2466A, &H115AF84, _
        &HE1B00428, &H95983A1D, &H6B89FB4, &HCE6EA048, _
        &H6F3F3B82, &H3520AB82, &H11A1D4B, &H277227F8, _
        &H611560B1, &HE7933FDC, &HBB3A792B, &H344525BD, _
        &HA08839E1, &H51CE794B, &H2F32C9B7, &HA01FBAC9, _
        &HE01CC87E, &HBCC7D1F6, &HCF0111C3, &HA1E8AAC7, _
        &H1A908749, &HD44FBD9A, &HD0DADECB, &HD50ADA38, _
        &H339C32A, &HC6913667, &H8DF9317C, &HE0B12B4F, _
        &HF79E59B7, &H43F5BB3A, &HF2D519FF, &H27D9459C, _
        &HBF97222C, &H15E6FC2A, &HF91FC71, &H9B941525, _
        &HFAE59361, &HCEB69CEB, &HC2A86459, &H12BAA8D1, _
        &HB6C1075E, &HE3056A0C, &H10D25065, &HCB03A442, _
        &HE0EC6E0E, &H1698DB3B, &H4C98A0BE, &H3278E964, _
        &H9F1F9532, &HE0D392DF, &HD3A0342B, &H8971F21E)
        
    For i = 0 To 63     ' 128 To 191
        blf_S(3, i + 128) = vntA(i)
    Next
        
    vntA = Array( _
        &H1B0A7441, &H4BA3348C, &HC5BE7120, &HC37632D8, _
        &HDF359F8D, &H9B992F2E, &HE60B6F47, &HFE3F11D, _
        &HE54CDA54, &H1EDAD891, &HCE6279CF, &HCD3E7E6F, _
        &H1618B166, &HFD2C1D05, &H848FD2C5, &HF6FB2299, _
        &HF523F357, &HA6327623, &H93A83531, &H56CCCD02, _
        &HACF08162, &H5A75EBB5, &H6E163697, &H88D273CC, _
        &HDE966292, &H81B949D0, &H4C50901B, &H71C65614, _
        &HE6C6C7BD, &H327A140A, &H45E1D006, &HC3F27B9A, _
        &HC9AA53FD, &H62A80F00, &HBB25BFE2, &H35BDD2F6, _
        &H71126905, &HB2040222, &HB6CBCF7C, &HCD769C2B, _
        &H53113EC0, &H1640E3D3, &H38ABBD60, &H2547ADF0, _
        &HBA38209C, &HF746CE76, &H77AFA1C5, &H20756060, _
        &H85CBFE4E, &H8AE88DD8, &H7AAAF9B0, &H4CF9AA7E, _
        &H1948C25C, &H2FB8A8C, &H1C36AE4, &HD6EBE1F9, _
        &H90D4F869, &HA65CDEA0, &H3F09252D, &HC208E69F, _
        &HB74E6132, &HCE77E25B, &H578FDFE3, &H3AC372E6)

    For i = 0 To 63     ' 192 To 255
        blf_S(3, i + 192) = vntA(i)
    Next
    
    ' DEBUG: Check for zeroes
    Dim j As Integer
    For i = 0 To 3
        For j = 0 To 255
            If blf_S(i, j) = 0 Then
                'MsgBox "Zero value in S" & i & "," & j & ")"
            End If
        Next
    Next
    
End Function

