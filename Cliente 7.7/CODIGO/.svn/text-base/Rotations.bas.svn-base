Attribute VB_Name = "Rotations"
Public Function BigShiftLeft(value1 As String, shifts As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    shifts = shifts Mod 32
    
    If shifts = 0 Then
        BigShiftLeft = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    
    For i = 1 To shifts
        For j = 1 To 32
            Mid(value1, j, 1) = Mid(value1, j + 1, 1)
        Next j
        If Not Mid(value1, 32, 1) = "0" Then Mid(value1, 32, 1) = "0"
    Next i
    tempstr = value1

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    BigShiftLeft = Right(value1, 8)
End Function
Public Function BigShiftRight(value1 As String, shifts As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    shifts = shifts Mod 32
    
    If shifts = 0 Then
        BigShiftRight = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    
    For i = 1 To shifts
        For j = 32 To 2 Step -1
            Mid(value1, j, 1) = Mid(value1, j - 1, 1)
        Next j
        If Not Mid(value1, 1, 1) = "0" Then Mid(value1, 1, 1) = "0"
    Next i
    tempstr = value1

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    BigShiftRight = Right(value1, 8)
End Function


Public Function RotLeft(ByVal value1 As String, ByVal rots As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    rots = rots Mod 32
    
    If rots = 0 Then
        RotLeft = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    tempstr = Mid$(value1, rots + 1) + Left$(value1, rots)

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    RotLeft = Right(value1, 8)
End Function
Public Function RotRight(ByVal value1 As String, ByVal rots As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    rots = rots Mod 32
    
    If rots = 0 Then
        RotRight = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    tempstr = Right$(value1, rots) + Mid$(value1, 1, Len(value1) - rots)

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    RotRight = Right(value1, 8)
End Function
Public Function BigAdd(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value1 = Space$(Abs(tempnum)) + value1
    ElseIf tempnum > 0 Then
        value2 = Space$(Abs(tempnum)) + value2
    End If

    tempnum = 0
    For loopit = Len(value1) To 1 Step -1
        tempnum = tempnum + Val("&H" + Mid$(value1, loopit, 1)) + Val("&H" + Mid$(value2, loopit, 1))
        valueans = Hex$(tempnum Mod 16) + valueans
        tempnum = Int(tempnum / 16)
    Next loopit

    If tempnum <> 0 Then
        valueans = Hex$(tempnum) + valueans
    End If

    BigAdd = Right(valueans, 8)
End Function
Function BigAND(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) And Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigAND = valueans
End Function
Function BigMod32Add(ByVal value1 As String, ByVal value2 As String) As String
    BigMod32Add = Right$(BigAdd(value1, value2), 8)
End Function
Function BigNOT(ByVal value1 As String) As String
Dim valueans As String
Dim loopit As Integer

    value1 = Right$(value1, 8)
    value1 = String$(8 - Len(value1), "0") + value1
    For loopit = 1 To 8
        valueans = valueans + Hex$(15 Xor Val("&H" + Mid$(value1, loopit, 1)))
    Next loopit

    BigNOT = valueans
End Function
Function BigOR(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Or Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigOR = valueans
End Function
Function BigXOR(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Xor Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigXOR = Right(valueans, 8)
End Function
Public Function DeHex(inp As String) As String
For i = 1 To Len(inp) Step 2
    X = X & Chr(Val("&H" & Mid(inp, i, 2)))
Next i
DeHex = X
End Function
Public Function EnHex(X As String) As String
For i = 1 To Len(X)
    v = Hex(Asc(Mid(X, i, 1)))
    If Len(v) = 1 Then v = "0" & v
    inp = inp & v
Next i
EnHex = inp
End Function
