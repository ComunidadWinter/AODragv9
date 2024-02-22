Attribute VB_Name = "Codifica"
'pluto:2.5.0
DefInt A-Z
Option Explicit

'//For Action parameter in EncryptString
Public Const ENCRYPT = 1, DECRYPT = 2
'pluto:2.5.0
'---------------------------------------------------------------------
' EncryptString
' Modificado por Harvey T.
'---------------------------------------------------------------------
Public Function CodificaR( _
    UserKey As String, Text As String, UserIndex As Integer, Action As Single _
    ) As String
    On Error GoTo fallo
    Dim UserKey2 As String
    Dim temp     As Integer
    Dim Times    As Integer
    Dim i        As Integer
    Dim j        As Integer
    Dim n        As Integer
    Dim rtn      As String
     Dim a As String
    Dim b As String
    Dim Textito As String
    Dim Calcu As Integer
    'pluto:6.7
   'If Len(Text) > 6 Then UserKey = " 2222"
    'If Len(Text) > 12 Then UserKey = " 2332"
  'Mid$(UserKey, 3, 3) = "8"
  'End If
   '--------
    Textito = Text
   
    a = Left$(Text, 1)
    b = Right$(Text, 1)
    
'pluto:6.7-------------------------
If Action = 1 Then
'UserList(UserIndex).Counters.UserEnvia = UserList(UserIndex).Counters.UserEnvia + 1
'If UserList(UserIndex).Counters.UserEnvia > 50 Then UserList(UserIndex).Counters.UserEnvia = 1
'Text = Textito
Text = Mid$(Text, 2, Len(Text) - 2)
Calcu = val(UserKey) - (Len(Text) * 3) + UserList(UserIndex).MacClave
UserKey2 = str(Calcu)
'Debug.Print "Envia: " & UserList(UserIndex).Counters.UserEnvia & " KEY: " & UserKey

End If
    
If Action = 2 Then
'UserList(UserIndex).Counters.UserRecibe = UserList(UserIndex).Counters.UserRecibe + 1
'If UserList(UserIndex).Counters.UserRecibe > 35 Then UserList(UserIndex).Counters.UserRecibe = 1
Text = Mid$(Text, 2, Len(Text))
Calcu = val(UserKey) - (Len(Text) * 3) + UserList(UserIndex).MacClave
UserKey2 = str(Calcu)
'Debug.Print "Recibe: " & UserList(UserIndex).Counters.UserRecibe & " KEY: " & UserKey
'Call LogError("DECODIFICAR: " & Text & "Key: " & UserKey & " Key2: " & UserKey2 & " UserRecibe: " & UserList(UserIndex).Counters.UserRecibe)

End If
'----------------------------------
    
    '//Get UserKey characters
    n = Len(UserKey2)
    ReDim UserKeyASCIIS(1 To n)
    For i = 1 To n
        UserKeyASCIIS(i) = Asc(Mid$(UserKey2, i, 1))
    Next
        
    '//Get Text characters
    ReDim TextASCIIS(Len(Text)) As Integer
    For i = 1 To Len(Text)
        TextASCIIS(i) = Asc(Mid$(Text, i, 1))
    Next
    
    '//Encryption/Decryption
    If Action = ENCRYPT Then
   
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           'If TextASCIIS(i) < 32 Then UserKeyASCIIS(j) = 0
           temp = TextASCIIS(i) + UserKeyASCIIS(j) + Int(UserList(UserIndex).MacClave / 20)
           
           If temp > 255 Then
            temp = temp - 255
              'MsgBox ("limite: " & Temp)
           End If
           'If TextASCIIS(i) > 230 Then
                      
           
            rtn = rtn + Chr$(temp)
           
       Next
    CodificaR = a & rtn & b
    'Call LogError("CODIFICAR: " & Text & " --> " & CodificaR & " Key: " & UserKey & " Key2: " & UserKey2 & " Userenvia: " & UserList(UserIndex).Counters.UserEnvia)

    ElseIf Action = DECRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
   ' If TextASCIIS(i) < 32 Then UserKeyASCIIS(j) = 0
temp = TextASCIIS(i) - UserKeyASCIIS(j) - Int(UserList(UserIndex).MacClave / 10)
           If temp < 0 Then
              temp = temp + 255
           End If
      
           rtn = rtn + Chr$(temp)
       Next
        CodificaR = a & rtn
        'Call LogError("DECODIFICAR: " & Text & " --> " & CodificaR & " Key: " & UserKey & " Key2: " & UserKey2 & " UserRecibe: " & UserList(UserIndex).Counters.UserRecibe)

    End If
    
    '//Return
Exit Function
fallo:
Call LogError("CODIFICAR: " & Err.number & " D: " & Err.Description & " Text: " & Textito & "Key: " & UserKey & " Action: " & Action)

End Function



