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
    UserKey As String, Text As String, Action As Single _
    ) As String
    Dim UserKey2 As String
    Dim Temp     As Integer
    Dim Times    As Integer
    Dim i        As Integer
    Dim j        As Integer
    Dim n        As Integer
    Dim rtn      As String
    Dim TExtito As String
    Dim Calcu As Integer
    
  'DoEvents
    '//Get UserKey characters
    Dim a As String
    Dim b As String
     'pluto:6.7
     TExtito = Text
   'If Len(Text) > 6 Then UserKey = " 2222"
    'If Len(Text) > 12 Then UserKey = " 2332"
    a = Left$(Text, 1)
    b = Right$(Text, 1)

If Action = 2 Then
Text = Mid$(Text, 2, Len(Text))
'UserRecibe = UserRecibe + 1
'If UserRecibe > 50 Then UserRecibe = 1
Calcu = Val(UserKey) - (Len(Text) * 3) + MacClave
UserKey2 = " " & CStr(Calcu)
'Debug.Print "Recibe: " & UserRecibe & " KEY: " & UserKey
'Call AddtoRichTextBox(frmMain.RecTxt, "D: " & UserKey, 100, 100, 120, 0, 0)

End If

If Action = 1 Then
Text = Mid$(Text, 2, Len(Text) - 2)
'UserEnvia = UserEnvia + 1
'If UserEnvia > 35 Then UserEnvia = 1

Calcu = Val(UserKey) - (Len(Text) * 3) + MacClave
UserKey2 = " " & CStr(Calcu)
'Debug.Print "Envia: " & UserEnvia & " KEY: " & UserKey
'Call AddtoRichTextBox(frmMain.RecTxt, "C: " & UserKey, 100, 100, 120, 0, 0)

End If
   'pluto:6.7

  ' If Len(Text) > 6 Then
  'Mid$(UserKey, 3, 3) = "8"
  'End If
   '--------
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
           Temp = TextASCIIS(i) + UserKeyASCIIS(j) + Int(MacClave / 10)
           If Temp > 255 Then
              
            Temp = Temp - 255
           End If
            
           rtn = rtn + Chr$(Temp)
   
       Next
        CodificaR = a & rtn & b
       'Call LogError("CODIFICAR: " & Text & " --> " & CodificaR & "Key: " & UserKey & "Key2: " & UserKey2 & " Userenvia: " & UserEnvia)

    ElseIf Action = DECRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
          'If TextASCIIS(i) < 32 Then UserKeyASCIIS(j) = 0
            Temp = TextASCIIS(i) - UserKeyASCIIS(j) - Int(MacClave / 20)
           If Temp < 0 Then
              Temp = Temp + 255
           End If
           rtn = rtn + Chr$(Temp)
   
       Next
        CodificaR = a & rtn
        'Call LogError("DECODIFICAR: " & Text & " --> " & CodificaR & "Key: " & UserKey & "Key2: " & UserKey2 & " UserRecibe: " & UserRecibe)

    End If
    
    '//Return
    

End Function

