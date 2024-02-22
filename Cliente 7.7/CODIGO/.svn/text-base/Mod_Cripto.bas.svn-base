Attribute VB_Name = "Mod_Cripto"
Option Explicit
Public Seed As String

Function EncryptINI$(Strg$, Password$)
   Dim b$, S$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, P$
   j = 1
   For i = 1 To Len(Password$)
     P$ = P$ & Asc(Mid$(Password$, i, 1))
   Next
    
   For i = 1 To Len(Strg$)
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     A2 = Asc(Mid$(Strg$, i, 1))
     A3 = A1 Xor A2
     b$ = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b$
     S$ = S$ + b$
   Next
   EncryptINI$ = S$
End Function

Function DecryptINI$(Strg$, Password$)
   Dim b$, S$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, P$
   j = 1
   For i = 1 To Len(Password$)
     P$ = P$ & Asc(Mid$(Password$, i, 1))
   Next
   
   For i = 1 To Len(Strg$) Step 2
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     b$ = Mid$(Strg$, i, 2)
     A3 = Val("&H" + b$)
     A2 = A1 Xor A3
     S$ = S$ + Chr$(A2)
   Next
   DecryptINI$ = S$
End Function

Function Crypt$(Strg$, Password$)
   Dim S$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, P$
   j = 1
   For i = 1 To Len(Password$)
     P$ = P$ & Asc(Mid$(Password$, i, 1))
   Next
   For i = 1 To Len(Strg$)
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     A2 = Asc(Mid$(Strg$, i, 1))
     A3 = A1 Xor A2
     S$ = S$ + Chr$(A3)
     'If i Mod 4096 = 0 Then j = 1
   Next
   Crypt$ = S$
End Function
