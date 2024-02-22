VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'                    Metodos publicos
'
' Public sub Push(byval i as variant) mete el elemento i
' al final de la cola.
'
' Public Function Pop As Variant: quita de la cola el primer elem
' y lo devuelve
'
' Public Function VerElemento(ByVal Index As Integer) As Variant
' muestra el elemento numero Index de la cola sin quitarlo
'
' Public Function PopByVal() As Variant: muestra el primer
' elemento de la cola sin quitarlo
'
' Public Property Get Longitud() As Integer: devuelve la
' cantidad de elementos que tiene la cola.

Private Const FRENTE = 1

Private Cola As Collection

Public Sub Reset()
On Error Resume Next
'Call BDDResetGMsos
Dim i As Integer
For i = 1 To Me.Longitud
    Cola.Remove FRENTE
Next i

End Sub

Public Property Get Longitud() As Integer
Longitud = Cola.Count
End Property

Private Function IndexValido(ByVal i As Integer) As Boolean
IndexValido = i >= 1 And i <= Me.Longitud
End Function

Private Sub Class_Initialize()
Set Cola = New Collection
End Sub

Public Function VerElemento(ByVal index As Integer) As Variant
On Error Resume Next
If IndexValido(index) Then
    VerElemento = Cola(index)
Else
    VerElemento = 0
End If

End Function


Public Sub Push(ByVal Info As String, ByVal Nombre As String)
On Error Resume Next
'Mete elemento en la cola
'Call BDDAddGMsos(ReadField(1, Nombre, 59), ReadField(2, Nombre, 59))
Call Cola.Add(UCase$(Nombre))
End Sub

Public Function Pop() As Variant
On Error Resume Next
'Quita elemento de la cola
If Cola.Count > 0 Then
    Pop = Cola(FRENTE)
    Call Cola.Remove(FRENTE)
Else
    Pop = 0
End If
End Function

Public Function PopByVal() As Variant
On Error Resume Next
'Call LogTarea("PopByVal SOS")

'Quita elemento de la cola
If Cola.Count > 0 Then
    PopByVal = Cola.Item(1)
Else
    PopByVal = 0
End If

End Function

Public Function Existe(ByVal Nombre As String) As Boolean
On Error Resume Next
 Dim f As Boolean
 Dim V As Variant
 Dim i As Integer
 Dim NombreEnMayusculas As String
 NombreEnMayusculas = UCase$(Nombre)
 i = 1
 f = False
 
 For i = 1 To Me.Longitud
        V = Me.VerElemento(i)
        If ReadField(1, V, 59) = ReadField(1, NombreEnMayusculas, 59) Then
            Existe = True
            Exit Function
        End If
 Next
End Function

Public Sub Quitar(ByVal Nombre As String)
On Error Resume Next
Dim V As Variant
Dim i As Integer
Dim NombreEnMayusculas As String
'Call BDDDelGMsos(ReadField(1, Nombre, 59))
NombreEnMayusculas = UCase$(Nombre)
For i = 1 To Me.Longitud
    V = Me.VerElemento(i)
    If ReadField(1, V, 59) = ReadField(1, NombreEnMayusculas, 59) Then
        Call Cola.Remove(i)
        Exit Sub
    End If
Next i
 
End Sub

Private Sub Class_Terminate()
'Destruimos el objeto Cola
Set Cola = Nothing
End Sub

'Delzak Sos offline

'Public Sub Escribe(ByVal Nivel As Integer, Nombre As String, moTivo As String)

'Dim nfile As Integer

'nfile = FreeFile
'moTivo = ReadField(2, moTivo, Asc(":"))


'Open App.Path & "\sosoffline.sos" For Append Shared As #nfile
 '   Print #nfile, Nombre & ";(" & Nivel & ")" & ": " & Date & " " & Time & "; " & moTivo
'Close #nfile

'End Sub

'Public Sub Borra(ByVal Nombre As String)

'Dim Linea As String
'foundNombre As Boolean

'Open App.Path & "\sosoffline.sos" For Input As #1
'Open App.Path & "\sosoffline.soscpy" For Append Shared As #2
    
 '   Do Until EOF(1)
        
  '      Line Input #1, Linea
   '     If InStr(Linea, Nombre) = 0 Then Print #2, Linea
       
    'Loop

'Close #1
'Close #2

'Kill App.Path & "\sosoffline.sos"
'Name App.Path & "\sosoffline.soscpy" As App.Path & "\sosoffline.sos"

'End Sub

'Public Function ExisteDelzak(Nombre As String) As Boolean

'Dim Linea As String

'Open App.Path & "\sosoffline.sos" For Input As #1
    'Do Until EOF(1)
        
'        Line Input #1, Linea
        'If InStr(Linea, Nombre) = 0 Then ExisteDelzak = True
'
    'Loop

'Close #1

'End Function

'Public Property Get LongitudDelzak() As Integer

'Dim contador As Integer
'Dim Linea As String

'Open App.Path & "\sosoffline.sos" For Input As #1
    'Do While Not EOF(1)
    
     '   Line Input #1, Linea
        'contador = contador + 1
       
    'Loop

'Close #1

'LongitudDelzak = contador

'End Property

'Public Function VerElementoDelzak(ByVal index As Integer, Nombre As String) As Variant

'On Error Resume Next

'Dim Linea As String
'Dim contador As Integer
'Dim online As Boolean
'Dim User As Integer

'online = False
'Open App.Path & "\sosoffline.sos" For Input As #1
 '   Do Until EOF(1)
        
  '      Line Input #1, Linea
   '     contador = contador + 1
    '    If contador = index Then
     '       Nombre = ReadField(1, Linea, Asc(";"))
      '      For User = 1 To LastUser
       '         If UserList(User).Name = Nombre Then
        '                VerElementoDelzak = "ONLine --> " & Linea
         '           online = True
          '      End If
           ' Next
            
           ' If online = False Then VerElementoDelzak = "OFFLine--> " & Linea
            
        'End If
    'Loop

'Close #1

'End Function

'Public Function Respuesta(ByVal UserIndex As Integer) As Variant

'Dim Linea As String
'Dim Nombre As String

'Nombre = UserList(UserIndex).Name

'Open App.Path & "\SosOfflineRespuestas.sos" For Input As #1
 '   Do Until EOF(1)
        
  '      Line Input #1, Linea
   '     If ReadField(2, Linea, Asc(";")) = Nombre Then
    '        Respuesta = Linea
     '       Close #1
      '      Exit Function
       ' End If
    'Loop
'Close #1


'End Function

'Public Function CuantasVecesAparece(ByVal UserIndex As Integer) As Integer

'Dim Linea As String
'Dim Nombre As String
'Dim contador As Integer

'Nombre = UserList(UserIndex).Name
  
'Open App.Path & "\SosOfflineRespuestas.sos" For Input As #1
 '   Do Until EOF(1)
        
  '      Line Input #1, Linea
   '     If ReadField(2, Linea, Asc(";")) = Nombre Then contador = contador + 1
'
'    Loop
'Close #1

'CuantasVecesAparece = contador

'End Function

'Public Function Busca(n As Integer, UserIndex As Integer)

'Dim Linea As String
'Dim Nombre As String
'Dim contador As Integer

'Nombre = UserList(UserIndex).Name

'Open App.Path & "\SosOfflineRespuestas.sos" For Input As #1
 '   Do Until EOF(1)
        
  '      Line Input #1, Linea
   '     If ReadField(2, Linea, Asc(";")) = Nombre Then
    '        contador = contador + 1
     '       If contador = n Then
      '         Busca = Linea
       '        Close #1
        '       Exit Function
         '   End If

       ' End If
    'Loop
'Close #1

'End Function
