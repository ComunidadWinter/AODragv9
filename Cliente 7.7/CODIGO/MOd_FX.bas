Attribute VB_Name = "MOd_FX"
Option Explicit

Public Const MAX_FX = 300

Type tFX
    vida      As Integer
    x         As Integer
    y         As Integer
    userindex As Integer
    GrhIndex  As Integer
End Type

'Vector que contiene los DIALOGOS
Public FXs(1 To MAX_FX) As tFX

'Apunta a el ultimo mensaje
Public UltimoFX As Integer
'Contiene la cantidad de mensajes activos
Public CantidadFX As Integer

Function PrimerIndiceVacioFX() As Integer
Dim i As Byte
i = 1
Do While i <= MAX_FX And FXs(i).vida <> 0
       i = i + 1
Loop
If FXs(i).vida = 0 Then PrimerIndiceVacioFX = i
End Function



Public Sub CrearFX(ByVal User As Integer, ByVal GrhIndex As Integer, ByVal vida As Integer)
Dim MiUserIndex As Integer
Dim IndiceLibre As Integer

If BuscarUserIndexFX(User, MiUserIndex) Then
      FXs(MiUserIndex).vida = 0
      FXs(MiUserIndex).userindex = 0
End If
    
IndiceLibre = PrimerIndiceVacioFX
FXs(IndiceLibre).vida = Delay
FXs(IndiceLibre).userindex = User
FXs(IndiceLibre).GrhIndex = GrhIndex

If UltimoFX > IndiceLibre Then
        UltimoFX = IndiceLibre
End If
CantidadFX = CantidadFX + 1
  
End Sub

Function BuscarUserIndexFX(User As Integer, MiUser As Integer) As Boolean
Dim i As Integer
i = 1
Do While i < MAX_FX And FXs(i).userindex <> User
       i = i + 1
Loop
If FXs(i).userindex = User Then
        MiUser = i
        BuscarUserIndexFX = True
Else: BuscarUserIndexFX = False
End If
End Function

Public Sub Update_FX_Pos(x As Integer, y As Integer, Index As Integer)
Dim MiUserIndex As Integer
If BuscarUserIndexFX(Index, MiUserIndex) Then
  If FXs(MiUserIndex).vida > 0 Then
            FXs(MiUserIndex).x = x + 5
            FXs(MiUserIndex).y = y + 20
        If FXs(MiUserIndex).vida > 0 Then
           FXs(MiUserIndex).vida = FXs(MiUserIndex).vida - 1
        End If
        If FXs(MiUserIndex).vida = 0 Then
            If MiUserIndex = UltimoFX Then UltimoFX = UltimoFX - 1
            CantidadFX = CantidadFX - 1
        End If
  End If
End If
End Sub


Public Sub MostrarFX()
Dim i As Integer
For i = 1 To CantidadFX
    If FXs(i).vida > 0 Then
         Call DDrawTransGrhtoSurface(BackBufferSurface, Grh(FXs(i).GrhIndex), FXs(i).x, FXs(i).y, 0, 1)
    End If
Next
End Sub

Public Sub QuitarFX(CharIndex As Integer)
Dim i As Integer
i = 1
Do While i < MAX_FX And FXs(i).userindex <> CharIndex
       i = i + 1
Loop
If FXs(i).userindex = CharIndex Then
      FXs(i).vida = 0
End If
End Sub

