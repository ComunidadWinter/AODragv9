Attribute VB_Name = "Drag_Efectos"
Public Type EfectoInfo
  id As Integer
  Nombre As String
  tipo As Byte
  descripcion As String
  valor As New Collection
  trigger As New Collection 'sirve esto para que un efecto pueda triggear otros efectos??
  duracion As Integer 'duración en milisegundos
  intervalo As Integer 'cada cuantos milisegundos se ejecuta el efecto
  contador_intervalo As Integer 'contador de los milisegundos
  limite As Byte 'cantidad limite de este efecto que puedes tener
  limite_origen As Byte 'cantidad limite de este efecto que puede tener por cada mismo origen (jugador, zona, npc)
  origen As String 'de donde viene este efecto (jugador, zona entrar/salir/mover, npc...)
  beneficioso As Boolean 'si es beneficioso es un buff de lo contrario es un debuff
  grh As grh
  
  'cliente
  EfectoIndex As Integer 'esto guarda en el cliente el Indice del array EfectoList en donde está su efecto en el server por si se lo quita
  
End Type

Public EfectoList() As EfectoInfo 'aquí están los efectos que se crean y se asignan a jugadores y npcs, ...
Option Explicit



Public Sub HandleCrearEfecto()

    'Remove packet ID
    Call incomingData.ReadByte
   
    Dim i As Integer
    i = UBound(EfectoList)
    ReDim Preserve EfectoList(i + 1)
    
    EfectoList(i).id = incomingData.ReadInteger()
    EfectoList(i).EfectoIndex = incomingData.ReadInteger()
    EfectoList(i).tipo = incomingData.ReadByte()
    EfectoList(i).duracion = incomingData.ReadInteger()
    EfectoList(i).beneficioso = incomingData.ReadBoolean
    
    Dim TmpGrh As Long
    TmpGrh = incomingData.ReadLong
    Call InitGrh(EfectoList(i).grh, TmpGrh)
    
    
End Sub


Public Sub HandleQuitarEfecto()

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim EfectoIndex As Integer
    EfectoIndex = incomingData.ReadInteger()
    
    '15/12/2018 Irongete: Recorro el array y busco este indice
    Dim i As Integer
    For i = 0 To UBound(EfectoList)
      If EfectoList(i).EfectoIndex = EfectoIndex Then
        EfectoList(i).EfectoIndex = -1
      End If
    Next i
    
    
End Sub
