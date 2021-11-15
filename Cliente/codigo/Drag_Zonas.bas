Attribute VB_Name = "Drag_Zonas"
Public Type ZonaInfo
  id As Long
  Nombre As String
  mapa As Integer
  x1 As Byte
  y1 As Byte
  x2 As Byte
  y2 As Byte
  jugador As New Collection 'lista de UserIndex de los jugadores que están dentro de la zona
  npc As New Collection 'lista de NpcIndex de los npcs que están dentro de la zona
  '14/12/2018 Irongete: Cada zona tiene una colección de efectos que se ejecutan cada vez que se entra, se está, se camina o se sale de la ella
  efecto_al_entrar As New Collection 'esto lo controla Drag_Zonas.comprobar_zona()
  efecto_al_moverse As New Collection 'esto lo controla Drag_Zonas.comprobar_zona()
  efecto_al_estar As New Collection 'esto lo controla Drag_Efectos.procesar_efectos()
  efecto_al_salir As New Collection 'esto lo controla Drag_Zonas.comprobar_zona()
  permisos As Integer
  prioridad As Byte 'para los permisos, si un jugador está en dos zonas a la vez, la que tenga este número mas alto aplicará los permisos
  grh As grh
  
  'cliente
  ZonaIndex As Long 'guarda el indice que tiene la zona en el servidor
End Type

Public ZonaList() As ZonaInfo

Public Enum permiso_zona
  no_invisibilidad = 1
  no_atacar = 2
End Enum

Option Explicit

Public Sub HandleCrearZona()

    'Remove packet ID
    Call incomingData.ReadByte

    Dim ZonaIndex As Long
    ZonaIndex = UBound(ZonaList)
    ReDim Preserve ZonaList(ZonaIndex + 1)
    
    ZonaList(ZonaIndex).ZonaIndex = incomingData.ReadLong()
    ZonaList(ZonaIndex).x1 = incomingData.ReadByte()
    ZonaList(ZonaIndex).x2 = incomingData.ReadByte()
    ZonaList(ZonaIndex).y1 = incomingData.ReadByte()
    ZonaList(ZonaIndex).y2 = incomingData.ReadByte()
    ZonaList(ZonaIndex).permisos = incomingData.ReadInteger()
    
    Dim TmpGrh As Long
    TmpGrh = incomingData.ReadLong()
    Call InitGrh(ZonaList(ZonaIndex).grh, TmpGrh)
    
End Sub
