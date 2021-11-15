Attribute VB_Name = "Drag_Zonas"
Public Type ZonaInfo
  nombre As String
  Mapa As Integer
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
  grh As Long
End Type

Public Enum permiso_zona
  no_invisibilidad = 1
  no_atacar = 2
End Enum

Public Enum ZonaEvento
  Entrar
  Moverse
  Estar
  Salir
End Enum

Public Type ZonaEfecto
  EfectoIndex As Long
  EfectoTipo As EfectoTipo
End Type

Public ZonaList() As ZonaInfo
Public ZonaEfecto() As ZonaEfecto

Option Explicit

Public Sub cargar_zonas_sql()
  On Error GoTo errHandler
  
  Dim i As Integer
  ReDim ZonaList(0) As ZonaInfo
  
  '11/12/2018 Irongete: Cargo las zonas en memoria
  Call CheckSQL
  Dim RS As ADODB.Recordset
  Set RS = New ADODB.Recordset
  Set RS = SQL.Execute("SELECT id, nombre, mapa, x1, y1, x2, y2, permisos, grh FROM zona")
  
  While Not RS.EOF
  
    '12/11/2018 Irongete: Añado la zona al array de Zonas
    i = UBound(ZonaList)
    ZonaList(i).nombre = RS!nombre
    ZonaList(i).Mapa = RS!Mapa
    ZonaList(i).x1 = RS!x1
    ZonaList(i).x2 = RS!x2
    ZonaList(i).y1 = RS!y1
    ZonaList(i).y2 = RS!y2
    ZonaList(i).permisos = RS!permisos
    ZonaList(i).grh = RS!grh
    ReDim Preserve ZonaList(i + 1) As ZonaInfo
    
    '12/11/2018 Irongete: Asigno la zona recién creada al mapa
    MapInfo(RS!Mapa).Zonas.Add i, "zona" & i
    
    '12/11/2018 Irongete: Tiene que spawnear algun npc?
    Set RS2 = SQL.Execute("SELECT id_npc FROM rel_zona_npc WHERE id_zona = '" & RS!id & "'")
    While Not RS2.EOF
      Dim Pos As worldPos
      Pos = RandomEnZona(RS!Mapa, RS!x1, RS!x2, RS!y1, RS!y2)
    
      Dim NpcId As Integer
      NpcId = SpawnNpc(RS2!id_npc, Pos, False, False, False, 0)
      NPCList(NpcId).flags.zona = RS!id

      RS2.MoveNext
    Wend
    RS.MoveNext
  Wend
  
  Exit Sub
  
errHandler:
  Debug.Print "error en cargar_zonas()"; Err.Description
End Sub
Public Function permiso_en_zona(ByVal UserIndex As Integer) As Integer
  
  Dim i As Integer
  For i = 1 To UserList(UserIndex).zona.count
    permiso_en_zona = ZonaList(UserList(UserIndex).zona(i)).permisos
  Next i
  
  permiso_en_zona = permiso_en_zona

End Function


Public Function jugador_en_zona(ByVal zona As Integer, ByVal UserIndex As Integer)
  Dim i As Integer
  Dim count As Integer
  count = ZonaList(zona).jugador.count
  For i = 1 To count
    If ZonaList(zona).jugador(i) = UserIndex Then
      jugador_en_zona = True
    End If
  Next
  jugador_en_zona = jugador_en_zona
End Function

Public Sub comprobar_zona(ByVal UserIndex As Integer)
  Dim TempIndex As Integer
  Dim TempZona As Integer
  Dim Map As Integer
  Dim X As Byte
  Dim Y As Byte
  Dim pisa_zona As Boolean
  Dim EfectoIndex As Variant
  
  Map = UserList(UserIndex).Pos.Map
  X = UserList(UserIndex).Pos.X
  Y = UserList(UserIndex).Pos.Y
  With MapInfo(Map)
    Dim zona As Variant
    For Each zona In .Zonas
      '11/12/2018 Irongete: Compruebo si está pisando la zona
      pisa_zona = False
      If X >= ZonaList(zona).x1 And X <= ZonaList(zona).x2 Then
        If Y >= ZonaList(zona).y1 And Y <= ZonaList(zona).y2 Then
          pisa_zona = True
        End If
      End If
      
      '12/12/2018 Irongete: Está pisando zona?
      If pisa_zona = True Then
      
        '12/12/2018 Irongete: La estaba pisando ya antes? Se mueve.
        If jugador_en_zona(zona, UserIndex) = True Then
                
           '14/12/2018 Irongete: Le creo y añado al jugador los efectos que tiene esta zona al salir
          For Each EfectoIndex In ZonaList(zona).efecto_al_moverse
            Call ejecutar_efecto(EfectoIndex, UserIndex, "zona_moverse" & zona)
          Next
          
        '12/12/2018 Irongete: No la pisaba. Entra.
        Else
          ZonaList(zona).jugador.Add UserIndex, "jugador" & UserIndex
          UserList(UserIndex).zona.Add zona, "zona" & zona
          Call WriteConsoleMsg(UserIndex, "ENTRAS EN " & ZonaList(zona).nombre & "(" & zona & ")", FontTypeNames.FONTTYPE_INFO)
          
          '14/12/2018 Irongete: Le creo y añado al jugador los efectos que tiene esta zona al entrar
          For Each EfectoIndex In ZonaList(zona).efecto_al_entrar
            Call ejecutar_efecto(EfectoIndex, "zona_entrar" & zona, UserIndex)
          Next
        End If
        
        '13/12/2018 Irongete: Da igual como haya llegado a la zona, comprobamos si está y puede estar invisible
        If UserList(UserIndex).flags.invisible = 1 And (ZonaList(zona).permisos And permiso_zona.no_invisibilidad) Then
          Call WriteConsoleMsg(UserIndex, "En esta zona no está permitida la invisibilidad.", FontTypeNames.FONTTYPE_INFO)
          Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
          UserList(UserIndex).flags.invisible = 0
        End If
      Else
      
        '12/12/2018 Irongete: No pisa zona. Está en alguna zona? Sale.
        If jugador_en_zona(zona, UserIndex) = True Then
          ZonaList(zona).jugador.Remove "jugador" & UserIndex
          UserList(UserIndex).zona.Remove "zona" & zona
          Call WriteConsoleMsg(UserIndex, "SALES DE " & ZonaList(zona).nombre & "(" & zona & ")", FontTypeNames.FONTTYPE_INFO)
          
          '14/12/2018 Irongete: Le creo y añado al jugador los efectos que tiene esta zona al salir
          For Each EfectoIndex In ZonaList(zona).efecto_al_salir
            Call ejecutar_efecto(EfectoIndex, "zona_salir" & zona, UserIndex)
          Next
          
        End If
      End If
    Next
  End With
End Sub


Public Sub RespawnNpcZona(NpcIndex As npc)
  Dim zona As Integer
  Dim Pos As worldPos
  zona = NpcIndex.flags.zona
  
  Call CheckSQL
  Dim RS As ADODB.Recordset
  Set RS = New ADODB.Recordset
  Set RS = SQL.Execute("SELECT mapa, x1, y1, x2, y2 FROM zona WHERE id = '" & zona & "'")
  If Not RS.EOF Then
    Pos = RandomEnZona(RS!Mapa, RS!x1, RS!x2, RS!y1, RS!y2)
    Dim NpcId As Integer
    NpcId = SpawnNpc(NpcIndex.Numero, Pos, False, False, False, 0)
    NPCList(NpcId).flags.zona = zona
    Debug.Print "ZONA "; zona; " RESPAWN "; NpcIndex.Numero; " (" & Pos.X & ","; Pos.Y & ")"
  End If
End Sub


Public Function RandomEnZona(ByVal Mapa As Integer, ByVal x1 As Byte, ByVal x2 As Byte, ByVal y1 As Byte, ByVal y2 As Byte) As worldPos
  Dim RandomX As Byte
  Dim RandomY As Byte
  RandomX = RandomNumber(x1, x2)
  RandomY = RandomNumber(y1, y2)
  RandomEnZona.Map = Mapa
  RandomEnZona.X = RandomX
  RandomEnZona.Y = RandomY
  RandomEnZona = RandomEnZona
End Function


'11/12/2018 Irongete: La funcion comprueba si el UserIndex está dentro del array de jugadores
' de ZonaList(mapa).jugadores()
Public Function ya_en_zona(ByVal UserIndex, ByVal zona_id) As Boolean
  Dim i As Integer
  Dim Mapa As Integer

  ya_en_zona = False
  
  Mapa = UserList(UserIndex).Pos.Map
    
  For i = LBound(MapInfo(Mapa).Zonas(zona_id).jugador) To UBound(MapInfo(Mapa).Zonas(zona_id).jugador)
    If MapInfo(Mapa).Zonas(zona_id).jugador(i) = UserIndex Then
      ya_en_zona = True
      Exit For
    End If
  Next
    
End Function

Public Function EnviarZonas(ByVal UserIndex As Integer, Mapa As Integer)
  Dim numZona As Variant
  For Each numZona In MapInfo(Mapa).Zonas
    Dim ZonaIndex As Long
    Call WriteCrearZona(UserIndex, numZona)
  Next


End Function
