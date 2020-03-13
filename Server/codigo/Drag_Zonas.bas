Attribute VB_Name = "Drag_Zonas"
Public Type ZonaInfo
  nombre As String
  Mapa As Integer
  x1 As Byte
  y1 As Byte
  x2 As Byte
  y2 As Byte
  jugador() As Long 'lista de UserIndex de los jugadores que están dentro de la zona
  npc() As Long 'lista de NpcIndex de los npcs que están dentro de la zona
  '14/12/2018 Irongete: Cada zona tiene una colección de efectos que se ejecutan cada vez que se entra, se está, se camina o se sale de la ella
  efecto_al_entrar As New Collection 'esto lo controla Drag_Zonas.comprobar_zona()
  efecto_al_moverse As New Collection 'esto lo controla Drag_Zonas.comprobar_zona()
  efecto_al_estar As New Collection 'esto lo controla Drag_Efectos.procesar_efectos()
  efecto_al_salir As New Collection 'esto lo controla Drag_Zonas.comprobar_zona()
  permisos As Integer
  prioridad As Byte 'para los permisos, si un jugador está en dos zonas a la vez, la que tenga este número mas alto aplicará los permisos
  grh As Long
  
  'servidor
  temporal As Boolean
  Duracion As Integer
  dueño_jugador As Boolean
  dueño As Long
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

Public Function crear_zona()

  '18/12/2018 Irongete: Creo la zona
  ReDim Preserve ZonaList(UBound(ZonaList) + 1)
  
  Dim ZonaIndex As Long
  ZonaIndex = UBound(ZonaList)
  ReDim ZonaList(ZonaIndex).jugador(0)
  crear_zona = ZonaIndex
        
End Function

'esto lo que hace es comprobar si a una
'zona temporal le llega la duración a 0 y hay que borrarla
Public Sub duracion_zonas()

  Dim ZonaIndex As Long
  For ZonaIndex = 0 To UBound(ZonaList)
    
    '18/12/2018 Irongete: Es temporal?
    If ZonaList(ZonaIndex).temporal = True Then
      
      '18/12/2018 Irongete: Ha llegado la duración a 0?
      If ZonaList(ZonaIndex).Duracion <= 0 Then
        Call quitar_zona(ZonaIndex)
        Exit Sub
      Else
        ZonaList(ZonaIndex).Duracion = ZonaList(ZonaIndex).Duracion - 100
      End If
    
    End If
  Next
  
End Sub


Public Sub quitar_zona_de_mapa(ByVal Mapa As Integer, ByVal QuitarIndex As Long)

  'busco el index
  Dim ZonaIndex As Long
  For ZonaIndex = 0 To UBound(MapInfo(Mapa).Zonas)
    If MapInfo(Mapa).Zonas(ZonaIndex) = QuitarIndex Then
      'me la apunto
      ZonaIndex = MapInfo(Mapa).Zonas(ZonaIndex)
      Exit For
    End If
  Next
  
  'reordeno el array para borrar la zona
  For ZonaIndex = ZonaIndex To UBound(MapInfo(ZonaIndex).Zonas) - 1
    MapInfo(Mapa).Zonas(ZonaIndex) = MapInfo(Mapa).Zonas(ZonaIndex + 1)
  Next
  
  If UBound(MapInfo(Mapa).Zonas) > 0 Then
    ReDim Preserve MapInfo(Mapa).Zonas(UBound(MapInfo(Mapa).Zonas) - 1)
  Else
    ReDim MapInfo(Mapa).Zonas(0)
  End If

End Sub

'esta funcion quita una zona
Public Sub quitar_zona(ByVal QuitarIndex As Long)


  Dim UserIndex As Integer
  For UserIndex = 1 To LastUser
    If UserList(UserIndex).ConnIDValida = True Then
      Call WriteQuitarZona(UserIndex, QuitarIndex)
    End If
  Next

  'quito la zona del MapInfo
  Call quitar_zona_de_mapa(ZonaList(QuitarIndex).Mapa, QuitarIndex)
  
  
  

  'quito la zona de zonalist
  Dim ZonaIndex As Long
  For ZonaIndex = QuitarIndex To UBound(ZonaList) - 1
    ZonaList(ZonaIndex) = ZonaList(ZonaIndex + 1)
  Next
  
  If UBound(ZonaList) > 0 Then
    ReDim Preserve ZonaList(UBound(ZonaList) - 1) As ZonaInfo
  Else
    ReDim ZonaList(0)
  End If
  
  
  
 
  
  
End Sub

Public Sub añadir_zona_a_mapa(ByVal Mapa As Integer, ByVal ZonaIndex As Long)

  'hago sitio
  ReDim Preserve MapInfo(Mapa).Zonas(UBound(MapInfo(Mapa).Zonas) + 1)
  
  Dim num_zona As Long
  num_zona = UBound(MapInfo(Mapa).Zonas)
  MapInfo(Mapa).Zonas(num_zona) = ZonaIndex

End Sub

Public Sub cargar_zonas_sql()

  'frmZonas.Show
  
  On Error GoTo errHandler
  
  Dim i As Integer
  ReDim ZonaList(0) As ZonaInfo
  
  '11/12/2018 Irongete: Cargo las zonas en memoria
  Call check_sql
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
    ZonaList(i).temporal = False 'por defecto del SQL es false
    
       
    ReDim Preserve ZonaList(i + 1) As ZonaInfo
    ReDim ZonaList(i).jugador(0) As Long
    ReDim ZonaList(i + 1).jugador(0) As Long
    
    
    '12/11/2018 Irongete: Asigno la zona recién creada al mapa
    Call añadir_zona_a_mapa(RS!Mapa, i)
    'MapInfo(RS!Mapa).Zonas.Add i, "zona" & i
    
    '12/11/2018 Irongete: Tiene que spawnear algun npc?
    Set RS2 = SQL.Execute("SELECT id_npc FROM rel_zona_npc WHERE id_zona = '" & RS!id & "'")
    While Not RS2.EOF
      Dim Pos As worldPos
      Pos = RandomEnZona(RS!Mapa, RS!x1, RS!x2, RS!y1, RS!y2)
    
      Dim NpcId As Integer
      NpcId = SpawnNpc(RS2!id_npc, Pos, False, False, False, 0)
      NPCList(NpcId).flags.Zona = RS!id

      RS2.MoveNext
    Wend
    RS2.Close
    Set RS2 = Nothing
    RS.MoveNext
  Wend
  RS.Close
  Set RS = Nothing
  
errHandler:
  Debug.Print "error en cargar_zonas()"; Err.Description
End Sub
Public Function permiso_en_zona(ByVal UserIndex As Integer) As Integer
  
  Dim i As Integer
  For i = 1 To UserList(UserIndex).Zona.count
    permiso_en_zona = ZonaList(UserList(UserIndex).Zona(i)).permisos
  Next i
  
  permiso_en_zona = permiso_en_zona

End Function

Public Sub añadir_jugador_a_zona(ByVal UserIndex As Integer, ByVal ZonaIndex As Long)

  
  Debug.Print LBound(ZonaList(ZonaIndex).jugador)
  Debug.Print UBound(ZonaList(ZonaIndex).jugador)
  
  
  'hago sitio
  ReDim Preserve ZonaList(ZonaIndex).jugador(UBound(ZonaList(ZonaIndex).jugador) + 1)
 
  'meto el userindex
  ZonaList(ZonaIndex).jugador(UBound(ZonaList(ZonaIndex).jugador)) = UserIndex
End Sub

Public Sub quitar_jugador_de_zona(ByVal UserIndex As Integer, ByVal ZonaIndex As Long)
  'busco el index
  Dim JugadorIndex As Long
  For JugadorIndex = 0 To UBound(ZonaList(ZonaIndex).jugador)
    If ZonaList(ZonaIndex).jugador(JugadorIndex) = UserIndex Then
      'me lo apunto
      JugadorIndex = ZonaList(ZonaIndex).jugador(JugadorIndex)
      Exit For
    End If
  Next
  
  'reordeno el array para borrar al jugador
  For JugadorIndex = JugadorIndex To UBound(ZonaList(ZonaIndex).jugador) - 1
    ZonaList(ZonaIndex).jugador(JugadorIndex) = ZonaList(ZonaIndex).jugador(JugadorIndex + 1)
  Next
  
  If UBound(ZonaList(ZonaIndex).jugador) > 0 Then
    ReDim Preserve ZonaList(ZonaIndex).jugador(UBound(ZonaList(ZonaIndex).jugador) - 1)
  Else
    ReDim ZonaList(ZonaIndex).jugador(0)
  End If
  
End Sub

Public Function jugador_en_zona(ByVal ZonaIndex As Integer, ByVal UserIndex As Integer) As Boolean

  'busco al jugador
  Dim JugadorIndex As Long
  
  For JugadorIndex = 0 To UBound(ZonaList(ZonaIndex).jugador)
    If ZonaList(ZonaIndex).jugador(JugadorIndex) = UserIndex Then
      jugador_en_zona = True
    End If
  Next
End Function

Public Sub comprobar_zona(ByVal UserIndex As Long)
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
    Dim Zona As Variant
    For Each Zona In .Zonas
      '11/12/2018 Irongete: Compruebo si está pisando la zona
      pisa_zona = False
      If X >= ZonaList(Zona).x1 And X <= ZonaList(Zona).x2 Then
        If Y >= ZonaList(Zona).y1 And Y <= ZonaList(Zona).y2 Then
          pisa_zona = True
        End If
      End If
      
      '12/12/2018 Irongete: Está pisando zona?
      If pisa_zona = True Then
      
        '12/12/2018 Irongete: La estaba pisando ya antes? Se mueve.
        '19/12/2018 Irongete: ***************** MOVERSE DENTRO DE LA ZONA ********************************
        If jugador_en_zona(Zona, UserIndex) = True Then
                
           '14/12/2018 Irongete: Le creo y añado al jugador los efectos que tiene esta zona al salir
          For Each EfectoIndex In ZonaList(Zona).efecto_al_moverse
            '17/12/2018 Irongete: El efecto es aplicado? Lo meto al usuario en vez de ejecutar
            If EfectoList(EfectoIndex).aplicado = True Then
              Call añadir_efecto_a_jugador(EfectoIndex, UserIndex, "zona_entrar")
            Else
              Dim TargetIndex As Long
              TargetIndex = UserIndex
              Call ejecutar_efecto(EfectoIndex, TargetIndex)
            End If
            
          Next
          
        '12/12/2018 Irongete: No la pisaba. Entra.
        '19/12/2018 Irongete: ***************** ENTRAR EN LA ZONA ********************************
        Else
        
          Call añadir_jugador_a_zona(UserIndex, Zona)
          'UserList(UserIndex).Zona.Add Zona, "zona" & Zona
          Call WriteConsoleMsg(UserIndex, "ENTRAS EN " & ZonaList(Zona).nombre & " (" & Zona & ")", FontTypeNames.FONTTYPE_INFO)
          
          '14/12/2018 Irongete: Le creo y añado al jugador los efectos que tiene esta zona al entrar
          For Each EfectoIndex In ZonaList(Zona).efecto_al_entrar
            
            '17/12/2018 Irongete: El efecto es aplicado? Lo meto al usuario en vez de ejecutar
            If EfectoList(EfectoIndex).aplicado = True Then
              Call añadir_efecto_a_jugador(EfectoIndex, UserIndex, "zona_entrar")
            Else
              TargetIndex = UserIndex
              Call ejecutar_efecto(EfectoIndex, TargetIndex)
            End If
          Next
        End If
        
        '13/12/2018 Irongete: Da igual como haya llegado a la zona, comprobamos si está y puede estar invisible
        If UserList(UserIndex).flags.invisible = 1 And (ZonaList(Zona).permisos And permiso_zona.no_invisibilidad) Then
          Call WriteConsoleMsg(UserIndex, "En esta zona no está permitida la invisibilidad.", FontTypeNames.FONTTYPE_INFO)
          Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
          UserList(UserIndex).flags.invisible = 0
        End If
      Else
      
        '12/12/2018 Irongete: No pisa zona. Está en alguna zona? Sale.
        '19/12/2018 Irongete: ***************** SALIR DE LA ZONA ********************************
        If jugador_en_zona(Zona, UserIndex) = True Then
          Call quitar_jugador_de_zona(UserIndex, Zona)
          'ZonaList(Zona).jugador.Remove "jugador" & UserIndex
          'UserList(UserIndex).Zona.Remove "zona" & Zona
          Call WriteConsoleMsg(UserIndex, "SALES DE " & ZonaList(Zona).nombre & " (" & Zona & ")", FontTypeNames.FONTTYPE_INFO)
          
          '14/12/2018 Irongete: Le creo y añado al jugador los efectos que tiene esta zona al salir
          For Each EfectoIndex In ZonaList(Zona).efecto_al_salir
            '17/12/2018 Irongete: El efecto es aplicado? Lo meto al usuario en vez de ejecutar
            If EfectoList(EfectoIndex).aplicado = True Then
              Call añadir_efecto_a_jugador(EfectoIndex, UserIndex, "zona_entrar")
            Else
              TargetIndex = UserIndex
              Call ejecutar_efecto(EfectoIndex, UserIndex)
            End If
          Next
          
        End If
      End If
    Next
  End With
End Sub


Public Sub RespawnNpcZona(NpcIndex As npc)
  Dim Zona As Integer
  Dim Pos As worldPos
  Zona = NpcIndex.flags.Zona
  
  Call CheckSQL
  Dim RS As ADODB.Recordset
  Set RS = New ADODB.Recordset
  Set RS = SQL.Execute("SELECT mapa, x1, y1, x2, y2 FROM zona WHERE id = '" & Zona & "'")
  If Not RS.EOF Then
    Pos = RandomEnZona(RS!Mapa, RS!x1, RS!x2, RS!y1, RS!y2)
    Dim NpcId As Integer
    NpcId = SpawnNpc(NpcIndex.Numero, Pos, False, False, False, 0)
    NPCList(NpcId).flags.Zona = Zona
    Debug.Print "ZONA "; Zona; " RESPAWN "; NpcIndex.Numero; " (" & Pos.X & ","; Pos.Y & ")"
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
  Dim ZonaIndex As Long
  
  For ZonaIndex = 0 To UBound(MapInfo(Mapa).Zonas)
    Call WriteCrearZona(UserIndex, MapInfo(Mapa).Zonas(ZonaIndex))
  Next


End Function
