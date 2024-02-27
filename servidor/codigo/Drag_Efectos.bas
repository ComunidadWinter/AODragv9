Attribute VB_Name = "Drag_Efectos"
Public Enum EfectoTipo
  daño 'causa daño
  curacion 'causa curación
  aumenta_velocidad
  disminuye_velocidad
End Enum

Public Enum EfectoObjetivo
  jugador
  objetivo
End Enum

Public Type EfectoInfo
  id As Integer
  nombre As String
  tipo As Byte
  descripcion As String
  valor As Long
  trigger As New Collection 'sirve esto para que un efecto pueda triggear otros efectos??
  duracion As Integer 'duración en milisegundos
  intervalo As Integer 'cada cuantos milisegundos se ejecuta el efecto
  limite As Byte 'cantidad limite de este efecto que puedes tener
  limite_origen As Byte 'cantidad limite de este efecto que puede tener por cada mismo origen (jugador, zona, npc)
  beneficioso As Boolean 'si es beneficioso es un buff de lo contrario es un debuff
  grh As Long
  enviar_a_cliente As Boolean
  aplicado As Boolean
  
  'valores que usa el juego
  origen As String 'de donde viene este efecto (jugador, zona entrar/salir/mover, npc...)
  contador_intervalo As Integer 'contador de los milisegundos
  'EfectoIndex As Integer 'esto guarda en el cliente el Indice del array EfectoList en donde está su efecto en el server por si se lo quita
  
End Type

Public Efecto() As EfectoInfo 'aquí están todos los efectos con su configuración
Public EfectoList() As EfectoInfo 'aquí están los efectos que se crean y se asignan a jugadores y npcs, ...

Option Explicit

Public Sub añadir_efecto_a_zona(ByVal ZonaIndex As Long, ByVal EfectoId As Integer, ByVal tipo As Byte)
     
  '16/12/2018 Irongete: Recorro Efecto() buscando la Id
  Dim EfectoIndex As Long
  For EfectoIndex = 0 To UBound(Efecto)
  
    '16/12/2018 Irongete: Lo encontré
    If Efecto(EfectoIndex).id = EfectoId Then
    
      '16/12/2018 Irongete: Añado el efecto a EfectoList
      Dim i As Long
      ReDim Preserve EfectoList(UBound(EfectoList) + 1)
      i = UBound(EfectoList)
      EfectoList(i) = Efecto(EfectoIndex)
      
      Select Case tipo
        Case ZonaEvento.Entrar
          ZonaList(ZonaIndex).efecto_al_entrar.Add i, "efecto" & i
        Case ZonaEvento.Estar
          ZonaList(ZonaIndex).efecto_al_estar.Add i, "efecto" & i
        Case ZonaEvento.Moverse
          ZonaList(ZonaIndex).efecto_al_moverse.Add i, "efecto" & i
        Case ZonaEvento.Salir
          ZonaList(ZonaIndex).efecto_al_salir.Add i, "efecto" & i
      End Select
    End If
  Next
  
  Exit Sub

End Sub


Public Sub cargar_efectos_sql()

  ReDim Efecto(0) As EfectoInfo
  ReDim EfectoList(0) As EfectoInfo
  
  '17/12/2018 Irongete: Cargo los efectos desde la base de datos
  Call CheckSQL
  Dim RS As ADODB.Recordset
  Set RS = New ADODB.Recordset
  Set RS = SQL.Execute("SELECT id, nombre, tipo, descripcion, valor, duracion, intervalo, limite, limite_origen, beneficioso, grh, aplicado, enviar_a_cliente FROM efecto")
  
  Dim i As Long
  While Not RS.EOF
    '17/12/2018 Irongete: Datos del hechizo
    i = UBound(Efecto)
    
    Efecto(i).id = RS!id
    Efecto(i).nombre = RS!nombre
    Efecto(i).tipo = RS!tipo
    Efecto(i).descripcion = RS!descripcion
    Efecto(i).valor = RS!valor
    Efecto(i).duracion = RS!duracion
    Efecto(i).intervalo = RS!intervalo
    Efecto(i).limite = RS!limite
    Efecto(i).limite_origen = RS!limite_origen
    Efecto(i).beneficioso = RS!beneficioso
    Efecto(i).grh = RS!grh
    Efecto(i).aplicado = RS!aplicado
    Efecto(i).enviar_a_cliente = RS!enviar_a_cliente
    
    '17/12/2018 Irongete: Triggers
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    Set RS2 = SQL.Execute("SELECT `trigger` FROM rel_efecto_trigger WHERE id_efecto = '" & RS!id & "'")
    
    While Not RS2.EOF
      Efecto(i).trigger.Add RS!trigger, "trigger" & RS!trigger
      RS2.MoveNext
    Wend
    RS2.Close
    Set RS2 = Nothing
    
    ReDim Preserve Efecto(i + 1) As EfectoInfo
    RS.MoveNext
  Wend
  RS.Close
  Set RS = Nothing
  
  
  Dim asdf As Integer
  For asdf = 0 To UBound(Efecto)
    Debug.Print "INDEX"; asdf; "ID:"; Efecto(asdf).id; "TIPO:"; Efecto(asdf).tipo
  Next
  
    
     
  'test velocidad
  Call añadir_efecto_a_zona(6, 1, ZonaEvento.Entrar) 'aumentar velocidad al entrar
  Call añadir_efecto_a_zona(6, 2, ZonaEvento.Salir) 'disminuir velocidad al salir
  
  
  '16/12/2018 Irongete: Todo cargado, activo el timer.
  frmMain.TimerEfectos.Enabled = True
End Sub

Public Function tiene_efecto(ByVal EfectoId As Integer, ByVal UserIndex As Integer) As Boolean
  Dim EfectoIndex As Variant
  For Each EfectoIndex In UserList(UserIndex).Efecto
    If EfectoList(EfectoIndex).id = EfectoId Then
      tiene_efecto = True
    End If
  Next
End Function


Public Function cantidad_efecto(ByVal efecto_id As Integer, ByVal UserIndex As Integer)
  Dim cantidad As Byte
  cantidad = 0
  Dim i As Variant
  For Each i In UserList(UserIndex).Efecto
    If EfectoList(i).id = efecto_id Then
      cantidad = cantidad + 1
    End If
  Next
cantidad_efecto = cantidad
End Function

Public Function cantidad_efecto_mismo_origen(ByVal efecto_id As Integer, ByVal UserIndex As Integer, ByVal origen As String)
  Dim cantidad As Byte
  cantidad = 0
  Dim i As Variant
  For Each i In UserList(UserIndex).Efecto
    If EfectoList(i).id = efecto_id Then
      If EfectoList(i).origen = origen Then
        cantidad = cantidad + 1
      End If
    End If
  Next
cantidad_efecto_mismo_origen = cantidad
End Function



Public Sub ejecutar_efecto(ByVal EfectoIndex As Integer, ByVal origen As String, ByVal TargetIndex As Integer)
  Dim valor As Variant
  Dim i As Integer
  
  '16/12/2018 Irongete: Es aplicado? Controlamos la duracion...
  If EfectoList(EfectoIndex).aplicado = True Then
  
    '14/12/2018 Irongete: Se ha acabado el efecto?
    If EfectoList(EfectoIndex).duracion >= 0 Then
      EfectoList(EfectoIndex).duracion = EfectoList(EfectoIndex).duracion - 500
      EfectoList(EfectoIndex).contador_intervalo = EfectoList(EfectoIndex).contador_intervalo + 500
  
      '15/12/2018 Irongete: El intervalo permite que se ejecute?
      If EfectoList(EfectoIndex).contador_intervalo >= EfectoList(EfectoIndex).intervalo Then
        EfectoList(EfectoIndex).contador_intervalo = 0
      Else
        Exit Sub
      End If
    Else
      Call WriteConsoleMsg(TargetIndex, "Se te quita: " & EfectoList(EfectoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
      UserList(TargetIndex).Efecto.Remove "efecto" & EfectoIndex
      
      '15/12/2018 Irongete: Mandar el paquete QuitarEfecto al jugador
      If EfectoList(EfectoIndex).enviar_a_cliente = True Then
        Call WriteQuitarEfecto(TargetIndex, EfectoIndex)
      End If
      
      '16/12/2018 Irongete: Elimino el efecto
      Exit Sub
    End If
  End If
  
  Select Case EfectoList(EfectoIndex).tipo
    Case EfectoTipo.daño
      Call WriteConsoleMsg(TargetIndex, EfectoIndex & " " & EfectoList(EfectoIndex).nombre & " te causa " & EfectoList(EfectoIndex).valor & " de daño.", FontTypeNames.FONTTYPE_INFO)
      
    Case EfectoTipo.curacion
      Call WriteConsoleMsg(TargetIndex, EfectoIndex & " " & EfectoList(EfectoIndex).nombre & " te restaura " & EfectoList(EfectoIndex).valor & " de salud.", FontTypeNames.FONTTYPE_INFO)
      
    Case EfectoTipo.aumenta_velocidad
      UserList(TargetIndex).flags.Speed = UserList(TargetIndex).flags.Speed + EfectoList(EfectoIndex).valor
      Call WriteChangeSpeed(TargetIndex, UserList(TargetIndex).Char.CharIndex, UserList(TargetIndex).flags.Speed)
      Call WriteConsoleMsg(TargetIndex, EfectoIndex & " " & "Se te aumenta la velocidad en " & EfectoList(EfectoIndex).valor, FontTypeNames.FONTTYPE_INFO)
            
     Case EfectoTipo.disminuye_velocidad
      UserList(TargetIndex).flags.Speed = UserList(TargetIndex).flags.Speed - EfectoList(EfectoIndex).valor
      Call WriteChangeSpeed(TargetIndex, UserList(TargetIndex).Char.CharIndex, UserList(TargetIndex).flags.Speed)
      Call WriteConsoleMsg(TargetIndex, EfectoIndex & " " & "Se te disminuye la velocidad en " & EfectoList(EfectoIndex).valor, FontTypeNames.FONTTYPE_INFO)
      
  End Select
  
  
  
  
  
  
  '16/12/2018 Irongete: El efecto triggerea otros efectos?
  Dim TriggerIndex As Variant
  For Each TriggerIndex In EfectoList(EfectoIndex).trigger
  
    Dim TmpEfecto As EfectoInfo
    TmpEfecto = crear_efecto(TriggerIndex) 'este es el efecto que se le va a aplicar
    
    '16/12/2018 Irongete: Es aplicado?
    If TmpEfecto.aplicado = True Then
    
      '15/12/2018 Irongete: Puede tener mas?
      Dim TieneIndex As Variant
      Dim cantidad As Byte
      Dim cantidad_origen As Byte
      For Each TieneIndex In UserList(TargetIndex).Efecto 'recorro todos sus efectos para ver si cumple los limites
        If EfectoList(TieneIndex).id = TmpEfecto.id Then
          cantidad = cantidad + 1
          If EfectoList(TieneIndex).origen = origen Then
            cantidad_origen = cantidad_origen + 1
          End If
        End If
      Next
      
      If cantidad >= TmpEfecto.limite Then
        EfectoList(EfectoIndex).duracion = get_duracion_efecto(TmpEfecto.id)
        Call WriteConsoleMsg(TargetIndex, "MISMO: Se te renueva: " & EfectoList(EfectoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
        Exit Sub
      End If
      
      '15/12/2018 Irongete: Puede tener mas del mismo origen?
      If cantidad_origen >= TmpEfecto.limite_origen Then
        EfectoList(EfectoIndex).duracion = get_duracion_efecto(TmpEfecto.id)
        Call WriteConsoleMsg(TargetIndex, "MISMO ORIGEN: Se te renueva: " & EfectoList(EfectoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
        Exit Sub
      End If
      
      '16/12/2018 Irongete: Puede tener otro, lo creo y se lo pongo.
      ReDim Preserve EfectoList(UBound(EfectoList) + 1)
      Dim NuevoIndex As Long
      NuevoIndex = UBound(EfectoList)
      EfectoList(NuevoIndex) = TmpEfecto
      UserList(TargetIndex).Efecto.Add NuevoIndex, "efecto" & NuevoIndex
      Call WriteConsoleMsg(TargetIndex, "Se te aplica: " & EfectoList(NuevoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
      
      '15/12/2018 Irongete: Mandar el paquete CrearEfecto al jugador
      If EfectoList(NuevoIndex).enviar_a_cliente = True Then
        Call WriteCrearEfecto(TargetIndex, NuevoIndex)
      End If
        
      Exit Sub
    Else
      Call ejecutar_efecto(TriggerIndex, origen, TargetIndex)
    End If
  Next
  
  
End Sub

Public Sub procesar_efectos_jugador()
  Dim UserIndex As Integer
  '14/12/2018 Irongete: Efectos que tienen los jugadores
  UserIndex = 1
  For UserIndex = 1 To LastUser
      If UserList(UserIndex).ConnID <> -1 Then
          If UserList(UserIndex).flags.UserLogged Then
              Dim EfectoIndex As Variant
              For Each EfectoIndex In UserList(UserIndex).Efecto
                Call ejecutar_efecto(EfectoIndex, "efecto" & EfectoIndex, UserIndex)
              Next
          End If
      End If
  Next UserIndex
End Sub

Public Sub procesar_efectos_npc()
End Sub

Public Sub procesar_efectos_suelo()
  '16/12/2018 Irongete: Para cada zona
  Dim ZonaIndex As Variant
  For ZonaIndex = 0 To UBound(ZonaList)
  
    '16/12/2018 Irongete: Para cada jugador dentro de la zona
    Dim jugador As Variant
    For Each jugador In ZonaList(ZonaIndex).jugador
    
      '16/12/2018 Irongete: Para cada efecto que tenga la zona al estar dentro de ella
      Debug.Print ZonaList(ZonaIndex).nombre
      
      Dim EfectoIndex As Variant
      For Each EfectoIndex In ZonaList(ZonaIndex).efecto_al_estar 'para cada efecto
        Call ejecutar_efecto(EfectoIndex, "zona_dentro" & ZonaIndex, ZonaList(ZonaIndex).jugador("jugador" & jugador))
      Next
    Next
  Next
End Sub

Public Function get_duracion_efecto(ByVal EfectoId As Long)
  Dim EfectoIndex As Long
  For EfectoIndex = 0 To UBound(Efecto)
    If Efecto(EfectoIndex).id = EfectoId Then
      get_duracion_efecto = Efecto(EfectoIndex).duracion
    End If
  Next
End Function

Public Function crear_efecto(ByVal EfectoId As Long) As EfectoInfo
  Dim EfectoIndex As Long
  For EfectoIndex = 0 To UBound(Efecto)
    If Efecto(EfectoIndex).id = EfectoId Then
      crear_efecto = Efecto(EfectoIndex)
    End If
  Next
End Function
