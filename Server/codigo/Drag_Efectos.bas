Attribute VB_Name = "Drag_Efectos"
Public Enum EfectoTipo
  daño 'causa daño
  curacion 'causa curación
  aumenta_velocidad
  disminuye_velocidad
  
  'relacionado con magias
  mod_daño_magico
  
  'relacionado con zonas
  crea_zona
End Enum

Public Type EfectoValor
 tipo_valor As EfectoTipoValor
 valor As Long
End Type

Public Enum EfectoTipoValor
  daño = 0
  curacion = 1
  velocidad = 2
  area = 3
  mod_daño_magico = 4
  posicion_mundo = 5
  grh = 6
  Duracion = 7
  añadir_efecto_al_entrar = 8
  añadir_efecto_al_estar = 9
  añadir_efecto_al_moverse = 10
  añadir_efecto_al_salir = 11
  intervalo = 12
End Enum

Public Type EfectoInfo
  id As Integer
  nombre As String
  tipo As Byte
  descripcion As String
  valor() As EfectoValor
  trigger As New Collection 'sirve esto para que un efecto pueda triggear otros efectos??
  Duracion As Integer 'duración en milisegundos
  intervalo As Integer 'cada cuantos milisegundos se ejecuta el efecto
  limite As Byte 'cantidad limite de este efecto que puedes tener
  limite_origen As Byte 'cantidad limite de este efecto que puede tener por cada mismo origen (jugador, zona, npc)
  beneficioso As Boolean 'si es beneficioso es un buff de lo contrario es un debuff
  grh As Long
  enviar_a_cliente As Boolean
  aplicado As Boolean
  
  
  'valores que usa el server
  es_habilidad As Boolean
  habilidad As Integer
  'tipo_dueño As TipoDueño 'esto es para poner que tipo de cosa ha lanzado la habilidad que ha creado este efecto (jugador, npc, entorno)
  dueño As Long 'esto lo pongo para ponerle el UserIndex o NpcIndex del que crea el efecto usando una habilidad
  lanzador As Integer
  'tipo_objetivo As habilidad_objetivo
  objetivo As Integer
  origen As String 'de donde viene este efecto (jugador, zona entrar/salir/mover, npc...)
  contador_intervalo As Integer 'contador de los milisegundos
  'EfectoIndex As Integer 'esto guarda en el cliente el Indice del array EfectoList en donde está su efecto en el server por si se lo quita
  posicion As worldPos
  
End Type

Public efecto() As EfectoInfo 'aquí están todos los efectos con su configuración
Public EfectoList() As EfectoInfo 'aquí están los efectos que se crean y se asignan a jugadores y npcs, ...

Option Explicit






Public Sub Añadir_Efecto_a_Jugador(ByVal EfectoIndex As Long, ByVal UserIndex As Integer, ByVal origen As String)


        '17/12/2018 Irongete: Puede tener mas?
        Dim cantidad As Byte
        Dim cantidad_origen As Byte
        Dim TieneIndex As Variant
        
        cantidad = 0
        cantidad_origen = 0
        For Each TieneIndex In UserList(UserIndex).efecto 'recorro todos sus efectos para ver si cumple los limites
                
          '17/12/2018 Irongete: Ya lo tiene
          If EfectoList(TieneIndex).id = EfectoList(EfectoIndex).id Then
            cantidad = cantidad + 1
            
            '17/12/2018 Irongete: Puede tener otro?
            If cantidad >= EfectoList(TieneIndex).limite Then
               EfectoList(TieneIndex).Duracion = Get_Duracion_Efecto(EfectoList(TieneIndex).id)
              Call WriteConsoleMsg(UserIndex, "LIMITE - MISMO: Se te renueva: " & EfectoList(TieneIndex).nombre, FontTypeNames.FONTTYPE_INFO)
              Exit Sub
            End If
              
            '17/12/2018 Ironete: Es del mismo origen?
            If EfectoList(TieneIndex).origen = origen Then
            cantidad_origen = cantidad_origen + 1
            
              '17/12/2018 Irongete: Puede tener otro?
              If cantidad_origen >= EfectoList(TieneIndex).limite_origen Then
                 EfectoList(TieneIndex).Duracion = Get_Duracion_Efecto(EfectoList(TieneIndex).id)
                Call WriteConsoleMsg(UserIndex, "LIMITE - MISMO ORIGEN: Se te renueva: " & EfectoList(TieneIndex).nombre, FontTypeNames.FONTTYPE_INFO)
                Exit Sub
              End If
              
            End If
          End If
         Next
        
        
        '17/12/2018 Irongete: Lo copio para el
        ReDim Preserve EfectoList(UBound(EfectoList) + 1)
        Dim NuevoEfectoIndex As Long
        NuevoEfectoIndex = UBound(EfectoList)
        EfectoList(NuevoEfectoIndex) = EfectoList(EfectoIndex)
        EfectoList(NuevoEfectoIndex).origen = origen
        
        '17/12/2018 Irongete: Se lo aplico al jugador
        UserList(UserIndex).efecto.Add NuevoEfectoIndex, "efecto" & NuevoEfectoIndex
        Call WriteConsoleMsg(UserIndex, NuevoEfectoIndex & " Se te aplica: " & EfectoList(NuevoEfectoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
        
        '15/12/2018 Irongete: Mandar el paquete CrearEfecto al jugador
        If EfectoList(EfectoIndex).enviar_a_cliente = True Then
          Call WriteCrearEfecto(UserIndex, NuevoEfectoIndex)
          Debug.Print "WRITE"; UserIndex; NuevoEfectoIndex
        End If

  
End Sub


Public Function Añadir_Efecto_a_Zona(ByVal ZonaIndex As Long, ByVal EfectoId As Integer, ByVal tipo As Byte, Optional ByVal dueño As Long, Optional ByVal habilidad As Long)
     
  '16/12/2018 Irongete: Recorro Efecto() buscando la Id
  Dim EfectoIndex As Long
  For EfectoIndex = 0 To UBound(efecto)
  
    '16/12/2018 Irongete: Lo encontré
    If efecto(EfectoIndex).id = EfectoId Then
    
      '16/12/2018 Irongete: Añado el efecto a EfectoList
      Dim i As Long
      ReDim Preserve EfectoList(UBound(EfectoList) + 1)
      i = UBound(EfectoList)
      EfectoList(i) = efecto(EfectoIndex)
      
      '19/12/2018 Irongete: Si hay dueño se lo pongo al efecto para que tenga dueño y poder hacer que el efecto de esta zona afecte
      ' a segun que jugadores (enemigos, guild, etc...)
      If dueño > 0 Then
        EfectoList(i).dueño = dueño
      End If
      
      '19/12/2018 Irongete: si hay habilidad se lo pongo al efecto para el mensaje
      If habilidad > -1 Then
        EfectoList(i).habilidad = habilidad
      End If
      
      
      
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
  
  Añadir_Efecto_a_Zona = i

End Function


Public Sub Cargar_Efectos_SQL()

  ReDim efecto(0) As EfectoInfo
  ReDim EfectoList(0) As EfectoInfo
  
  '17/12/2018 Irongete: Cargo los efectos desde la base de datos
  Call check_sql
  Dim RS As ADODB.Recordset
  Set RS = New ADODB.Recordset
  Set RS = SQL.Execute("SELECT id, nombre, tipo, descripcion, duracion, intervalo, limite, limite_origen, beneficioso, grh, aplicado, enviar_a_cliente FROM efecto")
  
  Dim i As Long
  While Not RS.EOF
    '17/12/2018 Irongete: Datos del hechizo
    i = UBound(efecto)
    
    efecto(i).id = RS!id
    efecto(i).nombre = RS!nombre
    efecto(i).tipo = RS!tipo
    efecto(i).descripcion = RS!descripcion
    
    ReDim efecto(i).valor(0) As EfectoValor
    
    '18/12/2018 Irongete: Los efectos deben tener multiples valores
    'efecto(i).valor = RS!valor
    'meto los valores en un array
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    Set RS2 = SQL.Execute("SELECT valor, tipo FROM rel_efecto_valor WHERE id_efecto = '" & RS!id & "'")
    
    While Not RS2.EOF
      Dim e As Byte
      e = UBound(efecto(i).valor)
      
      'valores efectos que tienen un valor único (daño, curacion, mod_daño_magico....)
      efecto(i).valor(e).tipo_valor = RS2!tipo
      efecto(i).valor(e).valor = RS2!valor
      
      ReDim Preserve efecto(i).valor(UBound(efecto(i).valor) + 1)
      RS2.MoveNext
    Wend
    RS2.Close
    
    efecto(i).Duracion = RS!Duracion
    efecto(i).intervalo = RS!intervalo
    efecto(i).limite = RS!limite
    efecto(i).limite_origen = RS!limite_origen
    efecto(i).beneficioso = RS!beneficioso
    efecto(i).grh = RS!grh
    efecto(i).aplicado = RS!aplicado
    efecto(i).enviar_a_cliente = RS!enviar_a_cliente
    
    '17/12/2018 Irongete: Triggers
    Set RS2 = New ADODB.Recordset
    Set RS2 = SQL.Execute("SELECT `trigger` FROM rel_efecto_trigger WHERE id_efecto = '" & RS!id & "'")
    
    While Not RS2.EOF
      efecto(i).trigger.Add RS2!trigger, "trigger" & RS2!trigger
      RS2.MoveNext
    Wend
    RS2.Close
    Set RS2 = Nothing
    
    ReDim Preserve efecto(i + 1) As EfectoInfo
    RS.MoveNext
  Wend
  RS.Close
  Set RS = Nothing
  
     
  'test velocidad
  Call Añadir_Efecto_a_Zona(6, 1, ZonaEvento.Entrar) 'aumentar velocidad al entrar
  Call Añadir_Efecto_a_Zona(6, 2, ZonaEvento.Salir) 'disminuir velocidad al salir
  
  'test aumento daño magico
  Call Añadir_Efecto_a_Zona(7, 5, ZonaEvento.Entrar) 'meter buff de daño magico
  
  
  '16/12/2018 Irongete: Todo cargado, activo el timer.
  frmMain.TimerEfectos.Enabled = True
End Sub

Public Function Tiene_Efecto(ByVal EfectoId As Integer, ByVal UserIndex As Integer) As Boolean
  Dim EfectoIndex As Variant
  For Each EfectoIndex In UserList(UserIndex).efecto
    If EfectoList(EfectoIndex).id = EfectoId Then
      Tiene_Efecto = True
    End If
  Next
End Function


Public Function Cantidad_Efecto(ByVal efecto_id As Integer, ByVal UserIndex As Integer)
  Dim cantidad As Byte
  cantidad = 0
  Dim i As Variant
  For Each i In UserList(UserIndex).efecto
    If EfectoList(i).id = efecto_id Then
      cantidad = cantidad + 1
    End If
  Next
Cantidad_Efecto = cantidad
End Function

Public Function Cantidad_Efecto_Mismo_Origen(ByVal efecto_id As Integer, ByVal UserIndex As Integer, ByVal origen As String)
  Dim cantidad As Byte
  cantidad = 0
  Dim i As Variant
  For Each i In UserList(UserIndex).efecto
    If EfectoList(i).id = efecto_id Then
      If EfectoList(i).origen = origen Then
        cantidad = cantidad + 1
      End If
    End If
  Next
Cantidad_Efecto_Mismo_Origen = cantidad
End Function


Public Function Get_Efecto_Valor(ByVal tipo_valor As EfectoTipoValor, ByVal EfectoIndex As Long)
  Dim ValorIndex As Byte
  
  For ValorIndex = 0 To UBound(EfectoList(EfectoIndex).valor)
    If EfectoList(EfectoIndex).valor(ValorIndex).tipo_valor = tipo_valor Then
      Get_Efecto_Valor = EfectoList(EfectoIndex).valor(ValorIndex).valor
      Exit For
    End If
  Next
  Get_Efecto_Valor = Get_Efecto_Valor

End Function


Public Sub Ejecutar_Efecto(ByVal EfectoIndex As Long, TargetIndex As Long)

  Debug.Print "EJECUTO EFECTO "; EfectoIndex; "EN "; TargetIndex
  
  
  Exit Sub

End Sub







Public Sub Ejecutar_Efecto_Old(ByVal EfectoIndex As Integer, TargetIndex As Long)
  Dim valor As Variant
  Dim i As Integer
  
  If TargetIndex = 0 And Not EfectoList(EfectoIndex).tipo = 5 Then Exit Sub
  
  
  '16/12/2018 Irongete: Es aplicado? Controlamos la duracion...
  If EfectoList(EfectoIndex).aplicado = True Then
  
    '14/12/2018 Irongete: Se ha acabado el efecto?
    If EfectoList(EfectoIndex).Duracion >= 0 Then
      EfectoList(EfectoIndex).Duracion = EfectoList(EfectoIndex).Duracion - 100
      EfectoList(EfectoIndex).contador_intervalo = EfectoList(EfectoIndex).contador_intervalo + 100
  
      '15/12/2018 Irongete: El intervalo permite que se ejecute?
      If Not EfectoList(EfectoIndex).contador_intervalo >= EfectoList(EfectoIndex).intervalo Then
        Exit Sub
      End If
      EfectoList(EfectoIndex).contador_intervalo = 0
      
    '17/12/2018 Irongete: La duracion ha llegado a 0, quitamos el efecto
    Else
      '17/12/2018 Irongete: Solo si el objetivo es jugador
      If EfectoList(EfectoIndex).tipo_objetivo = HabilidadObjetivo.jugador Then
        UserList(TargetIndex).efecto.Remove "efecto" & EfectoIndex
        Call WriteConsoleMsg(TargetIndex, "Se te quita: " & EfectoList(EfectoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
        
        '15/12/2018 Irongete: Mandar el paquete QuitarEfecto al jugador
        If EfectoList(EfectoIndex).enviar_a_cliente = True Then
          Call WriteQuitarEfecto(TargetIndex, EfectoIndex)
        End If
      End If
      
      '17/12/2018 Irongete: Se lo quito al NPC
      If EfectoList(EfectoIndex).tipo_objetivo = HabilidadObjetivo.npc Then
        NPCList(TargetIndex).efecto.Remove "efecto" & EfectoIndex
      End If
      Exit Sub
    End If
  End If
  '------------------------ Fin de la comprobación del efecto si es aplicado
  
    
  
  Select Case EfectoList(EfectoIndex).tipo
  
    '18/12/2018 Irongete: El efecto crea una zona
    Case EfectoTipo.crea_zona ' ***** ZONA
      Dim area As Byte
      Dim grh As Long
      Dim Pos As worldPos
      Dim Duracion As Long
      
            
      Pos = EfectoList(EfectoIndex).posicion
      area = Get_Efecto_Valor(EfectoTipoValor.area, EfectoIndex)
      grh = Get_Efecto_Valor(EfectoTipoValor.grh, EfectoIndex)
      Duracion = Get_Efecto_Valor(EfectoTipoValor.Duracion, EfectoIndex)
      
            
      Dim ZonaIndex As Long
      ZonaIndex = crear_zona() 'creo una nueva zona
  
      ZonaList(ZonaIndex).nombre = "zona generada por Efecto" & EfectoIndex
      ZonaList(ZonaIndex).Mapa = Pos.Map
      ZonaList(ZonaIndex).x1 = Pos.X - area
      ZonaList(ZonaIndex).x2 = Pos.X + area
      ZonaList(ZonaIndex).y1 = Pos.Y - area
      ZonaList(ZonaIndex).y2 = Pos.Y + area
      ZonaList(ZonaIndex).grh = grh
      ZonaList(ZonaIndex).permisos = 0
      ZonaList(ZonaIndex).Duracion = Duracion
      ZonaList(ZonaIndex).temporal = True
     
      '18/12/2018 Irongete: Añado a las zonas los efectos al entrar, moverse, estar o salir
      Dim EfectoValorIndex As Long
      Dim NuevoEfectoIndex As Long
      For EfectoValorIndex = 0 To UBound(EfectoList(EfectoIndex).valor)
      
        '18/12/2018 Irongete: Le pongo a la variable dueño del efecto el UserIndex de quien ha creado la zona
        Select Case EfectoList(EfectoIndex).valor(EfectoValorIndex).tipo_valor
          Case EfectoTipoValor.añadir_efecto_al_entrar
            NuevoEfectoIndex = Añadir_Efecto_a_Zona(ZonaIndex, EfectoList(EfectoIndex).valor(EfectoValorIndex).valor, ZonaEvento.Entrar)
          Case EfectoTipoValor.añadir_efecto_al_estar
            NuevoEfectoIndex = Añadir_Efecto_a_Zona(ZonaIndex, EfectoList(EfectoIndex).valor(EfectoValorIndex).valor, ZonaEvento.Estar)
          Case EfectoTipoValor.añadir_efecto_al_moverse
            NuevoEfectoIndex = Añadir_Efecto_a_Zona(ZonaIndex, EfectoList(EfectoIndex).valor(EfectoValorIndex).valor, ZonaEvento.Moverse)
          Case EfectoTipoValor.añadir_efecto_al_salir
            NuevoEfectoIndex = Añadir_Efecto_a_Zona(ZonaIndex, EfectoList(EfectoIndex).valor(EfectoValorIndex).valor, ZonaEvento.Salir)
          Case EfectoTipoValor.intervalo
            EfectoList(EfectoIndex).intervalo = EfectoList(EfectoIndex).valor(EfectoValorIndex).valor
        End Select
        
        
        '19/12/2018 Irongete: Añadir propiedades al efecto
        EfectoList(NuevoEfectoIndex).habilidad = EfectoList(EfectoIndex).habilidad 'hereda la habilidad
        EfectoList(NuevoEfectoIndex).dueño = EfectoList(EfectoIndex).dueño 'hereda el dueño
        
        
        Debug.Print "INTERVALO"; EfectoList(EfectoIndex).intervalo
        
      Next
      
      Call añadir_zona_a_mapa(Pos.Map, ZonaIndex)
      
      '19/12/2018 Irongete: Envio la zona a todos los jugadores del mapa
      Dim EnviarZonaIndex As Integer
      For EnviarZonaIndex = 1 To LastUser
        If UserList(EnviarZonaIndex).ConnIDValida = True Then
          Call WriteCrearZona(EnviarZonaIndex, ZonaIndex)
        End If
      Next
              


    '17/12/2018 Irongete: El efecto hace daño
    Case EfectoTipo.daño '***** DAÑO
      Dim daño As Long
      
      daño = Get_Efecto_Valor(EfectoTipoValor.daño, EfectoIndex)
      
      If TargetIndex = 0 Then
        TargetIndex = EfectoList(EfectoIndex).objetivo
      End If
      
      '17/12/2018 Irongete: El daño se modifica segun los efectos que tenga
      Dim ModDaño As Variant
      For Each ModDaño In UserList(EfectoList(EfectoIndex).dueño).efecto
        
        '17/12/2018 Irongete: Aumenta daño magico?
        If EfectoList(ModDaño).tipo = EfectoTipo.mod_daño_magico Then
          Dim Aumento As Long
          Aumento = Get_Efecto_Valor(EfectoTipo.mod_daño_magico, ModDaño)
          daño = daño + Aumento
        End If
        
      Next
      
      '17/12/2018 Irongete: A quien le pegamos?
      Select Case EfectoList(EfectoIndex).tipo_objetivo
      
        '17/12/2018 Irongete: A un jugador
        Case HabilidadObjetivo.jugador
          
          '19/12/2018 Irongete: Es una habilidad, mostramos el nombre de la habilidad
          If EfectoList(EfectoIndex).habilidad > 0 Then
            Call WriteConsoleMsg(TargetIndex, EfectoIndex & " " & habilidad(EfectoList(EfectoIndex).habilidad).nombre & " de " & UserList(EfectoList(EfectoIndex).dueño).Name & " te causa " & daño & " de daño.", FontTypeNames.FONTTYPE_INFO)
          Else
            Call WriteConsoleMsg(TargetIndex, EfectoIndex & " " & EfectoList(EfectoIndex).nombre & " de " & UserList(EfectoList(EfectoIndex).dueño).Name & " te causa " & daño & " de daño.", FontTypeNames.FONTTYPE_INFO)
          End If
          
        '17/12/2028 Irongete: A un npc
        Case HabilidadObjetivo.npc
          Dim NpcIndex As Integer
          NpcIndex = TargetIndex
          
  
          '17/12/2018 Irongete: Le quito vida al NPC
          NPCList(NpcIndex).Stats.MinHP = NPCList(TargetIndex).Stats.MinHP - daño
          
          '17/12/2018 Irongete: Mensaje
          Call WriteConsoleMsg(EfectoList(EfectoIndex).dueño, EfectoIndex & " Tu " & habilidad(EfectoList(EfectoIndex).habilidad).nombre & " causa " & daño & " daño a " & NPCList(NpcIndex).Name & "(" & NPCList(NpcIndex).Stats.MinHP & "/" & NPCList(NpcIndex).Stats.MaxHP & ")", FontTypeNames.FONTTYPE_FIGHT)
                    
          '17/12/2018 Irongete: El mensaje del daño encima del NPC
          Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateRenderValue(NPCList(NpcIndex).Pos.X, NPCList(NpcIndex).Pos.Y, daño, DAMAGE_NORMAL))
          

          '17/12/2018 Irongete: Muere el NPC?
          If NPCList(NpcIndex).Stats.MinHP <= 0 Then
            Call MuereNpc(NpcIndex, EfectoList(EfectoIndex).dueño)
            
            '17/12/2018 Irongete: Le quito los posibles efectos que pueda tener
            Dim PosibleEfecto As Variant
            For Each PosibleEfecto In NPCList(NpcIndex).efecto
              NPCList(NpcIndex).efecto.Remove "efecto" & PosibleEfecto
            Next
          End If
      
      
      End Select
      
      
    '17/12/2018 Irongete: El efecto causa curacion
    Case EfectoTipo.curacion
      Call WriteConsoleMsg(TargetIndex, EfectoList(EfectoIndex).nombre & " te restaura " & 1 & " de salud.", FontTypeNames.FONTTYPE_INFO)
      
    Case EfectoTipo.aumenta_velocidad
      UserList(TargetIndex).flags.Speed = UserList(TargetIndex).flags.Speed + 1
      Call WriteChangeSpeed(TargetIndex, UserList(TargetIndex).Char.CharIndex, UserList(TargetIndex).flags.Speed)
      Call WriteConsoleMsg(TargetIndex, EfectoIndex & " " & "Se te aumenta la velocidad en " & 1, FontTypeNames.FONTTYPE_INFO)
            
     Case EfectoTipo.disminuye_velocidad
      UserList(TargetIndex).flags.Speed = UserList(TargetIndex).flags.Speed - 1
      Call WriteChangeSpeed(TargetIndex, UserList(TargetIndex).Char.CharIndex, UserList(TargetIndex).flags.Speed)
      Call WriteConsoleMsg(TargetIndex, EfectoIndex & " " & "Se te disminuye la velocidad en " & 1, FontTypeNames.FONTTYPE_INFO)
      
    Case EfectoTipo.mod_daño_magico
      Debug.Print "DAÑO MAGICOOO"
      
  End Select
  
  
  
  
  
  
  '16/12/2018 Irongete: El efecto triggerea otros efectos?
  Dim TriggerIndex As Variant
  For Each TriggerIndex In EfectoList(EfectoIndex).trigger
  
    Dim TmpEfecto As EfectoInfo
    TmpEfecto = Crear_Efecto(TriggerIndex) 'este es el efecto que se le va a aplicar
    
    '16/12/2018 Irongete: Es aplicado?
    If TmpEfecto.aplicado = True Then
    
      '15/12/2018 Irongete: Puede tener mas?
      Dim TieneIndex As Variant
      Dim cantidad As Byte
      Dim cantidad_origen As Byte
      For Each TieneIndex In UserList(TargetIndex).efecto 'recorro todos sus efectos para ver si cumple los limites
        If EfectoList(TieneIndex).id = TmpEfecto.id Then
          cantidad = cantidad + 1
          If EfectoList(TieneIndex).origen = origen Then
            cantidad_origen = cantidad_origen + 1
          End If
        End If
      Next
      
      If cantidad >= TmpEfecto.limite Then
        EfectoList(EfectoIndex).Duracion = Get_Duracion_Efecto(TmpEfecto.id)
        Call WriteConsoleMsg(TargetIndex, "MISMO: Se te renueva: " & EfectoList(EfectoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
        Exit Sub
      End If
      
      '15/12/2018 Irongete: Puede tener mas del mismo origen?
      If cantidad_origen >= TmpEfecto.limite_origen Then
        EfectoList(EfectoIndex).Duracion = Get_Duracion_Efecto(TmpEfecto.id)
        Call WriteConsoleMsg(TargetIndex, "MISMO ORIGEN: Se te renueva: " & EfectoList(EfectoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
        Exit Sub
      End If
      
      '16/12/2018 Irongete: Puede tener otro, lo creo y se lo pongo.
      ReDim Preserve EfectoList(UBound(EfectoList) + 1)
      Dim NuevoIndex As Long
      NuevoIndex = UBound(EfectoList)
      EfectoList(NuevoIndex) = TmpEfecto
      UserList(TargetIndex).efecto.Add NuevoIndex, "efecto" & NuevoIndex
      Call WriteConsoleMsg(TargetIndex, NuevoIndex & "Se te aplica: " & EfectoList(NuevoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
      
      '15/12/2018 Irongete: Mandar el paquete CrearEfecto al jugador
      If EfectoList(NuevoIndex).enviar_a_cliente = True Then
        Call WriteCrearEfecto(TargetIndex, NuevoIndex)
      End If
        
      Exit Sub
    Else
      Call Ejecutar_Efecto(TriggerIndex, origen, TargetIndex)
    End If
  Next
  
  
  
  
  
End Sub
'#End Region


Public Sub Procesar_Efectos_Jugador()
  Dim UserIndex As Integer
  '14/12/2018 Irongete: Efectos que tienen los jugadores
  UserIndex = 1
  For UserIndex = 1 To LastUser
      If UserList(UserIndex).ConnID <> -1 Then
          If UserList(UserIndex).flags.UserLogged Then
              Dim EfectoIndex As Variant
              For Each EfectoIndex In UserList(UserIndex).efecto
                Dim TargetIndex As Long
                TargetIndex = EfectoIndex
                Call Ejecutar_Efecto(EfectoIndex, TargetIndex)
              Next
          End If
      End If
  Next UserIndex
End Sub

Public Sub Procesar_Efectos_NPC()
  Dim NpcIndex As Integer
  '17/12/2018 Irongete:
  NpcIndex = 1
  For NpcIndex = 1 To UBound(NPCList)
    Dim EfectoIndex As Variant
    For Each EfectoIndex In NPCList(NpcIndex).efecto
      Dim TargetIndex As Long
      TargetIndex = EfectoIndex
      Call Ejecutar_Efecto(EfectoIndex, TargetIndex)
    Next
  Next
End Sub

Public Sub Procesar_Efectos_Suelo()
  '16/12/2018 Irongete: Para cada zona
  Dim ZonaIndex As Variant
  For ZonaIndex = 0 To UBound(ZonaList)
    
      '16/12/2018 Irongete: Para cada efecto que tenga la zona al estar dentro de ella
      Dim EfectoIndex As Variant
      For Each EfectoIndex In ZonaList(ZonaIndex).efecto_al_estar 'para cada efecto
      
         '16/12/2018 Irongete: Para cada jugador dentro de la zona
        Dim JugadorIndex As Long
        For JugadorIndex = 0 To UBound(ZonaList(ZonaIndex).jugador)
         If JugadorIndex > 0 Then
           Call Ejecutar_Efecto(EfectoIndex, ZonaList(ZonaIndex).jugador(JugadorIndex))
        End If
        
      Next
      
      
    Next
  Next
End Sub

Public Function Get_Duracion_Efecto(ByVal EfectoId As Long)
  Dim EfectoIndex As Long
  For EfectoIndex = 0 To UBound(efecto)
    If efecto(EfectoIndex).id = EfectoId Then
      Get_Duracion_Efecto = efecto(EfectoIndex).Duracion
    End If
  Next
End Function

Public Function Crear_Efecto(ByVal EfectoId As Variant) As EfectoInfo
  Dim EfectoIndex As Long
  For EfectoIndex = 0 To UBound(efecto)
    If efecto(EfectoIndex).id = EfectoId Then
      Crear_Efecto = efecto(EfectoIndex)
    End If
  Next
End Function
