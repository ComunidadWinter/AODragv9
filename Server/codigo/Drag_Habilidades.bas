Attribute VB_Name = "Drag_Habilidades"
Public Type habilidad
  id As Long 'id unica de la habilidad
  nombre As String 'nombre
  palabras_magicas As String 'el texto que sale al lanzarla
  objetivo As tipo_objetivo 'jugador, npc, suelo....
  beneficiosa As Boolean
  'cambio esto por beneficiosa 'solo_amigo As Boolean 'el objetivo (jugador, npc, suelo(zona?)) ha de ser amigo (clan, grupo...)
  'es mejor que el efecto decida si es buff o debuff??? -- tipo As habilidad_tipo 'negativa/positiva. segun esto, si es aplicable saldrá como buff o debuff dependiendo de esto.
  efecto As New Collection 'efectos que ejecuta esta habilidad
  fx As Long 'grh de la animación
  wav As Integer 'numero del sonido
End Type

Option Explicit

Public habilidad() As habilidad


Public Sub cargar_habilidades_sql()

  '20/12/2018 Irongete: Dimensionar el array
  ReDim habilidad(1) As habilidad

  '17/12/2018 Irongete: Cargo los efectos desde la base de datos
  Call check_sql
  Dim RS As ADODB.Recordset
  Set RS = New ADODB.Recordset
  Set RS = SQL.Execute("SELECT id, nombre, palabras_magicas, objetivo, beneficiosa, fx, wav FROM habilidad")
  
  Dim habilidad_index As Long
  While Not RS.EOF
  
    '17/12/2018 Irongete: Datos de la habilidad
    habilidad_index = UBound(habilidad)
    
    habilidad(habilidad_index).id = RS!id
    habilidad(habilidad_index).nombre = RS!nombre
    habilidad(habilidad_index).palabras_magicas = RS!palabras_magicas
    habilidad(habilidad_index).objetivo = RS!objetivo
    habilidad(habilidad_index).beneficiosa = RS!beneficiosa
    habilidad(habilidad_index).fx = RS!fx
    habilidad(habilidad_index).wav = RS!wav

    '17/12/2018 Irongete: Triggers
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    Set RS2 = SQL.Execute("SELECT id_efecto FROM rel_habilidad_efecto WHERE id_habilidad = '" & RS!id & "'")

    While Not RS2.EOF
      Dim id_efecto As Integer
      id_efecto = RS2!id_efecto
      habilidad(habilidad_index).efecto.Add id_efecto, "efecto" & id_efecto
      RS2.MoveNext
    Wend
    RS2.Close
    Set RS2 = Nothing
    
    
    ReDim Preserve habilidad(habilidad_index + 1) As habilidad
    RS.MoveNext
  Wend
  RS.Close
  Set RS = Nothing
  

  
End Sub

Public Sub lanzar_habilidad(ByVal habilidad_id As Integer, ByVal UserIndex As Integer)
  
  '17/12/2018 Irongete: Busco la habilidad
  Dim habilidad_index As Integer

  For habilidad_index = 0 To UBound(habilidad)
  
    If habilidad(habilidad_index).id = habilidad_id Then
      Dim TargetIndex As Long
  
      '20/12/2018 Irongete: Comprobar si tiene la habilidad aprendida.
      ' - es necesario? supongo que si...



      '**************************************************************************************
      '******************************************* JUGADOR **************************************
      '**************************************************************************************
      If (habilidad(habilidad_index).objetivo And tipo_objetivo.jugador) And UserList(UserIndex).flags.targetUser > 0 Then
      
        '20/12/2018 Irongete: Quien es el jugador?
        TargetIndex = UserList(UserIndex).flags.targetUser
      
        '20/12/2018 Irongete: Si no es beneficiosa no puede tirarsela a si mismo
        If (habilidad(habilidad_index).beneficiosa = False And UserIndex = TargetIndex) Then
          Call WriteConsoleMsg(UserIndex, "a ti no", FontTypeNames.FONTTYPE_INFO)
  
        Else
          
          Call WriteConsoleMsg(UserIndex, "LANZAS A " & UserList(TargetIndex).Name, FontTypeNames.FONTTYPE_INFO)
          
          '20/12/2018 Irongete: Creo el efecto y se lo meto al target
          Dim efecto_id As Variant

          For Each efecto_id In habilidad(habilidad_index).efecto
          
            Dim TmpEfecto As EfectoInfo
            TmpEfecto = crear_efecto(efecto_id)
            
            ReDim Preserve EfectoList(UBound(EfectoList) + 1)
            Dim EfectoIndex As Long
            EfectoIndex = UBound(EfectoList)
            EfectoList(EfectoIndex) = TmpEfecto
            
            '20/12/2018 Irongete: Le pongo info al efecto para ejecutarlo
            EfectoList(EfectoIndex).dueño = UserIndex
           
            '20/12/2018 Irongete: Si el efecto es aplicado se lo meto al jugador
            If EfectoList(EfectoIndex).aplicado = True Then
            
            Else
              Call ejecutar_efecto(EfectoIndex, TargetIndex)
            End If
            
            Call WriteConsoleMsg(UserIndex, EfectoIndex & " EFECTO" & EfectoList(EfectoIndex).nombre, FontTypeNames.FONTTYPE_INFO)
          Next
       
        End If
        Exit Sub
      End If
    
    
    
      '**************************************************************************************
      '************************************NPC **********************************************
      '**************************************************************************************
      If (habilidad(habilidad_index).objetivo And tipo_objetivo.npc) And UserList(UserIndex).flags.TargetNPC > 0 Then
    
        '20/12/2018 Irongete: Quien es el npc?
        TargetIndex = UserList(UserIndex).flags.TargetNPC
        
        '20/12/2018 Irongete: Puede lanzarselo a ese npc? no es su mascota, o de su amigo, o rey de su castillo, etc.....
        If habilidad(habilidad_index).beneficiosa = True Then Exit Sub
        
        '20/12/2018 Ironete: Si ha llegado aquí es que puede lanzarla
        Call WriteConsoleMsg(UserIndex, "LANZAS A NPC OK " & NPCList(TargetIndex).Name, FontTypeNames.FONTTYPE_INFO)
        
      End If
    End If
  Next
End Sub

Public Sub montar_follon()
  '17/12/2018 Irongete: Es una habilidad que lanza un jugador?
  If EfectoList(EfectoIndex).es_habilidad = True Then
  
    '17/12/2018 Irongete: Palabras magicas
    Dim palabras_magicas As String
    palabras_magicas = habilidad(EfectoList(EfectoIndex).habilidad).palabras_magicas
    Call SendData(SendTarget.ToPCArea, EfectoList(EfectoIndex).dueño, PrepareMessageChatOverHead(palabras_magicas, UserList(EfectoList(EfectoIndex).dueño).Char.CharIndex, vbCyan))
        
    '17/12/2018 Irongete: FX y sonido de la habilidad
    If EfectoList(EfectoIndex).tipo_objetivo = jugador Then
      Call SendData(SendTarget.ToNPCArea, EfectoList(EfectoIndex).objetivo, PrepareMessageCreateFX(UserList(EfectoList(EfectoIndex).objetivo).Char.CharIndex, habilidad(EfectoList(EfectoIndex).habilidad).fx, 0))
      Call SendData(SendTarget.ToNPCArea, EfectoList(EfectoIndex).objetivo, PrepareMessagePlayWave(habilidad(EfectoList(EfectoIndex).habilidad).wav, NPCList(EfectoList(EfectoIndex).objetivo).Pos.X, NPCList(EfectoList(EfectoIndex).objetivo).Pos.Y))
    End If

    If EfectoList(EfectoIndex).tipo_objetivo = npc Then
      Call SendData(SendTarget.ToNPCArea, EfectoList(EfectoIndex).objetivo, PrepareMessageCreateFX(NPCList(EfectoList(EfectoIndex).objetivo).Char.CharIndex, habilidad(EfectoList(EfectoIndex).habilidad).fx, 0))
      Call SendData(SendTarget.ToNPCArea, EfectoList(EfectoIndex).objetivo, PrepareMessagePlayWave(habilidad(EfectoList(EfectoIndex).habilidad).wav, NPCList(EfectoList(EfectoIndex).objetivo).Pos.X, NPCList(EfectoList(EfectoIndex).objetivo).Pos.Y))
    End If
      
  End If

End Sub

