Attribute VB_Name = "Acciones"

Option Explicit

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

'¿Posicion valida?
If InMapBounds(Map, X, Y) Then
   
    Dim foundchar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
       
    '¿Es un obj?
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        
        Select Case ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType
            
            Case OBJTYPE_PUERTAS 'Es una puerta
                Call AccionParaPuerta(Map, X, Y, UserIndex)
            Case OBJTYPE_CARTELES 'Es un cartel
                Call AccionParaCartel(Map, X, Y, UserIndex)
            Case OBJTYPE_FOROS 'Foro
                Call AccionParaForo(Map, X, Y, UserIndex)
            Case OBJTYPE_LEÑA 'Leña
                If MapData(Map, X, Y).OBJInfo.ObjIndex = FOGATA_APAG Then
                    Call AccionParaRamita(Map, X, Y, UserIndex)
                End If
            
        End Select
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.ObjIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
        End Select
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
        End Select
    ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.ObjIndex
        Call SendData(ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).Name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
            
        End Select
        
    Else
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        Call SendData(ToIndex, UserIndex, 0, "M9")
    End If
    
End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error GoTo fallo

Dim suerte As Byte
Dim exito As Byte
Dim obj As obj
Dim raise As Integer
'pluto:2.15
If UserList(UserIndex).flags.Muerto > 0 Then Exit Sub

If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
            suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
            suerte = 1
End If

exito = RandomNumber(1, suerte)

If exito = 1 Then
    obj.ObjIndex = FOGATA
    obj.Amount = 1
    
    Call SendData(ToIndex, UserIndex, 0, "||Has prendido la fogata." & "´" & FontTypeNames.FONTTYPE_info)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "FO")
    
    Call MakeObj(ToMap, 0, Map, obj, Map, X, Y)
    
    
Else
    Call SendData(ToIndex, UserIndex, 0, "||No has podido hacer fuego." & "´" & FontTypeNames.FONTTYPE_info)
End If

'Sino tiene hambre o sed quizas suba el skill supervivencia
If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
    Call SubirSkill(UserIndex, Supervivencia)
End If

Exit Sub
fallo:
Call LogError("ACCIONRAMITA " & Err.number & " D: " & Err.Description)

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error GoTo fallo

'¿Hay mensajes?
Dim f As String, tit As String, men As String, base As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(f, "INFO", "CantMSG"))
    base = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim n As Integer
    For i = 1 To num
        n = FreeFile
        f = base & i & ".for"
        Open f For Input Shared As #n
        Input #n, tit
        men = ""
        auxcad = ""
        Do While Not EOF(n)
            Input #n, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #n
        Call SendData2(ToIndex, UserIndex, 0, 52, tit & Chr(176) & men)
    Next
End If
Call SendData2(ToIndex, UserIndex, 0, 53)

Exit Sub
fallo:
Call LogError("ACCIONPARAFORO " & Err.number & " D: " & Err.Description)


End Sub


Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error GoTo fallo


Dim MiObj As obj
Dim wp As WorldPos
'pluto:hoy
Dim son As Integer
If Map > 177 Then son = 133 Else son = SND_PUERTA

If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 1 Then
        If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Cerrada = 1 And Cuentas(UserIndex).Llave = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Clave Then
            MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexAbierta
            Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)
            MapData(Map, X, Y).Blocked = 0
            MapData(Map, X - 1, Y).Blocked = 0
                     
            Call Bloquear(ToMap, 0, Map, Map, X, Y, 0)
            Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 0)
            SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & son
            Exit Sub
        End If
    End If
    
    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then
                     MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexAbierta
                                  
                     Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                     
                     'Desbloquea
                     MapData(Map, X, Y).Blocked = 0
                     MapData(Map, X - 1, Y).Blocked = 0
                     
                     'Bloquea todos los mapas
                     Call Bloquear(ToMap, 0, Map, Map, X, Y, 0)
                     Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 0)
                     
                       
                     'Sonido
                     SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & son
                    
                Else
                     Call SendData(ToIndex, UserIndex, 0, "E8")
                End If
        Else
                'Cierra puerta
                MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexCerradaLlave
                
                Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                
                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 1)
                Call Bloquear(ToMap, 0, Map, Map, X, Y, 1)
                
                SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & son
        End If
        
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
    Else
        Call SendData(ToIndex, UserIndex, 0, "E8")
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "L2")
End If

Exit Sub
fallo:
Call LogError("ACCIONPARAPUERTA " & Err.number & " D: " & Err.Description)

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error GoTo fallo

Dim MiObj As obj

If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = 8 Then
 
If ServerPrimario = 1 Then
 'pluto:6.3 textos carteles casas
 If MapData(Map, X, Y).OBJInfo.ObjIndex = 79 And Map = 59 Then
 ObjData(79).texto = "Emporio Los Ramones. ¡¡Tenemos soluciones!!"
 End If
 If MapData(Map, X, Y).OBJInfo.ObjIndex = 77 And Map = 64 Then
 ObjData(77).texto = "Steve"
 End If
 'pluto:6.9
  If MapData(Map, X, Y).OBJInfo.ObjIndex = 78 And Map = 34 Then
 ObjData(78).texto = "Tras advertir el lúgubre aspecto de la vivienda, te diriges hacia la inscripción hallada cerca del umbral y a duras penas logras leer su misiva.                                   Si aquí usted se amparare...             Estertor encontrare...                                                                    Alexstrasza Clamacielos. "
 End If
 
End If
 
  If Len(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto) > 0 Then
       Call SendData2(ToIndex, UserIndex, 0, 44, _
        ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto & _
        Chr(176) & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhSecundario)
  End If

End If

Exit Sub
fallo:
Call LogError("ACCIONPARACARTEL " & Err.number & " D: " & Err.Description)


End Sub

