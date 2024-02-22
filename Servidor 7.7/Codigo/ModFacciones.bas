Attribute VB_Name = "ModFacciones"
Option Explicit

Public ArmaduraImperial1 As Integer 'Primer jerarquia
Public ArmaduraImperial2 As Integer 'Segunda jerarquía
Public ArmaduraImperial3 As Integer 'Enanos
Public TunicaMagoImperial As Integer 'Magos
Public TunicaMagoImperialEnanos As Integer 'Magos


Public ArmaduraCaos1 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer

Public ArmaduraLegion1 As Integer
Public TunicaMagoLegion As Integer
Public TunicaMagoLegionEnanos As Integer
Public ArmaduraLegion2 As Integer
Public ArmaduraLegion3 As Integer

Public Const ExpAlUnirse = 100000
Public Const ExpX100 = 100000


Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal = 2 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya perteneces a las tropas de la Legión !!! Ve a combatir criminales!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 2 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya has pertenecido a las tropas de la Legión, no puedes entrar a la Armada. !!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Maldito insolente!!! vete de aqui seguidor de las sombras!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6ºNo se permiten criminales en el ejercito imperial!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CriminalesMatados < 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Para unirte a nuestras fuerzas debes matar al menos 1 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 30 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Para unirte a nuestras fuerzas debes ser al menos de nivel 30!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 5 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Has asesinado más de 5 ciudadanos, no aceptamos asesinos en las tropas reales!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Faccion.ArmadaReal = 1
'pluto:2.4.7.
'UserList(userindex).Faccion.RecompensasReal = UserList(userindex).Faccion.CriminalesMatados \ 100
UserList(UserIndex).Faccion.RecompensasReal = 1

Call SendData(ToIndex, UserIndex, 0, "||6°Bienvenido a al Ejercito Imperial!!!, aqui tienes tu armadura. Por cada centena de criminales que acabes te dare un recompensa, buena suerte soldado!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As obj
    MiObj.Amount = 1
    If UCase$(UserList(UserIndex).clase) = "MAGO" Or UCase$(UserList(UserIndex).clase) = "DRUIDA" Then
           If UCase$(UserList(UserIndex).raza) = "ENANO" Or UCase$(UserList(UserIndex).raza) = "GNOMO" Or UCase$(UserList(UserIndex).raza) = "GOBLIN" Then
                  MiObj.ObjIndex = TunicaMagoImperialEnanos
           Else
                  MiObj.ObjIndex = TunicaMagoImperial
           If UCase$(UserList(UserIndex).Genero) = "MUJER" Then MiObj.ObjIndex = 516

           End If
    ElseIf UCase$(UserList(UserIndex).clase) = "GUERRERO" Or _
           UCase$(UserList(UserIndex).clase) = "PALADIN" Then
                         If UCase$(UserList(UserIndex).raza) = "ENANO" Or UCase$(UserList(UserIndex).raza) = "GNOMO" Or UCase$(UserList(UserIndex).raza) = "GOBLIN" Then
                  MiObj.ObjIndex = ArmaduraImperial3
              Else
                  MiObj.ObjIndex = ArmaduraImperial1
              End If
    Else
              If UCase$(UserList(UserIndex).raza) = "ENANO" Or UCase$(UserList(UserIndex).raza) = "GNOMO" Or UCase$(UserList(UserIndex).raza) = "GOBLIN" Then
                  MiObj.ObjIndex = 522
              Else
                  MiObj.ObjIndex = ArmaduraImperial2
              If UCase$(UserList(UserIndex).Genero) = "MUJER" Then MiObj.ObjIndex = 719
              End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.exp, ExpAlUnirse, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 1
    Call CheckUserLevel(UserIndex)
'pluto:2.17
UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + 50
Call SendData(ToIndex, UserIndex, 0, "||Has ganado 50 SkillPoints." & "´" & FontTypeNames.FONTTYPE_info)
'--------------

End If


Call LogEjercitoReal(UserList(UserIndex).Name)
Exit Sub
fallo:
Call LogError("enlistararmadareal " & Err.number & " D: " & Err.Description)

End Sub
Public Sub Enlistarlegion(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).Faccion.ArmadaReal = 2 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya perteneces a las tropas de la Legión!!! Ve a combatir criminales!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya perteneces a las tropas de la Armada Real!!! No puedes pertenecer a la Legión.°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya has pertenecido a las tropas de la Armada Real!!! No puedes pertenecer a la Legión.°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Maldito insolente!!! vete de aqui seguidor de las sombras!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6°No se permiten criminales en la Legión.!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If


If UserList(UserIndex).Stats.ELV < 30 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Para unirte a nuestras fuerzas debes ser al menos de nivel 30!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 5 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Has asesinado más de 5 inocentes, no aceptamos asesinos en las tropas de la Legión!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Faccion.ArmadaReal = 2
'pluto:2.4.7
'UserList(userindex).Faccion.RecompensasReal = (UserList(userindex).Stats.ELV - 28) \ 2
UserList(UserIndex).Faccion.RecompensasReal = 1

Call SendData(ToIndex, UserIndex, 0, "||6°Bienvenido a al Ejercito de la Legión!!!, aqui tienes tu armadura. Por cada dos niveles que subas te dare una recompensa, buena suerte soldado!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
'pluto:2.3
If UserList(UserIndex).Faccion.RecibioArmaduraLegion = 0 Then
    Dim MiObj As obj
    MiObj.Amount = 1
    If UCase$(UserList(UserIndex).clase) = "MAGO" Or UCase$(UserList(UserIndex).clase) = "DRUIDA" Then
           If UCase$(UserList(UserIndex).raza) = "ENANO" Or UCase$(UserList(UserIndex).raza) = "GNOMO" Or UCase$(UserList(UserIndex).raza) = "GOBLIN" Then
                  MiObj.ObjIndex = TunicaMagoLegionEnanos
           Else
                  MiObj.ObjIndex = TunicaMagoLegion
           End If
    ElseIf UCase$(UserList(UserIndex).clase) = "GUERRERO" Or _
    UCase$(UserList(UserIndex).clase) = "PALADIN" Then

              If UCase$(UserList(UserIndex).raza) = "ENANO" Or UCase$(UserList(UserIndex).raza) = "GNOMO" Or UCase$(UserList(UserIndex).raza) = "GOBLIN" Then
                  MiObj.ObjIndex = ArmaduraLegion3
              Else
                  MiObj.ObjIndex = ArmaduraLegion1
              End If
    Else
              If UCase$(UserList(UserIndex).raza) = "ENANO" Or UCase$(UserList(UserIndex).raza) = "GNOMO" Or UCase$(UserList(UserIndex).raza) = "GOBLIN" Then
                  MiObj.ObjIndex = ArmaduraLegion3
              Else
                  MiObj.ObjIndex = ArmaduraLegion2
              End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraLegion = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.exp, ExpAlUnirse, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 2
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoReal(UserList(UserIndex).Name)
Exit Sub
fallo:
Call LogError("enlistarlegion " & Err.number & " D: " & Err.Description)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).Faccion.CriminalesMatados \ 15 <= _
   UserList(UserIndex).Faccion.RecompensasReal Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya has recibido tu recompensa, mata 30 criminales mas para recibir la proxima!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
Exit Sub
End If


'pluto:2.3
Dim dar As Integer
Dim Qui As Integer
Dim recibida As Byte
Dim clase As String
Dim raza As String
Dim Genero As String
Dim recompensa As Byte
'Dim alli As Byte
recibida = UserList(UserIndex).Faccion.RecibioArmaduraReal
recompensa = UserList(UserIndex).Faccion.RecompensasReal
clase = UCase$(UserList(UserIndex).clase)
raza = UCase$(UserList(UserIndex).raza)
Genero = UCase$(UserList(UserIndex).Genero)

'pluto:2.17
If recompensa > 9 Then Exit Sub

If recompensa <> 4 And recompensa <> 7 Then GoTo alli

Select Case recompensa

Case 4
If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then

If clase = "MAGO" Or clase = "DRUIDA" Then
    dar = 743
    Qui = 549
    ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
    dar = 616
    Qui = 492
Else
    dar = 955
    Qui = 522
  End If

Else 'raza
   
If clase = "MAGO" Or clase = "DRUIDA" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 618
            Qui = 517
            Case "MUJER"
            'pluto:7.0
            dar = 701
            Qui = 516
        End Select 'GENERO
  ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 620
            Qui = 370
            Case "MUJER"
            dar = 620
            Qui = 370
            
        End Select 'GENERO
   Else
    Select Case Genero
            Case "HOMBRE"
            dar = 715
            Qui = 372
            Case "MUJER"
            dar = 520
            Qui = 719
        End Select 'GENERO
   End If 'CLASE

End If 'RAZA

Case 7

If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then

If clase = "MAGO" Or clase = "DRUIDA" Then
    dar = 742
    Qui = 743
   ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
    dar = 740
    Qui = 616
Else
dar = 956
Qui = 955
End If 'CLASE


Else 'RAZA no enana
   
  If clase = "MAGO" Or clase = "DRUIDA" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 369
            Qui = 618
            Case "MUJER"
            dar = 369
            Qui = 618
        End Select 'GENERO
   ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 704
            Qui = 620
            Case "MUJER"
            dar = 704
            Qui = 620

        End Select 'GENERO
     Else
    Select Case Genero
            Case "HOMBRE"
            dar = 621
            Qui = 715
            Case "MUJER"
            dar = 521
            Qui = 520
        End Select 'GENERO
   End If 'CLASE

End If 'RAZA

End Select 'recompensa

'comprueba objeto y lo cambia
If dar = 0 Or Qui = 0 Then
Call SendData(ToIndex, UserIndex, 0, "|| No existe la ropa que te corresponde." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
Dim Slot As Integer

If UserList(UserIndex).Invent.ArmourEqpObjIndex = Qui Then
Slot = UserList(UserIndex).Invent.ArmourEqpSlot
Call QuitarUserInvItem(UserIndex, Slot, 1)
Call UpdateUserInv(False, UserIndex, Slot)

Call DarCuerpoDesnudo(UserIndex)
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Dim MiObj As obj
MiObj.Amount = 1
MiObj.ObjIndex = dar
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If

Else
Call SendData(ToIndex, UserIndex, 0, "|| No tienes la ropa del rango anterior equipada, vuelve cuando la tengas." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

alli:
    Call SendData(ToIndex, UserIndex, 0, "||6°Aqui tienes tu recompensa noble guerrero!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Call AddtoVar(UserList(UserIndex).Stats.exp, ExpX100, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    'pluto:2.17
UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + 10
Call SendData(ToIndex, UserIndex, 0, "||Has ganado 10 SkillPoints." & "´" & FontTypeNames.FONTTYPE_info)
'--------------

    Call CheckUserLevel(UserIndex)

Exit Sub
fallo:
Call LogError("recompensa armada real " & Err.number & " D: " & Err.Description)

End Sub
Public Sub Recompensalegion(ByVal UserIndex As Integer)
On Error GoTo fallo
If (UserList(UserIndex).Stats.ELV - 28) \ 2 = _
   UserList(UserIndex).Faccion.RecompensasReal Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya has recibido tu recompensa,sube más nivel para subir de rango.!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
'pluto:2.4.7 --> faltaba un exit sub
Exit Sub

End If
    
  'pluto:2.3
Dim dar As Integer
Dim Qui As Integer
Dim recibida As Byte
Dim clase As String
Dim raza As String
Dim Genero As String
Dim recompensa As Byte
'Dim alli As Byte
recibida = UserList(UserIndex).Faccion.RecibioArmaduraLegion
recompensa = UserList(UserIndex).Faccion.RecompensasReal
clase = UCase$(UserList(UserIndex).clase)
raza = UCase$(UserList(UserIndex).raza)
Genero = UCase$(UserList(UserIndex).Genero)

'pluto:2.4.7 -->arreglado fallo legion
If recompensa <> 2 And recompensa <> 5 Then GoTo alli

Select Case recompensa

Case 2
If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then

If clase = "MAGO" Or clase = "DRUIDA" Then
    dar = 885
    Qui = 810
    ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
    dar = 869
    Qui = 809
Else
    dar = 869
    Qui = 809
  End If

Else 'raza
   
If clase = "MAGO" Or clase = "DRUIDA" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 706
            Qui = 707
            Case "MUJER"
            dar = 706
            Qui = 707
        End Select 'GENERO
  ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 702
            Qui = 701
            Case "MUJER"
            dar = 702
            Qui = 701
        End Select 'GENERO
   Else
    Select Case Genero
            Case "HOMBRE"
            dar = 702
            Qui = 701
            Case "MUJER"
            dar = 702
            Qui = 701
        End Select 'GENERO
   End If 'CLASE

End If 'RAZA

Case 5

If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then

If clase = "MAGO" Or clase = "DRUIDA" Then
    dar = 886
    Qui = 885
   ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
    dar = 870
    Qui = 869
Else
    dar = 870
    Qui = 869
End If 'CLASE


Else 'RAZA no enana
   
  If clase = "MAGO" Or clase = "DRUIDA" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 708
            Qui = 706
            Case "MUJER"
            dar = 708
            Qui = 706
        End Select 'GENERO
   ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 703
            Qui = 702
            Case "MUJER"
            dar = 703
            Qui = 702
        End Select 'GENERO
     Else
    Select Case Genero
            Case "HOMBRE"
            dar = 703
            Qui = 702
            Case "MUJER"
            dar = 703
            Qui = 702
        End Select 'GENERO
   End If 'CLASE

End If 'RAZA

End Select 'recompensa

'comprueba objeto y lo cambia
If dar = 0 Or Qui = 0 Then
Call SendData(ToIndex, UserIndex, 0, "|| No existe la ropa que te corresponde." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

If UserList(UserIndex).Invent.ArmourEqpObjIndex = Qui Then
Call QuitarObjetos(Qui, 1, UserIndex)
Call DarCuerpoDesnudo(UserIndex)
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Dim MiObj As obj
MiObj.Amount = 1
MiObj.ObjIndex = dar
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If

Else
Call SendData(ToIndex, UserIndex, 0, "|| No tienes la ropa del rango anterior equipada, vuelve cuando la tengas." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
    
'pluto:2.4.7 --> poner alli:
alli:
    Call SendData(ToIndex, UserIndex, 0, "||6°Has subido de rango en las tropas de la Legión!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    'Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpX100, MAXEXP)
   ' Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPENAMES.FONTTYPE_fight)
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    'Call CheckUserLevel(UserIndex)

Exit Sub
fallo:
Call LogError("recompensalegion " & Err.number & " D: " & Err.Description)

End Sub
Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).Faccion.ArmadaReal = 0
UserList(UserIndex).Faccion.CriminalesMatados = 0
Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado de las tropas reales.!!!." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Exit Sub
fallo:
Call LogError("expulsarfaccionreal " & Err.number & " D: " & Err.Description)

End Sub
Public Sub ExpulsarFaccionlegion(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).Faccion.ArmadaReal = 0
Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado de las tropas de la Legión.!!!." & "´" & FontTypeNames.FONTTYPE_FIGHT)
Exit Sub
fallo:
Call LogError("expulsarfaccionlegion " & Err.number & " D: " & Err.Description)

End Sub
Public Function Titulolegion(ByVal UserIndex As Integer) As String
On Error GoTo fallo
Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 0
        Titulolegion = "Recluta"
    Case 1
        Titulolegion = "Soldado"
    Case 2
        Titulolegion = "Sargento"
    Case 3
        Titulolegion = "Brigada"
    Case 4
        Titulolegion = "Alferez"
    Case 5
        Titulolegion = "Teniente"
    Case 6
        Titulolegion = "Capitán"
    Case 7
        Titulolegion = "Comandante"
    Case 8
        Titulolegion = "General"
    Case Else
        Titulolegion = "Almirante"
End Select
Exit Function
fallo:
Call LogError("titulolegion " & Err.number & " D: " & Err.Description)

End Function
Public Function TituloReal(ByVal UserIndex As Integer) As String
On Error GoTo fallo
Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 1
        TituloReal = "Guerrero Imperial"
    Case 2
        TituloReal = "Teniente Imperial"
    Case 3
        TituloReal = "Capitán Imperial"
    Case 4
        TituloReal = "Comandante Imperial"
    Case 5
        TituloReal = "General Imperial"
    Case 6
        TituloReal = "Elite Imperial"
    Case 7
        TituloReal = "Protector del Imperio"
    Case 8
        TituloReal = "Caballero de la Luz"
    Case 9
        TituloReal = "Escolta del Imperio"
    Case Else
        TituloReal = "Garante del Orden Imperial"
End Select
Exit Function
fallo:
Call LogError("tituloreal " & Err.number & " D: " & Err.Description)

End Function
Public Sub EnlistarCaos(ByVal UserIndex As Integer)
On Error GoTo fallo
If Not Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Largate de aqui, bufon!!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya perteneces a las tropas del caos!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Las sombras reinaran en Argentum, largate de aqui estupido ciudadano.!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
If UserList(UserIndex).Faccion.RecibioExpInicialReal > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°No queremos antiguos miembros del Bién en nuestras filas.°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
Exit Sub
End If
'pluto:hoy
If UserList(UserIndex).Faccion.RecibioExpInicialCaos > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°No queremos antiguos miembros del Caos en nuestras filas.°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
Exit Sub
End If

If Not Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ja ja ja tu no eres bienvenido aqui!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CiudadanosMatados < 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Para unirte a nuestras fuerzas debes matar al menos 1 ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 30 Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Para unirte a nuestras fuerzas debes ser al menos de nivel 30!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Faccion.FuerzasCaos = 1
'pluto:2.4.7 --> enlistar con muertes justas
'UserList(userindex).Faccion.RecompensasCaos = UserList(userindex).Faccion.CiudadanosMatados \ 100
UserList(UserIndex).Faccion.RecompensasCaos = 1

Call SendData(ToIndex, UserIndex, 0, "||6°Bienvenido a al lado oscuro!!!, aqui tienes tu armadura. Por cada centena de ciudadanos que acabes te dare un recompensa, buena suerte soldado!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As obj
    MiObj.Amount = 1
    If UCase$(UserList(UserIndex).clase) = "MAGO" Or UCase$(UserList(UserIndex).clase) = "DRUIDA" Then
                  MiObj.ObjIndex = TunicaMagoCaos
    'pluto:2.4.7 --> Poner genero
    If UCase$(UserList(UserIndex).Genero) = "MUJER" Then MiObj.ObjIndex = 509

    'pluto:7.0 GOBLIN
    If UCase$(UserList(UserIndex).raza) = "ENANO" Or UCase$(UserList(UserIndex).raza) = "GNOMO" Or UCase$(UserList(UserIndex).raza) = "GOBLIN" Then
                  MiObj.ObjIndex = 524 'TunicaMagoCaosEnanos
    End If
    
    
    ElseIf UCase$(UserList(UserIndex).clase) = "GUERRERO" Or _
        UCase$(UserList(UserIndex).clase) = "PALADIN" Then
                         If UCase$(UserList(UserIndex).raza) = "ENANO" Or UCase$(UserList(UserIndex).raza) = "GNOMO" Or UCase$(UserList(UserIndex).raza) = "GOBLIN" Then
                  MiObj.ObjIndex = ArmaduraCaos3
              Else
                  MiObj.ObjIndex = ArmaduraCaos1
              End If
    Else
              If UCase$(UserList(UserIndex).raza) = "ENANO" Or UCase$(UserList(UserIndex).raza) = "GNOMO" Or UCase$(UserList(UserIndex).raza) = "GOBLIN" Then
                  MiObj.ObjIndex = 957
              Else
                  MiObj.ObjIndex = ArmaduraCaos2
              End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.exp, ExpAlUnirse, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoCaos(UserList(UserIndex).Name)
Exit Sub
fallo:
Call LogError("enlistarcaos " & Err.number & " D: " & Err.Description)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
On Error GoTo fallo

If UserList(UserIndex).Faccion.CiudadanosMatados \ 15 <= _
   UserList(UserIndex).Faccion.RecompensasCaos Then
    Call SendData(ToIndex, UserIndex, 0, "||6°Ya has recibido tu recompensa, mata 30 ciudadanos mas para recibir la proxima!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
Exit Sub
End If


'pluto:2.3
Dim dar As Integer
Dim Qui As Integer
Dim recibida As Byte
Dim clase As String
Dim raza As String
Dim Genero As String
Dim recompensa As Byte
'Dim alli As Byte
recibida = UserList(UserIndex).Faccion.RecibioArmaduraCaos
recompensa = UserList(UserIndex).Faccion.RecompensasCaos
clase = UCase$(UserList(UserIndex).clase)
raza = UCase$(UserList(UserIndex).raza)
Genero = UCase$(UserList(UserIndex).Genero)


If recompensa <> 4 And recompensa <> 7 Then GoTo alli

Select Case recompensa

Case 4
If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then

If clase = "MAGO" Or clase = "DRUIDA" Then
    dar = 562
    Qui = 524
    ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
    dar = 615
    Qui = 593
    Else
    dar = 958
    Qui = 957
  End If

Else 'raza
   
If clase = "MAGO" Or clase = "DRUIDA" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 613
            Qui = 518
            Case "MUJER"
            dar = 613
            Qui = 509
        End Select 'GENERO
  ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 808
            Qui = 379
            Case "MUJER"
            dar = 494
            Qui = 379
            
       End Select 'GENERO
   Else
    Select Case Genero
            Case "HOMBRE"
            dar = 614
            Qui = 523
            Case "MUJER"
            dar = 614
            Qui = 523
        End Select 'GENERO
   End If 'CLASE

End If 'RAZA

Case 7

If raza = "ENANO" Or raza = "GNOMO" Or raza = "GOBLIN" Then

If clase = "MAGO" Or clase = "DRUIDA" Then
    dar = 739
    Qui = 562
   ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
    dar = 953
    Qui = 615
Else
    dar = 959
    Qui = 958
End If 'CLASE


Else 'RAZA no enana
   
  If clase = "MAGO" Or clase = "DRUIDA" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 714
            Qui = 613
            Case "MUJER"
            dar = 380
            Qui = 613
        End Select 'GENERO
   ElseIf clase = "GUERRERO" Or clase = "PALADIN" Then
        Select Case Genero
            Case "HOMBRE"
            dar = 617
            Qui = 808
            Case "MUJER"
            dar = 617
            Qui = 494
            
        End Select 'GENERO
     Else
    Select Case Genero
            Case "HOMBRE"
            dar = 954
            Qui = 614
            Case "MUJER"
            dar = 954
            Qui = 614
        End Select 'GENERO
   End If 'CLASE

End If 'RAZA

End Select 'recompensa

'comprueba objeto y lo cambia
If dar = 0 Or Qui = 0 Then
Call SendData(ToIndex, UserIndex, 0, "|| No existe la ropa que te corresponde." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
Dim Slot As Integer

If UserList(UserIndex).Invent.ArmourEqpObjIndex = Qui Then
Slot = UserList(UserIndex).Invent.ArmourEqpSlot
Call QuitarUserInvItem(UserIndex, Slot, 1)
Call UpdateUserInv(False, UserIndex, Slot)


'Call QuitarObjetos(Qui, 1, UserIndex)
Call DarCuerpoDesnudo(UserIndex)
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Dim MiObj As obj
MiObj.Amount = 1
MiObj.ObjIndex = dar
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If

Else
Call SendData(ToIndex, UserIndex, 0, "|| No tienes la ropa del rango anterior equipada, vuelve cuando la tengas." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

alli:
    Call SendData(ToIndex, UserIndex, 0, "||6°Aqui tienes tu recompensa noble guerrero!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Call AddtoVar(UserList(UserIndex).Stats.exp, ExpX100, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & "´" & FontTypeNames.FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.RecompensasCaos + 1
    Call CheckUserLevel(UserIndex)
    
        'pluto:2.17
UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + 10
Call SendData(ToIndex, UserIndex, 0, "||Has ganado 10 SkillPoints." & "´" & FontTypeNames.FONTTYPE_info)

Exit Sub
fallo:
Call LogError("recompensacaos " & Err.number & " D: " & Err.Description)


End Sub

Public Sub ExpulsarCaos(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).Faccion.FuerzasCaos = 0
Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado del ejercito del caos!!!." & "´" & FontTypeNames.FONTTYPE_FIGHT)

Exit Sub
fallo:
Call LogError("expulsarcaos " & Err.number & " D: " & Err.Description)

End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
On Error GoTo fallo
Select Case UserList(UserIndex).Faccion.RecompensasCaos
    
    Case 1
        TituloCaos = "Guerrero del caos"
    Case 2
        TituloCaos = "Teniente del caos"
    Case 3
        TituloCaos = "Capitán del caos"
    Case 4
        TituloCaos = "Comandante del caos"
    Case 5
        TituloCaos = "General del caos"
    Case 6
        TituloCaos = "Elite caos"
    Case 7
        TituloCaos = "Asolador de las sombras"
    Case 8
        TituloCaos = "Caballero Oscuro"
    Case 9
        TituloCaos = "Asesino del caos"
    Case Else
        TituloCaos = "Adorador del demonio"
End Select

Exit Function
fallo:
Call LogError("titulocaos " & Err.number & " D: " & Err.Description)

End Function

