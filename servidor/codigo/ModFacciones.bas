Attribute VB_Name = "ModFacciones"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Public Const ExpAlUnirse As Long = 50000
Public Const ExpX100 As Integer = 5000


Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 15/03/2009
'15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
'Handles the entrance of users to the "Armada Real"
'***************************************************
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Maldito insolente!!! vete de aqui seguidor de las sombras", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If criminal(UserIndex) Then
    Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten criminales en el ejército imperial!!!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 5 Then
    Call WriteChatOverHead(UserIndex, "¡Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.Reenlistadas > 1 Then
    Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas reales demasiadas veces!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Dim Slot As Integer
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
    Call WriteChatOverHead(UserIndex, "Tu inventario esta lleno, vuelve cuando tengas un hueco para recibir la armadura.", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
       End If
   Loop
End If

With UserList(UserIndex)
    If .GuildIndex > 0 Then
        If modGuilds.GuildFounder(.GuildIndex) = .Name Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutro" Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Eres el fundador de un clan neutro!!!", str(NPCList(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
    End If
End With

UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1

Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al Ejército Imperial!!!, aqui tienes tu uniforme. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
Call WriteConsoleMsg(UserIndex, "¡Recuerda que si desertas no podras volver a formar parte de la Armada!", FontTypeNames.FONTTYPE_INFOBOLD)

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    
    With UserList(UserIndex)
        If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
            MiObj.ObjIndex = 676
        Else
            If .genero = Hombre Then
                MiObj.ObjIndex = 675
            Else
                MiObj.ObjIndex = 679
            End If
        End If
    End With
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
    UserList(UserIndex).Faccion.NextRecompensa = 30
    UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).Faccion.FechaIngreso = Date
    'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
    UserList(UserIndex).Faccion.MatadosIngreso = UserList(UserIndex).Faccion.CiudadanosMatados

End If

'Agregado para que no hayan armadas en un clan Neutro
If UserList(UserIndex).GuildIndex > 0 Then
    If modGuilds.GuildAlignment(UserList(UserIndex).GuildIndex) = "Neutro" Then
        Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
        Call WriteConsoleMsg(UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

Call LogEjercitoReal(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the way of gaining new ranks in the "Armada Real"
'***************************************************
Dim Crimis As Long
Dim Lvl As Byte
Dim NextRecom As Long
Dim Nobleza As Long
Lvl = UserList(UserIndex).Stats.ELV
NextRecom = UserList(UserIndex).Faccion.NextRecompensa
Nobleza = UserList(UserIndex).Reputacion.NobleRep

If Lvl < NextRecom Then
    Call WriteChatOverHead(UserIndex, "Vuelve cuando seas nivel " & NextRecom & " para recibir la próxima Recompensa", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Select Case NextRecom
    Case 30:
        UserList(UserIndex).Faccion.RecompensasReal = 1
        UserList(UserIndex).Faccion.NextRecompensa = 32
    
    Case 32:
        UserList(UserIndex).Faccion.RecompensasReal = 2
        UserList(UserIndex).Faccion.NextRecompensa = 34
    
    Case 34:
        UserList(UserIndex).Faccion.RecompensasReal = 3
        UserList(UserIndex).Faccion.NextRecompensa = 36
    
    Case 36:
        UserList(UserIndex).Faccion.RecompensasReal = 4
        UserList(UserIndex).Faccion.NextRecompensa = 38
    
    Case 38:
        UserList(UserIndex).Faccion.RecompensasReal = 5
        UserList(UserIndex).Faccion.NextRecompensa = 40
    
    Case 40:
        UserList(UserIndex).Faccion.RecompensasReal = 6
        UserList(UserIndex).Faccion.NextRecompensa = 45
    
    Case 45:
        UserList(UserIndex).Faccion.RecompensasReal = 7
        UserList(UserIndex).Faccion.NextRecompensa = 50
    
    Case 50:
        UserList(UserIndex).Faccion.RecompensasReal = 8
        UserList(UserIndex).Faccion.NextRecompensa = 1000
    
    Case Else:
        Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores Soldados. Ya no tengo más recompensa para darte que mi agradescimiento. ¡Felicidades!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
End Select

Call WriteChatOverHead(UserIndex, "¡¡¡Aqui tienes tu recompensa " + TituloReal(UserIndex) + "!!!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
If UserList(UserIndex).Stats.Exp > MAXEXP Then
    UserList(UserIndex).Stats.Exp = MAXEXP
End If
Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpX100 & " puntos de experiencia.", FontTypeNames.FONTTYPE_exp)

Call CheckUserLevel(UserIndex)

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)

    UserList(UserIndex).Faccion.ArmadaReal = 0
    'Call PerderItemsFaccionarios(UserIndex)
    If Expulsado Then
        Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado de las tropas reales!!!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado de las tropas reales!!!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        'Desequipamos el escudo de caos si está equipado
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)

    UserList(UserIndex).Faccion.FuerzasCaos = 0
    'Call PerderItemsFaccionarios(UserIndex)
    If Expulsado Then
        Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado de la Legión Oscura!!!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado de la Legión Oscura!!!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura de caos si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        'Desequipamos el escudo de caos si está equipado
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'Handles the titles of the members of the "Armada Real"
'***************************************************
Select Case UserList(UserIndex).Faccion.RecompensasReal
'Rango 1: Aprendiz (30 Criminales)
'Rango 2: Escudero (70 Criminales)
'Rango 3: Soldado (130 Criminales)
'Rango 4: Sargento (210 Criminales)
'Rango 5: Caballero (320 Criminales)
'Rango 6: Comandante (460 Criminales)
'Rango 7: Capitán (640 Criminales + > lvl 27)
'Rango 8: Senescal (870 Criminales)
'Rango 9: Mariscal (1160 Criminales)
'Rango 10: Condestable (2000 Criminales + > lvl 30)
'Rangos de Honor de la Armada Real: (Consejo de Bander)
'Rango 11: Ejecutor Imperial (2500 Criminales + 2.000.000 Nobleza)
'Rango 12: Protector del Reino (3000 Criminales + 3.000.000 Nobleza)
'Rango 13: Avatar de la Justicia (3500 Criminales + 4.000.000 Nobleza + > lvl 35)
'Rango 14: Guardián del Bien (4000 Criminales + 5.000.000 Nobleza + > lvl 36)
'Rango 15: Campeón de la Luz (5000 Criminales + 6.000.000 Nobleza + > lvl 37)
    
    Case 0
        TituloReal = "Soldado"
    Case 1
        TituloReal = "Teniente"
    Case 2
        TituloReal = "Capitán"
    Case 3
        TituloReal = "Mariscal"
    Case 4
        TituloReal = "Ejecutor del Imperio"
    Case 5
        TituloReal = "Protector del Reino"
    Case 6
        TituloReal = "Avatar de la Justicia"
    Case 7
        TituloReal = "Guardián del Bien"
    Case 8
        TituloReal = "Protector de Newbies"
End Select


End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 15/3/2009
'15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
'Handles the entrance of users to the "Legión Oscura"
'***************************************************
If Not criminal(UserIndex) Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Lárgate de aqui, bufón!!!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a la legión oscura!!!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call WriteChatOverHead(UserIndex, "Las sombras reinarán en las tierras del Dragón. ¡¡¡Fuera de aqui insecto Real!!!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

'[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
    Call WriteChatOverHead(UserIndex, "No permitiré que ningún insecto real ingrese a mis tropas.", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If
'[/Barrin]

If Not criminal(UserIndex) Then
    Call WriteChatOverHead(UserIndex, "¡¡Ja ja ja!! Tu no eres bienvenido aqui asqueroso Ciudadano", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Dim Slot As Integer
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
    Call WriteChatOverHead(UserIndex, "Tu inventario esta lleno, vuelve cuando tengas un hueco para recibir la armadura.", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
       End If
   Loop
End If

With UserList(UserIndex)
    If .GuildIndex > 0 Then
        If modGuilds.GuildFounder(.GuildIndex) = .Name Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutro" Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Eres el fundador de un clan neutro!!!", str(NPCList(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
    End If
End With


If UserList(UserIndex).Faccion.Reenlistadas > 1 Then
    If UserList(UserIndex).Faccion.Reenlistadas = 200 Then
        Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Else
        Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas oscuras demasiadas veces!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    End If
    Exit Sub
End If

UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
UserList(UserIndex).Faccion.FuerzasCaos = 1

Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al lado oscuro!!! Aqui tienes tu uniforme. Derrama sangre Ciudadana y Real y serás recompensado, lo prometo.", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
Call WriteConsoleMsg(UserIndex, "¡Recuerda que si desertas del Caos no podras volver!", FontTypeNames.FONTTYPE_INFOBOLD)
If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    
    With UserList(UserIndex)
        If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
            MiObj.ObjIndex = 678 'Vestimenta Fuerzas del Caos (E/G)
        Else
            If .genero = Hombre Then
                MiObj.ObjIndex = 677 'Vestimenta Fuerzas del Caos
             Else
                MiObj.ObjIndex = 680 'Vestimenta Fuerzas del Caos (M)
            End If
        End If
    End With
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If

    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
    UserList(UserIndex).Faccion.NextRecompensa = 25
    UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).Faccion.FechaIngreso = Date

End If

'Agregado para que no hayan armadas en un clan Neutro
If UserList(UserIndex).GuildIndex > 0 Then
    If modGuilds.GuildAlignment(UserList(UserIndex).GuildIndex) = "Neutro" Then
        Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
        Call WriteConsoleMsg(UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

Call LogEjercitoCaos(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)

'***************************************************
'Author: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the way of gaining new ranks in the "Legión Oscura"
'***************************************************
Dim Ciudas As Long
Dim Lvl As Byte
Dim NextRecom As Long
Lvl = UserList(UserIndex).Stats.ELV
NextRecom = UserList(UserIndex).Faccion.NextRecompensa

If Lvl < NextRecom Then
    Call WriteChatOverHead(UserIndex, "Necesitas ser nivel " & NextRecom & " para recibir la próxima Recompensa", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Select Case NextRecom
    Case 25:
        UserList(UserIndex).Faccion.RecompensasCaos = 1
        UserList(UserIndex).Faccion.NextRecompensa = 30
    
    Case 30:
        UserList(UserIndex).Faccion.RecompensasCaos = 2
        UserList(UserIndex).Faccion.NextRecompensa = 32
    
    Case 32:
        UserList(UserIndex).Faccion.RecompensasCaos = 3
        UserList(UserIndex).Faccion.NextRecompensa = 34
    
    Case 34:
        UserList(UserIndex).Faccion.RecompensasCaos = 4
        UserList(UserIndex).Faccion.NextRecompensa = 36
    
    Case 36:
        UserList(UserIndex).Faccion.RecompensasCaos = 5
        UserList(UserIndex).Faccion.NextRecompensa = 38
        
    Case 38:
        UserList(UserIndex).Faccion.RecompensasCaos = 7
        UserList(UserIndex).Faccion.NextRecompensa = 40
        
    Case 40:
        UserList(UserIndex).Faccion.RecompensasCaos = 8
        UserList(UserIndex).Faccion.NextRecompensa = 45
        
    Case 45:
        UserList(UserIndex).Faccion.RecompensasCaos = 9
        UserList(UserIndex).Faccion.NextRecompensa = 50
    
    Case 50:
        UserList(UserIndex).Faccion.RecompensasCaos = 10
        UserList(UserIndex).Faccion.NextRecompensa = 100

    Case Else:
        Exit Sub
        
End Select

Call WriteChatOverHead(UserIndex, "¡¡¡Bien hecho " + TituloCaos(UserIndex) + ", aquí tienes tu recompensa!!!", str(NPCList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
If UserList(UserIndex).Stats.Exp > MAXEXP Then
    UserList(UserIndex).Stats.Exp = MAXEXP
End If
Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpX100 & " puntos de experiencia.", FontTypeNames.FONTTYPE_exp)
Call CheckUserLevel(UserIndex)


End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'Handles the titles of the members of the "Legión Oscura"
'***************************************************
'Rango 1: Acólito (70)
'Rango 2: Alma Corrupta (160)
'Rango 3: Paria (300)
'Rango 4: Condenado (490)
'Rango 5: Esbirro (740)
'Rango 6: Sanguinario (1100)
'Rango 7: Corruptor (1500 + lvl 27)
'Rango 8: Heraldo Impio (2010)
'Rango 9: Caballero de la Oscuridad (2700)
'Rango 10: Señor del Miedo (4600 + lvl 30)
'Rango 11: Ejecutor Infernal (5800 + lvl 31)
'Rango 12: Protector del Averno (6990 + lvl 33)
'Rango 13: Avatar de la Destrucción (8100 + lvl 35)
'Rango 14: Guardián del Mal (9300 + lvl 36)
'Rango 15: Campeón de la Oscuridad (11500 + lvl 37)

Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Acólito"
    Case 1
        TituloCaos = "Alma Corrupta"
    Case 2
        TituloCaos = "Esbirro"
    Case 3
        TituloCaos = "Sanguinario"
    Case 4
        TituloCaos = "Caballero de la Oscuridad"
    Case 5
        TituloCaos = "Señor del Miedo"
    Case 6
        TituloCaos = "Avatar de la Destrucción"
    Case 7
        TituloCaos = "Ejecutor Infernal"
    Case Else
        TituloCaos = "Campeón de la Oscuridad"
End Select

End Function

