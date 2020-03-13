Attribute VB_Name = "SistemaCombate"
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
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat


Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a
    End If
End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b
    End If
End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
    PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModClase(UserList(UserIndex).clase).Evasion) / 2
End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
    Dim lTemp As Long
    With UserList(UserIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * AgilidadMaxima(UserIndex)) * ModClase(.clase).Evasion
          
        '<Edurne> 'En montura...
        If UserList(UserIndex).flags.QueMontura Then lTemp = lTemp + UserList(UserIndex).flags.Montura(UserList(UserIndex).flags.QueMontura).Evasion
        
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
    
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Armas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Armas) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 2 * AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueArmas
        Else
           PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 3 * AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueArmas
        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Proyectiles) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Proyectiles) * ModClase(.clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 2 * AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 3 * AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueProyectiles
        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Wrestling) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Wrestling) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 2 * AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 3 * AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueArmas
        End If
        
        PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderTrabajo(ByVal UserIndex As Integer, ByVal Skill As Byte) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(Skill) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(Skill) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(Skill) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(Skill) + AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(Skill) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(Skill) + 2 * AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(Skill) + 3 * AgilidadMaxima(UserIndex)) * ModClase(.clase).AtaqueArmas
        End If
        
        PoderTrabajo = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim Skill As eSkill
    Dim ProbExito As Long
    
    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
    
    If UserTrabajando = True Then 'Si el usuario lo hacemos con el skill del trabajo
        Skill = SkillTrabajando
        PoderAtaque = PoderTrabajo(UserIndex, SkillTrabajando)
        
    Else
        If Arma > 0 Then 'Usando un arma
            If ObjData(Arma).proyectil = 1 Then
                PoderAtaque = PoderAtaqueProyectil(UserIndex)
                Skill = eSkill.Proyectiles
                Dim MunicionObjIndex    As Integer
                MunicionObjIndex = UserList(UserIndex).Invent.MunicionEqpObjIndex
                'Tiene munición?
                If MunicionObjIndex <> 0 Then
                    Call WriteProyectil(UserIndex, UserList(UserIndex).Char.CharIndex, NPCList(NpcIndex).Char.CharIndex, ObjData(MunicionObjIndex).GrhIndex)
                End If
            Else
                PoderAtaque = PoderAtaqueArma(UserIndex)
                Skill = eSkill.Armas
            End If
        Else 'Peleando con puños
            PoderAtaque = PoderAtaqueWrestling(UserIndex)
            Skill = eSkill.Wrestling
        End If
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - NPCList(NpcIndex).PoderEvasion) * 0.4)))
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    If UserImpactoNpc Then
            Call SubirSkill(UserIndex, Skill)
    End If
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'*************************************************
    Dim Rechazo As Boolean
    Dim ProbRechazo As Long
    Dim ProbExito As Long
    Dim UserEvasion As Long
    Dim NpcPoderAtaque As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = NPCList(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
    
    SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)
    
    'Esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    Call WriteBlockedWithShieldUser(UserIndex)
                    Call SubirSkill(UserIndex, Defensa)
                End If
            End If
        End If
    End If
End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
    Dim DañoArma As Long
    Dim DañoUsuario As Long
    Dim Arma As ObjData
    Dim ModifClase As Single
    Dim proyectil As ObjData
    Dim DañoMaxArma As Long
    
    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean
    matoDragon = False
    
    With UserList(UserIndex)
        If .Invent.WeaponEqpObjIndex > 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If NpcIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.clase).DañoProyectiles
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                    
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = ModClase(.clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mata Dragones?
                        If NPCList(NpcIndex).NPCType = DRAGON Then 'Ataca Dragon?
                            DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                            DañoMaxArma = Arma.MaxHIT
                            matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                        Else ' Sino es Dragon daño es 1
                            DañoArma = 1
                            DañoMaxArma = 1
                        End If
                    Else
                        DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        DañoMaxArma = Arma.MaxHIT
                    End If
                End If
            Else ' Ataca usuario
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.clase).DañoProyectiles
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                     
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = ModClase(.clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        ModifClase = ModClase(.clase).DañoArmas
                        DañoArma = 1 ' Si usa la espada mataDragones daño es 1
                        DañoMaxArma = 1
                    Else
                        DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        DañoMaxArma = Arma.MaxHIT
                    End If
                End If
            End If
        Else
            ModifClase = ModClase(.clase).DañoWrestling
            DañoArma = RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
            DañoMaxArma = 3
        End If
        
        DañoUsuario = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        
        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        If matoDragon Then
            CalcularDaño = NPCList(NpcIndex).Stats.MinHP + NPCList(NpcIndex).Stats.def
        Else
            CalcularDaño = (3 * DañoArma + ((DañoMaxArma / 5) * MaximoInt(0, FuerzaMaxima(UserIndex) - 15)) + DañoUsuario) * ModifClase
        End If
    End With
End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    Dim daño As Long
    Dim i As Byte
    
    daño = CalcularDaño(UserIndex, NpcIndex)
        
    'esta navegando? si es asi le sumamos el daño del barco
    If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
        daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)
    End If
    
    With NPCList(NpcIndex)
        
        '<Edurne> 'En montura...
        If UserList(UserIndex).flags.QueMontura Then _
            daño = daño + UserList(UserIndex).flags.Montura(UserList(UserIndex).flags.QueMontura).Ataque
        
        If GranPoder = UserIndex Then daño = daño * MultiplicadorGPN
        daño = daño - .Stats.def
        
        If daño < 0 Then daño = 0
               
        'Irongete: Si el jugador está invisible hace la mitad de daño
        'Lorwik: Si estas oculto tambien.
        If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
            daño = daño * 0.5
        End If
        
        
        If NPCList(NpcIndex).Numero = NPCReyCastle Then
            If UserList(UserIndex).GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, MSG_ATK_CASTILLO_NOCLAN, FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                Call ReyEsAtacado(UserIndex, NpcIndex, daño)
            End If
        End If
        
        '16/11/2015 Irongete: Está atacando la puerta del castillo
        If NPCList(NpcIndex).NPCType = 10 Then
            If UserList(UserIndex).GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, MSG_ATK_CASTILLO_NOCLAN, FontTypeNames.FONTTYPE_INFO)
            Else
                Call PuertaEsAtacada(UserIndex, NpcIndex, daño)
            End If
            
        End If
        
        '26/02/2016 Irongete: Está atacando el defensor de la fortaleza
        If NPCList(NpcIndex).Numero = NPCDefensorFortaleza Then
            If UserList(UserIndex).GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, MSG_ATK_CASTILLO_NOCLAN, FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                Call DefensorEsAtacado(UserIndex, NpcIndex, daño)
            End If
        End If
    
        
                
        Call EventosDaño(UserIndex, NpcIndex, daño)
        
        .Stats.MinHP = .Stats.MinHP - daño
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateRenderValue(.Pos.X, .Pos.Y, daño, DAMAGE_NORMAL))
        
        If .Stats.MinHP > 0 Then
            Dim LeQueda As String
            LeQueda = NPCList(NpcIndex).Name & " (" & NPCList(NpcIndex).Stats.MinHP & "/" & NPCList(NpcIndex).Stats.MaxHP & ")"
            Call WriteUserHitNPC(UserIndex, daño, LeQueda)
        End If
        
        If .Stats.MinHP > 0 Then
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(UserIndex) Then
               Call DoApuñalar(UserIndex, NpcIndex, 0, daño)
               Call SubirSkill(UserIndex, Apuñalar)
            End If
            
            'trata de dar golpe crítico
            Call DoGolpeCritico(UserIndex, NpcIndex, 0, daño)
        End If
        
        If .Stats.MinHP <= 0 Then
            ' Si era un Dragon perdemos la espada mataDragones
            If .NPCType = DRAGON Then
                'Si tiene equipada la matadracos se la sacamos
                If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                    Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
                End If
                If .Stats.MaxHP > 100000 Then Call LogDesarrollo(UserList(UserIndex).Name & " mató un dragón")
            End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If UserList(UserIndex).MascotasIndex(j) > 0 Then
                    If NPCList(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then
                        NPCList(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
                        NPCList(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
                    End If
                End If
            Next j
            
            Call MuereNpc(NpcIndex, UserIndex)
        End If
        
        'Llegados a este pundo ya debio restarle vida o el NPC murio, asi que lo ponemos en false
        UserTrabajando = False
        SkillTrabajando = 0
    End With
End Sub

Public Sub EventosDaño(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal daño As Long)
'***************************************************************************
'Autor: Lorwik
'Descripción: Comprueba si se ejecuta algun evento al provocar daño al NPC
'***************************************************************************
    With NPCList(NpcIndex)
        
        If (.Stats.MinHP / .Stats.MaxHP) * 100 <= 10 Then
            If .flags.LanzaMensaje = 1 Then
                If .flags.DijoMensaje = False Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(.flags.Mensaje, .Char.CharIndex, vbBlue))
                    .flags.DijoMensaje = True
                End If
            End If
        End If
        
        If (.Stats.MinHP / .Stats.MaxHP) * 100 <= 50 Then
            If .flags.AumentaPotencia = True Then
                If .flags.ActivoPotencia = False Then
                    .Stats.MaxHIT = .Stats.MaxHIT * 2
                    .Stats.MinHIT = .Stats.MinHIT * 2
                    .flags.ActivoPotencia = True
                End If
            End If
        End If
        
    End With
End Sub

Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    Dim daño As Integer
    Dim Lugar As Integer
    Dim absorbido As Integer
    Dim defbarco As Integer
    Dim Obj As ObjData
    
    daño = RandomNumber(NPCList(NpcIndex).Stats.MinHIT, NPCList(NpcIndex).Stats.MaxHIT)
    
    With UserList(UserIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
            Obj = ObjData(.Invent.BarcoObjIndex)
            defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If .Invent.CascoEqpObjIndex > 0 Then
                   Obj = ObjData(.Invent.CascoEqpObjIndex)
                   absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If
          Case Else
                'Si tiene armadura absorbe el golpe
                If .Invent.ArmourEqpObjIndex > 0 Then
                    Dim Obj2 As ObjData
                    Obj = ObjData(.Invent.ArmourEqpObjIndex)
                    If .Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(.Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                   End If
                End If
        End Select
        
        absorbido = absorbido + defbarco
        daño = daño - absorbido
        If daño < 1 Then daño = 1
        
        Call WriteNPCHitUser(UserIndex, Lugar, daño)
        
        If .flags.Privilegios And PlayerType.User Then
            .Stats.MinHP = .Stats.MinHP - daño
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, daño, DAMAGE_NORMAL))
        End If
        
        If .flags.Meditando Then
            If daño > Fix(.Stats.MinHP / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                .Char.fx = 0
                .Char.loops = 0
                
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            End If

            If .flags.Navegando = 0 Then
                If daño < 100 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
                Else
                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGREXXL, 0))
                End If
            End If
        End If
        
        'Muere el usuario
        If .Stats.MinHP <= 0 Then
            Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
            
            'Si lo mato un guardia
            If criminal(UserIndex) And NPCList(NpcIndex).NPCType = eNPCType.GuardiaReal Then
                Call RestarCriminalidad(UserIndex)
                If Not criminal(UserIndex) And .Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)
            End If
            
            If NPCList(NpcIndex).MaestroUser > 0 Then
                Call AllFollowAmo(NPCList(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
                If NPCList(NpcIndex).Stats.Alineacion = 0 Then
                    NPCList(NpcIndex).Movement = NPCList(NpcIndex).flags.OldMovement
                    NPCList(NpcIndex).Hostile = NPCList(NpcIndex).flags.OldHostil
                    NPCList(NpcIndex).flags.AttackedBy = vbNullString
                End If
            End If
            
            Call UserDie(UserIndex)
        End If
    End With
End Sub

Public Sub RestarCriminalidad(ByVal UserIndex As Integer)
    Dim EraCriminal As Boolean
    EraCriminal = criminal(UserIndex)
    
    With UserList(UserIndex).Reputacion
        If .BandidoRep > 0 Then
             .BandidoRep = .BandidoRep - vlASALTO
             If .BandidoRep < 0 Then .BandidoRep = 0
        ElseIf .LadronesRep > 0 Then
             .LadronesRep = .LadronesRep - (vlCAZADOR * 10)
             If .LadronesRep < 0 Then .LadronesRep = 0
        End If
    End With
    
    If EraCriminal And Not criminal(UserIndex) Then
        Call RefreshCharStatus(UserIndex)
    End If
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)
    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) > 0 Then
           If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
            If CheckElementales Or (NPCList(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALFUEGO And NPCList(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
                If NPCList(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0 Then NPCList(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex
                NPCList(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
            End If
           End If
        End If
    Next j
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
        End If
    Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.User) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
    
    If NPCList(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
    
    If UserList(UserIndex).flags.Makro <> 0 Then
        Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
        UserList(UserIndex).flags.Makro = 0
    End If
    
    
    
    With NPCList(NpcIndex)
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            NpcAtacaUser = True
            
            If UserList(UserIndex).flags.Meditando Then
               UserList(UserIndex).flags.Meditando = False
               UserList(UserIndex).Char.fx = 0
               UserList(UserIndex).Char.loops = 0
                        
                Call WriteMeditateToggle(UserIndex)
                        
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
            End If
            
            
            Call CheckPets(NpcIndex, UserIndex, False)
            
            If .Target = 0 Then .Target = UserIndex
            
            If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then
                UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
            End If
        Else
            NpcAtacaUser = False
            Exit Function
        End If
        
        .CanAttack = 0
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
    End With
    
    If NpcImpacto(NpcIndex, UserIndex) Then
        With UserList(UserIndex)
        
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            Call NpcDaño(NpcIndex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            
            '¿Puede envenenar?
            If NPCList(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
        End With
    Else
        Call WriteNPCSwing(UserIndex)
    End If
    
    Call WriteAtaqueNPC(UserIndex, NpcIndex)
    
    '-----Tal vez suba los skills------
    Call SubirSkill(UserIndex, Tacticas)
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
    Dim PoderAtt As Long
    Dim PoderEva As Long
    Dim ProbExito As Long
    
    PoderAtt = NPCList(Atacante).PoderAtaque
    PoderEva = NPCList(Victima).PoderEvasion
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    Dim daño As Integer
    
    With NPCList(Atacante)
        daño = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        NPCList(Victima).Stats.MinHP = NPCList(Victima).Stats.MinHP - daño
        Call SendData(SendTarget.ToPCArea, Victima, PrepareMessageCreateRenderValue(UserList(Victima).Pos.X, UserList(Victima).Pos.Y, daño, DAMAGE_NORMAL))
        
        If NPCList(Victima).Stats.MinHP < 1 Then
            .Movement = .flags.OldMovement
            
            If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil
            End If
            
            If .MaestroUser > 0 Then
                Call FollowAmo(Atacante)
            End If
            
            Call MuereNpc(Victima, .MaestroUser)
        End If
    End With
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
'*************************************************
'Author: Unknown
'Last modified: 01/03/2009
'01/03/2009: ZaMa - Las mascotas no pueden atacar al rey si quedan pretorianos vivos.
'*************************************************
    
    With NPCList(Atacante)
        
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            .CanAttack = 0
            If cambiarMOvimiento Then
                NPCList(Victima).TargetNPC = Atacante
                NPCList(Victima).Movement = TipoAI.NpcAtacaNpc
            End If
        Else
            Exit Sub
        End If
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
        
        If NpcImpactoNpc(Atacante, Victima) Then
            If NPCList(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(NPCList(Victima).flags.Snd2, NPCList(Victima).Pos.X, NPCList(Victima).Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, NPCList(Victima).Pos.X, NPCList(Victima).Pos.Y))
            End If
        
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, NPCList(Victima).Pos.X, NPCList(Victima).Pos.Y))
            End If
            
            Call NpcDañoNpc(Atacante, Victima)
        Else
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, NPCList(Victima).Pos.X, NPCList(Victima).Pos.Y))
                'Call SendData(SendTarget.ToPCArea, Victima, PrepareMessageCreateRenderValue(UserList(VictimNpcIndex).Pos.X, UserList(VictimNpcIndex).Pos.Y, "Fallado", COLOR_DAÑO))
            End If
        End If
    End With
End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

On Error GoTo errHandler
Dim i As Byte

    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        Exit Sub
    End If
    
    'NUEVO SISTEMA DE EXTRACCION DE MINERALES
    Call EsTrabajo(UserIndex, NpcIndex)
    
    If NPCList(NpcIndex).NPCType = eNPCType.Arbol Or NPCList(NpcIndex).NPCType = eNPCType.Yacimiento Then
        Debug.Print UserTrabajando
    
        If UserTrabajando = False Then Exit Sub
    End If
    
    If NPCList(NpcIndex).Numero = NPCReyCastle Then
        If UserList(UserIndex).GuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes clan. Para participar en la conquista de castillos necesitas uno!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
       
         End If
    End If
    
    If NPCList(NpcIndex).flags.ArenasRinkel = 1 And UserList(UserIndex).flags.ArenaRinkel = False Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacar a ese NPC.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If NPCList(NpcIndex).flags.SoloParty = 1 And UserList(UserIndex).PartyId = 0 Then
        Call WriteConsoleMsg(UserIndex, "Para atacar a esta criatura necesitas pertenecer a una Party.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' No podras pasar de mapa por un rato
    Call IntervaloPermiteCambiardeMapa(UserIndex, True)
    
    Call NPCAtacado(NpcIndex, UserIndex)
    
    If UserImpactoNpc(UserIndex, NpcIndex) Then
        If NPCList(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(NPCList(NpcIndex).flags.Snd2, NPCList(NpcIndex).Pos.X, NPCList(NpcIndex).Pos.Y))
        Else
            If SkillTrabajando = eSkill.Mineria Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_MINERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            ElseIf SkillTrabajando = eSkill.talar Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            End If
        End If
        
        Call UserDañoNpc(UserIndex, NpcIndex)
    Else
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(RandomNumber(177, 179), UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        End If
        Call WriteUserSwing(UserIndex)
        'Ha fallado
        UserTrabajando = False
        SkillTrabajando = 0
    End If
    
Exit Sub
    
errHandler:
    Call LogError("Error en UsuarioAtacaNpc. Error " & Err.Number & " : " & Err.Description)
    
End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
    Dim Index As Integer
    Dim AttackPos As worldPos
    
    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
    
    'Check Spell-Magic interval
    If Not IntervaloPermiteMagiaGolpe(UserIndex) Then
        'Check Attack interval
        If Not IntervaloPermiteAtacar(UserIndex) Then
            Exit Sub
        End If
    End If
    
    With UserList(UserIndex)
        'Quitamos stamina
        If .Stats.MinSta >= 10 Then
            Call QuitarSta(UserIndex, RandomNumber(1, 10))
        Else
            If .genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        
        AttackPos = .Pos
        Call HeadtoPos(.Char.heading, AttackPos)
        
        'Exit if not legal
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(RandomNumber(177, 179), .Pos.X, .Pos.Y))
            End If
            Exit Sub
        End If
        
        Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
        
        'Look for user
        If Index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, Index)
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(Index)
            Exit Sub
        End If
        
        Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex
        
        'Look for NPC
        If Index > 0 Then
            If NPCList(Index).Attackable Then
                If NPCList(Index).MaestroUser > 0 And MapInfo(NPCList(Index).Pos.Map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No podés atacar mascotas en zonas seguras", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                Call UsuarioAtacaNpc(UserIndex, Index)
            Else
                Call WriteConsoleMsg(UserIndex, "No podés atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)
            End If
            
            Call WriteUpdateUserStats(UserIndex)
            
            Exit Sub
        End If
        
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(RandomNumber(177, 179), .Pos.X, .Pos.Y))
            End If
        Call WriteUpdateUserStats(UserIndex)
        
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
            
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    End With
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
    
On Error GoTo errHandler

    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long
    Dim PoderAtaque As Long
    Dim UserPoderEvasion As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma As Integer
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)
    
    Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion(VictimaIndex)
    
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
       UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
       UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
       Call WriteProyectil(UserList(AtacanteIndex).Char.CharIndex, VictimaIndex, ObjData(UserList(AtacanteIndex).Invent.MunicionEqpObjIndex).GrhIndex)
    Else
        UserPoderEvasionEscudo = 0
    End If
    
    'Esta usando un arma ???
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
        Else
            PoderAtaque = PoderAtaqueArma(AtacanteIndex)
        End If
    Else
        PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
    
    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
        'Fallo ???
        If Not UsuarioImpacto Then
            ' Chances are rounded
            ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
                  
                Call WriteBlockedWithShieldOther(AtacanteIndex)
                Call WriteBlockedWithShieldUser(VictimaIndex)
                
                Call SubirSkill(VictimaIndex, Defensa)
            End If
        End If
    End If
    
    Call FlushBuffer(VictimaIndex)
    
    Exit Function
    
errHandler:
    Call LogError("Error en UsuarioImpacto. Error " & Err.Number & " : " & Err.Description)
End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

On Error GoTo errHandler

Dim SoundATACK As Byte
Dim UserProtected As Boolean

    UserProtected = Not IntervaloPermiteSerAtacado(VictimaIndex) And UserList(VictimaIndex).flags.NoPuedeSerAtacado
    
    If UserProtected Then Exit Sub

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub
    
    With UserList(AtacanteIndex)
        
        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
           Call WriteConsoleMsg(AtacanteIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
           Exit Sub
        End If
        
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            Call UserDañoUser(AtacanteIndex, VictimaIndex)
        Else
        
            If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
                SoundATACK = SND_SWING
            Else
                SoundATACK = RandomNumber(177, 179)
            End If
        
            ' Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible = 1 Then
                Call EnviarDatosASlot(AtacanteIndex, PrepareMessagePlayWave(SoundATACK, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SoundATACK, .Pos.X, .Pos.Y))
            End If
            
            Call WriteUserSwing(AtacanteIndex)
            Call WriteUserAttackedSwing(VictimaIndex, AtacanteIndex)
        End If
        
        ' No podras pasar de mapa por un rato
        Call IntervaloPermiteCambiardeMapa(AtacanteIndex, True)
        Call IntervaloPermiteCambiardeMapa(VictimaIndex, True)
        
    End With
Exit Sub
    
errHandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.Description)
End Sub

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    
On Error GoTo errHandler

    Dim daño As Long
    Dim absorbido As Long
    Dim defbarco As Integer
    Dim Obj As ObjData
    Dim Resist As Byte

    daño = CalcularDaño(AtacanteIndex)
    If GranPoder = AtacanteIndex Then daño = daño * MultiplicadorGP

    '<Edurne>'En montura...
    If UserList(AtacanteIndex).flags.QueMontura Then _
        daño = daño + UserList(AtacanteIndex).flags.Montura(UserList(AtacanteIndex).flags.QueMontura).Ataque

    Call UserEnvenena(AtacanteIndex, VictimaIndex)
    
    With UserList(AtacanteIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(.Invent.BarcoObjIndex)
             daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
        End If

        If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
             defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If

        If .Invent.WeaponEqpObjIndex > 0 Then
            Resist = ObjData(.Invent.WeaponEqpObjIndex).Refuerzo
        End If

        'Casco
        If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
            Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
            absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If

        'Escudo
        If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
            Obj = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
            absorbido = absorbido + RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If

        'Armadura
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
            Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
            absorbido = absorbido + RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        '<Edurne> 'En montura...
        If UserList(VictimaIndex).flags.QueMontura Then _
            absorbido = absorbido + UserList(VictimaIndex).flags.Montura(UserList(VictimaIndex).flags.QueMontura).Defensa

        absorbido = absorbido - Resist
        Debug.Print "Absorbido: " & absorbido
        daño = daño - absorbido
        
        If daño < 0 Then daño = 1

        If UserList(VictimaIndex).flags.Navegando = 0 Then
            If daño < 100 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            Else
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGREXXL, 0))
            End If
        End If
   
        UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - daño
        
        Call SubirSkill(VictimaIndex, Tacticas)

        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            'Si usa un arma quizas suba "Combate con armas"
            If .Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(.Invent.WeaponEqpObjIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
                    Call SubirSkill(AtacanteIndex, Proyectiles)
                Else
                    'Sube combate con armas.
                    Call SubirSkill(AtacanteIndex, Armas)
                End If
            Else
            'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, Wrestling)
            End If

            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(AtacanteIndex) Then
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño)
                Call SubirSkill(AtacanteIndex, Apuñalar)
            End If
            'e intenta dar un golpe crítico [Pablo (ToxicWaste)]
            Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, daño)
        End If
 
        If Not PuedeApuñalar(AtacanteIndex) Then
            Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateRenderValue(UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, daño, DAMAGE_NORMAL)) ' GSZAO
        End If

        If UserList(VictimaIndex).Stats.MinHP <= 0 Then

            'Store it!
            'Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
            
            Call ContarMuerte(VictimaIndex, AtacanteIndex)

            'Lorwik> Comprobamos si esta en torneo
            Call MuerteEnTorneo(AtacanteIndex, VictimaIndex)

            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If .MascotasIndex(j) > 0 Then
                    If NPCList(.MascotasIndex(j)).Target = VictimaIndex Then
                        NPCList(.MascotasIndex(j)).Target = 0
                        Call FollowAmo(.MascotasIndex(j))
                    End If
                End If
            Next j

            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex, AtacanteIndex)
        Else
            'Está vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)

        End If
    End With
    
    'Controla el nivel del usuario
    Call CheckUserLevel(AtacanteIndex)
    
    Call FlushBuffer(VictimaIndex)

errHandler:

    If Err.Number = 0 Then Exit Sub

    Dim AtacanteNick As String
    Dim VictimaNick As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UserDañoUser. Error " & Err.Number & " : " & Err.Description & " AtacanteIndex: " & _
             AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 10/01/08
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'***************************************************

    If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal As Boolean
    
    If Not criminal(attackerIndex) And Not criminal(VictimIndex) Then
        Call VolverCriminal(attackerIndex)
    End If
    
    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        UserList(VictimIndex).Char.fx = 0
        UserList(VictimIndex).Char.loops = 0
                
        Call WriteMeditateToggle(VictimIndex)
        Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
        Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(UserList(VictimIndex).Char.CharIndex, 0, 0))
    End If
    
    '¿Está trabajando?
    If UserList(VictimIndex).flags.Makro <> 0 Then
        Call WriteMultiMessage(VictimIndex, eMessages.NoTrabaja)
        UserList(VictimIndex).flags.Makro = 0
    End If
    
    Dim EstaenCastillo As Boolean
    Dim i As Byte
    
    '¿Esta en castillo?
    For i = 1 To NUMCASTILLOS
        If UserList(attackerIndex).Pos.Map = Castillos(i).Mapa Then
            EstaenCastillo = True
        Else
            EstaenCastillo = False
        End If
    Next i
    
    '14/02/2016 Lorwik Si la victima tiene el Gran Poder nos saltamos esta parte.
    If Not VictimIndex = GranPoder Or EstaenCastillo = False Then
        EraCriminal = criminal(attackerIndex)
        
        With UserList(attackerIndex).Reputacion
            If Not criminal(VictimIndex) Then
                .BandidoRep = .BandidoRep + vlASALTO
                If .BandidoRep > MAXREP Then .BandidoRep = MAXREP
                
                .NobleRep = .NobleRep / 2
                If .NobleRep < 0 Then .NobleRep = 0
            Else
                .NobleRep = .NobleRep + vlNoble
                If .NobleRep > MAXREP Then .NobleRep = MAXREP
            End If
        End With
        
        If criminal(attackerIndex) Then
            If UserList(attackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(attackerIndex)
            
            If Not EraCriminal Then Call RefreshCharStatus(attackerIndex)
        ElseIf EraCriminal Then
            Call RefreshCharStatus(attackerIndex)
        End If
    End If
    
    Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
    'Reaccion de las mascotas
    Dim iCount As Integer
    
    For iCount = 1 To MAXMASCOTAS
        If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            NPCList(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
            NPCList(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            NPCList(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
        End If
    Next iCount
End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 24/02/2009
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
'24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
'24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
'***************************************************
On Error GoTo errHandler

    'MUY importante el orden de estos "IF"...
    
    Dim EstaenCastillo As Boolean
    Dim i As Byte
    
    '¿Esta en castillo?
    For i = 1 To NUMCASTILLOS
        If UserList(attackerIndex).Pos.Map = Castillos(i).Mapa Then
            EstaenCastillo = True
        End If
    Next i
    
    'Estas muerto no podes atacar
    If UserList(attackerIndex).flags.Muerto = 1 Then
        Call WriteMultiMessage(attackerIndex, eMessages.Muerto)
        PuedeAtacar = False
        Exit Function
    End If
    
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(attackerIndex, "No podés atacar a un espiritu", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    '¿Está trabajando?
    If UserList(attackerIndex).flags.Makro <> 0 Then
        Call WriteConsoleMsg(attackerIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(attackerIndex, VictimIndex)
        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = True
            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            PuedeAtacar = False
            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE
            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 Then
                If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
    End Select
    '12/03/2016 Lorwik> Si esta en castillos nos saltamos estos If
    If EstaenCastillo = False Then
        'Sos un Armada atacando un ciudadano?
        If (Not criminal(VictimIndex)) And (esArmada(attackerIndex)) Then
            Call WriteConsoleMsg(attackerIndex, "Los soldados del Ejército Real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function
        End If
    
        'Tenes puesto el seguro?
        If UserList(attackerIndex).flags.Seguro = 1 Then
            If Not criminal(VictimIndex) Then
                Call WriteConsoleMsg(attackerIndex, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
        End If
    End If
    
    If UserList(attackerIndex).PartyId > 0 Then
        If UserList(VictimIndex).PartyId = UserList(attackerIndex).PartyId Then
            Call WriteConsoleMsg(attackerIndex, "¡No puedes atacar a miembros de tu party.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
    End If
    
    If UserList(attackerIndex).GuildIndex > 0 Then
        If UserList(VictimIndex).GuildIndex = UserList(attackerIndex).GuildIndex Then
            Call WriteConsoleMsg(attackerIndex, "¡No puedes atacar a miembros de tu clan.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
    End If
    
    If EsNewbie(attackerIndex) = True Then
        Call WriteConsoleMsg(attackerIndex, "Los Newbie no pueden atacar a otros usuarios.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacar = False
        Exit Function
    End If
        
    If EsNewbie(VictimIndex) = True Then
        Call WriteConsoleMsg(attackerIndex, "No puedes atacar a los Newbie.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacar = False
        Exit Function
    End If
    
    '13/12/2018 Irongete: Está en una zona segura?
    'Debug.Print permiso_en_zona(attackerIndex)
    If permiso_en_zona(attackerIndex) And permiso_zona.no_atacar Then
        If esArmada(attackerIndex) Then
            If UserList(attackerIndex).Faccion.RecompensasReal > 11 Then
                If UserList(VictimIndex).Pos.Map = 7 Or UserList(VictimIndex).Pos.Map = 8 Or UserList(VictimIndex).Pos.Map = 40 Then
                Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                Exit Function
                End If
            End If
        End If
        If esCaos(attackerIndex) Then
            If UserList(attackerIndex).Faccion.RecompensasCaos > 11 Then
                If UserList(VictimIndex).Pos.Map = 151 Or UserList(VictimIndex).Pos.Map = 156 Then
                Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                Exit Function
                End If
            End If
        End If
        Call WriteConsoleMsg(attackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or _
        MapData(UserList(attackerIndex).Pos.Map, UserList(attackerIndex).Pos.X, UserList(attackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(attackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    PuedeAtacar = True
Exit Function

errHandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.Description)
End Function

Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown Author (Original version)
'Returns True if AttackerIndex can attack the NpcIndex
'Last Modification: 24/01/2007
'24/01/2007 Pablo (ToxicWaste) - Orden y corrección de ataque sobre una mascota y guardias
'14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
'esta función para todo lo referente a ataque a un NPC. Ya sea Magia, Físico o a Distancia.
'***************************************************

    'Estas muerto?
    If UserList(attackerIndex).flags.Muerto = 1 Then
        Call WriteMultiMessage(attackerIndex, eMessages.Muerto)
        PuedeAtacarNPC = False
        Exit Function
    End If
    
    'Sos consejero?
    If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
        'No pueden atacar NPC los Consejeros.
        PuedeAtacarNPC = False
        Exit Function
    End If
    
    'Es una criatura atacable?
    If NPCList(NpcIndex).Attackable = 0 Then
        Call WriteConsoleMsg(attackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacarNPC = False
        Exit Function
    End If
    
    '¿Está trabajando?
    If UserList(attackerIndex).flags.Makro <> 0 Then
        Call WriteConsoleMsg(attackerIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    'Es valida la distancia a la cual estamos atacando?
    If Distancia(UserList(attackerIndex).Pos, NPCList(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
       Call WriteConsoleMsg(attackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
       PuedeAtacarNPC = False
       Exit Function
    End If
    
    'Es una criatura No-Hostil?
    If NPCList(NpcIndex).Hostile = 0 Then
        'Es Guardia del Caos?
        If NPCList(NpcIndex).NPCType = eNPCType.Guardiascaos Then
            'Lo quiere atacar un caos?
            If esCaos(attackerIndex) Then
                Call WriteConsoleMsg(attackerIndex, "No puedes atacar Guardias del Caos siendo Legionario", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        'Es guardia Real?
        ElseIf NPCList(NpcIndex).NPCType = eNPCType.GuardiaReal Then
            'Lo quiere atacar un Armada?
            If esArmada(attackerIndex) Then
                Call WriteConsoleMsg(attackerIndex, "No puedes atacar Guardias Reales siendo Armada Real", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
            'Tienes el seguro puesto?
            If UserList(attackerIndex).flags.Seguro = 1 Then
                Call WriteConsoleMsg(attackerIndex, "Debes quitar el seguro para poder Atacar Guardias Reales utilizando /seg", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            Else
                Call WriteConsoleMsg(attackerIndex, "Atacaste un Guardia Real! Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
                Call VolverCriminal(attackerIndex)
                PuedeAtacarNPC = True
                Exit Function
            End If
    
        'No era un Guardia, asi que es una criatura No-Hostil común.
        'Para asegurarnos que no sea una Mascota:
        ElseIf NPCList(NpcIndex).MaestroUser = 0 Then
            'Si sos ciudadano tenes que quitar el seguro para atacarla.
            If Not criminal(attackerIndex) Then
                'Sos ciudadano, tenes el seguro puesto?
                If UserList(attackerIndex).flags.Seguro = 1 Then
                    Call WriteConsoleMsg(attackerIndex, "Para atacar a este NPC debés quitar el seguro", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                Else
                    'No tiene seguro puesto. Puede atacar pero es penalizado.
                    Call WriteConsoleMsg(attackerIndex, "Atacaste un NPC No-Hostil. Continua haciendolo y serás Criminal.", FontTypeNames.FONTTYPE_INFO)
                    'NicoNZ: Cambio para que al atacar npcs no hostiles no bajen puntos de nobleza
                    'Call DisNobAuBan(attackerIndex, 1000, 1000)
                    Call DisNobAuBan(attackerIndex, 0, 1000)
                    PuedeAtacarNPC = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    Dim i As Byte
    
    '<Edurne>
    'Los castillos no son atacables por sus dueños
    If NPCList(NpcIndex).Numero = NPCReyCastle Or NPCList(NpcIndex).Numero = NPCDefensorFortaleza Or NPCList(NpcIndex).NPCType = 10 Then
        If UserList(attackerIndex).GuildIndex > 0 Then
            For i = 1 To NUMCASTILLOS
                If UserList(attackerIndex).Pos.Map = Castillos(i).Mapa Then
                    If UserList(attackerIndex).GuildIndex = Castillos(i).Dueño Then
                        Call WriteConsoleMsg(attackerIndex, "Este castillo ya pertenece a tu clan.", FontTypeNames.FONTTYPE_INFO)
                        PuedeAtacarNPC = False
                        Exit Function
                    End If
                End If
            Next i
        Else
            Call WriteConsoleMsg(attackerIndex, "Sin clan no pueden atacarse castillos.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarNPC = False
            Exit Function
        End If
    End If
    '/<Edurne>
    
    'Es el NPC mascota de alguien?
    If NPCList(NpcIndex).MaestroUser > 0 Then
        If Not criminal(NPCList(NpcIndex).MaestroUser) Then
            'Es mascota de un Ciudadano.
            If esArmada(attackerIndex) Then
                'El atacante es Armada y esta intentando atacar mascota de un Ciudadano
                Call WriteConsoleMsg(attackerIndex, "Los Armadas no pueden atacar mascotas de Ciudadanos.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
            If Not criminal(attackerIndex) Then
                'El atacante es Ciudadano y esta intentando atacar mascota de un Ciudadano.
                If UserList(attackerIndex).flags.Seguro = 1 Then
                    'El atacante tiene el seguro puesto. No puede atacar.
                    Call WriteConsoleMsg(attackerIndex, "Para atacar mascotas de Ciudadanos debes quitar el seguro.", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                Else
                'El atacante no tiene el seguro puesto. Recibe penalización.
                    Call WriteConsoleMsg(attackerIndex, "Has atacado la Mascota de un Ciudadano. Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
                    Call VolverCriminal(attackerIndex)
                    PuedeAtacarNPC = True
                    Exit Function
                End If
            Else
                'El atacante es criminal y quiere atacar un elemental ciuda, pero tiene el seguro puesto (NicoNZ)
                If UserList(attackerIndex).flags.Seguro = 1 Then
                    Call WriteConsoleMsg(attackerIndex, "Para atacar mascotas de Ciudadanos debes quitar el seguro.", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                End If
            End If
        Else
            'Es mascota de un Criminal.
            If esCaos(NPCList(NpcIndex).MaestroUser) Then
                'Es Caos el Dueño.
                If esCaos(attackerIndex) Then
                    'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
                    Call WriteConsoleMsg(attackerIndex, "Los miembros de la Legión Oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    PuedeAtacarNPC = True
End Function

Public Function TriggerZonaPelea(ByVal origen As Integer, ByVal Destino As Integer) As eTrigger6
'TODO: Pero que rebuscado!!
'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
On Error GoTo errHandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(origen).Pos.Map, UserList(origen).Pos.X, UserList(origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
errHandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.Description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
        End If
        
        If ObjInd > 0 Then
            If ObjData(ObjInd).Envenena = 1 Then
                
                If RandomNumber(1, 100) < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call WriteConsoleMsg(VictimaIndex, UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(AtacanteIndex, "Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
                End If
            End If
        End If
    End If
    
    Call FlushBuffer(VictimaIndex)
End Sub
