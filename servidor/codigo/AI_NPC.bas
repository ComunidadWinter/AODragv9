Attribute VB_Name = "AI"
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

Public Enum TipoAI
    Estatico = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    NpcObjeto = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92

'Damos a los NPCs el mismo rango de visión que un PJ
Private Const NPC_RANGO_VISION_X As Byte = 11
Private Const NPC_RANGO_VISION_Y As Byte = 8

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo AI_NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'AI de los NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Private Sub GuardiasAI(ByVal NPCIndex As Integer, ByVal DelCaos As Boolean)
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim UserProtected As Boolean
    
    With NPCList(NPCIndex)
        For headingloop = eHeading.SOUTH To eHeading.EAST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or headingloop = .Char.heading Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                            '¿ES CRIMINAL?
                            If Not DelCaos Then
                                If criminal(UI) Then
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                    
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            Else
                                If Not criminal(UI) Then
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                      
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If  'not inmovil
        Next headingloop
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

''
' Handles the evil npcs' artificial intelligency.
'
' @param NpcIndex Specifies reference to the npc
Private Sub HostilMalvadoAI(ByVal NPCIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 28/04/2009
'28/04/2009: ZaMa - Now those NPCs who doble attack, have 50% of posibility of casting a spell on user.
'**************************************************************
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim NPCI As Integer
    Dim atacoPJ As Boolean
    Dim UserProtected As Boolean
    
    atacoPJ = False
    
    With NPCList(NPCIndex)
        For headingloop = eHeading.SOUTH To eHeading.EAST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .flags.Paralizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NPCIndex
                    If UI > 0 And Not atacoPJ Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                    
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            atacoPJ = True
                            If .flags.LanzaSpells Then
                                If RandomNumber(0, 100) < 50 Then
                                    If .flags.AtacaDoble Then
                                        If (RandomNumber(0, 1)) Then
                                            If NpcAtacaUser(NPCIndex, UI) Then
                                                Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                            End If
                                            Exit Sub
                                        End If
                                    End If
                                
                                    Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                    Call NpcLanzaUnSpell(NPCIndex, UI)
                                End If
                            End If
                            If NpcAtacaUser(NPCIndex, UI) Then
                                Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                            End If
                            Exit Sub
                        End If
                    ElseIf NPCI > 0 Then
                        If NPCList(NPCI).MaestroUser > 0 And NPCList(NPCI).flags.Paralizado = 0 Then
                            Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                            Call SistemaCombate.NpcAtacaNpc(NPCIndex, NPCI, False)
                            Exit Sub
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub HostilBuenoAI(ByVal NPCIndex As Integer)
    Dim nPos As WorldPos
    Dim headingloop As eHeading
    Dim UI As Integer
    Dim UserProtected As Boolean
    
    With NPCList(NPCIndex)
        For headingloop = eHeading.SOUTH To eHeading.EAST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        If UserList(UI).Name = .flags.AttackedBy Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NPCIndex, UI)
                                End If
                                
                                If NpcAtacaUser(NPCIndex, UI) Then
                                    Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub IrUsuarioCercano(ByVal NPCIndex As Integer)
    Dim tHeading As Byte
    Dim UI As Integer
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim i As Long
    Dim UserProtected As Boolean
    
    With NPCList(NPCIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= NPC_RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= NPC_RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        
                        If UserList(UI).flags.Muerto = 0 Then
                            If Not UserProtected Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NPCIndex, UI)
                                Exit Sub
                            End If
                        End If
                        
                    End If
                End If
            Next i
        Else
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= NPC_RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= NPC_RANGO_VISION_Y Then
                        
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        
                        If UserList(UI).flags.Muerto = 0 Then
                            If UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Or (NPCList(NPCIndex).flags.VerInvi = 1) And Not UserProtected Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NPCIndex, UI)
                                If Not NPCList(NPCIndex).PFINFO.PathLenght > 0 Then tHeading = FindDirection(NPCList(NPCIndex).Pos, .Pos)
                                If tHeading = 0 Then
                                    Call PathFindingAI(NPCIndex)
                                    If Not ReCalculatePath(NPCIndex) Then
                                        If Not PathEnd(NPCIndex) Then
                                            Call FollowPath(NPCIndex)
                                        Else
                                            NPCList(NPCIndex).PFINFO.PathLenght = 0
                                        End If
                                    End If
                                Else
                                    If Not NPCList(NPCIndex).PFINFO.PathLenght > 0 Then Call MoveNPCChar(NPCIndex, tHeading)
                                    Exit Sub
                                End If
                                Exit Sub
                            End If
                        End If
                        
                    End If
                End If
            Next i
            
            'Si llega aca es que no había ningún usuario cercano vivo.
            'A bailar. Pablo (ToxicWaste)
            If RandomNumber(0, 10) = 0 Then
                Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))
            End If
        End If
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

''
' Makes a Pet / Summoned Npc to Follow an enemy
'
' @param NpcIndex Specifies reference to the npc
Private Sub SeguirAgresor(ByVal NPCIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: Marco Vanotti (MarKoxX)
'Last Modify Date: 08/16/2008
'08/16/2008: MarKoxX - Now pets that do melé attacks have to be near the enemy to attack.
'**************************************************************
    Dim tHeading As Byte
    Dim UI As Integer
    
    Dim i As Long
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer

    With NPCList(NPCIndex)
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select

            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)

                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= NPC_RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= NPC_RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then

                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacará a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado", FontTypeNames.FONTTYPE_INFO)
                                    Call FlushBuffer(.MaestroUser)
                                    .flags.AttackedBy = vbNullString
                                    Exit Sub
                                End If
                            End If

                            If UserList(UI).flags.Muerto = 0 Then
                                If (NPCList(NPCIndex).flags.VerInvi = 1) Or UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                                    If .flags.LanzaSpells > 0 Then
                                         Call NpcLanzaUnSpell(NPCIndex, UI)
                                    Else
                                       If Distancia(UserList(UI).Pos, NPCList(NPCIndex).Pos) <= 1 Then
                                           ' TODO : Set this a separate AI for Elementals and Druid's pets
                                           If NPCList(NPCIndex).Numero <> 92 Then
                                               Call NpcAtacaUser(NPCIndex, UI)
                                           End If
                                       End If
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                    End If
                End If
                
            Next i
        Else
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= NPC_RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= NPC_RANGO_VISION_Y Then
                        
                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacará a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado", FontTypeNames.FONTTYPE_INFO)
                                    Call FlushBuffer(.MaestroUser)
                                    .flags.AttackedBy = vbNullString
                                    Call FollowAmo(NPCIndex)
                                    Exit Sub
                                End If
                            End If
                            
                            If UserList(UI).flags.Muerto = 0 Then
                                If (NPCList(NPCIndex).flags.VerInvi = 1) Or UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                                    If .flags.LanzaSpells > 0 Then
                                           Call NpcLanzaUnSpell(NPCIndex, UI)
                                    Else
                                       If Distancia(UserList(UI).Pos, NPCList(NPCIndex).Pos) <= 1 Then
                                           ' TODO : Set this a separate AI for Elementals and Druid's pets
                                           If NPCList(NPCIndex).Numero <> 92 Then
                                               Call NpcAtacaUser(NPCIndex, UI)
                                           End If
                                       End If
                                    End If
                                    
                                    tHeading = FindDirection(.Pos, UserList(UI).Pos)
                                    Call MoveNPCChar(NPCIndex, tHeading)
                                    
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                    End If
                End If
                
            Next i
        End If
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub RestoreOldMovement(ByVal NPCIndex As Integer)
    With NPCList(NPCIndex)
        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
        End If
    End With
End Sub

Private Sub PersigueCiudadano(ByVal NPCIndex As Integer)
    Dim UI As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim UserProtected As Boolean
    
    With NPCList(NPCIndex)
        For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
            UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
            'Is it in it's range of vision??
            If Abs(UserList(UI).Pos.X - .Pos.X) <= NPC_RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= NPC_RANGO_VISION_Y Then
                    
                    If Not criminal(UI) Then
                        
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                    
                        If UserList(UI).flags.Muerto = 0 Then
                            If UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Or Not (.flags.VerInvi = 1) Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NPCIndex, UI)
                                End If
                                    If Not NPCList(NPCIndex).PFINFO.PathLenght > 0 Then tHeading = FindDirection(NPCList(NPCIndex).Pos, .Pos)
                                    If tHeading = 0 Then
                                        Call PathFindingAI(NPCIndex)
                                        If Not ReCalculatePath(NPCIndex) Then
                                            If Not PathEnd(NPCIndex) Then
                                                Call FollowPath(NPCIndex)
                                            Else
                                                NPCList(NPCIndex).PFINFO.PathLenght = 0
                                            End If
                                        End If
                                    Else
                                        If Not NPCList(NPCIndex).PFINFO.PathLenght > 0 Then Call MoveNPCChar(NPCIndex, tHeading)
                                        Exit Sub
                                    End If
    
                                Exit Sub
                            End If
                        Else
                            If InMapBounds(NPCList(NPCIndex).Orig.Map, NPCList(NPCIndex).Orig.X, NPCList(NPCIndex).Orig.Y) And NPCList(NPCIndex).Numero = Guardias Then
                                tHeading = FindDirection(.Pos, .Orig)
                                Call MoveNPCChar(NPCIndex, tHeading)
                            End If
                        End If
                    End If
                    
               End If
            End If
            
        Next i
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub PersigueCriminal(ByVal NPCIndex As Integer)
    Dim UI As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim UserProtected As Boolean
    
    With NPCList(NPCIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= NPC_RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= NPC_RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        If criminal(UI) Then
                        
                            UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        
                            If UserList(UI).flags.Muerto = 0 Then
                                If UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Or (.flags.VerInvi = 1) And Not UserProtected Then
                                    If .flags.LanzaSpells > 0 Then
                                          Call NpcLanzaUnSpell(NPCIndex, UI)
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                   End If
                End If
                    
            Next i
        Else
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= NPC_RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= NPC_RANGO_VISION_Y Then
                        
                        If criminal(UI) Then
                        
                            UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        
                            If UserList(UI).flags.Muerto = 0 Then
                                If UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Or (NPCList(NPCIndex).flags.VerInvi = 1) And Not UserProtected Then
                                    If .flags.LanzaSpells > 0 Then
                                        Call NpcLanzaUnSpell(NPCIndex, UI)
                                    End If
                                    If .flags.Inmovilizado = 1 Then Exit Sub
                                    
                                    If Not NPCList(NPCIndex).PFINFO.PathLenght > 0 Then tHeading = FindDirection(NPCList(NPCIndex).Pos, .Pos)
                                    If tHeading = 0 Then
                                        Call PathFindingAI(NPCIndex)
                                        If Not ReCalculatePath(NPCIndex) Then
                                            If Not PathEnd(NPCIndex) Then
                                                Call FollowPath(NPCIndex)
                                            Else
                                                NPCList(NPCIndex).PFINFO.PathLenght = 0
                                            End If
                                        End If
                                    Else
                                        If Not NPCList(NPCIndex).PFINFO.PathLenght > 0 Then Call MoveNPCChar(NPCIndex, tHeading)
                                        Exit Sub
                                    End If
                                    
                                    Exit Sub
                                End If
                            Else
                                If InMapBounds(NPCList(NPCIndex).Orig.Map, NPCList(NPCIndex).Orig.X, NPCList(NPCIndex).Orig.Y) And NPCList(NPCIndex).Numero = Guardias Then
                                    tHeading = FindDirection(.Pos, .Orig)
                                    Call MoveNPCChar(NPCIndex, tHeading)
                                End If
                            End If
                        End If
                        
                   End If
                End If
                
            Next i
        End If
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub SeguirAmo(ByVal NPCIndex As Integer)
    Dim tHeading As Byte
    Dim UI As Integer
    
    With NPCList(NPCIndex)
        If .Target = 0 And .TargetNPC = 0 Then
            UI = .MaestroUser
            
            'Is it in it's range of vision??
            If Abs(UserList(UI).Pos.X - .Pos.X) <= NPC_RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= NPC_RANGO_VISION_Y Then
                    If UserList(UI).flags.Muerto = 0 _
                            And UserList(UI).flags.invisible = 0 _
                            And UserList(UI).flags.Oculto = 0 _
                            And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                        tHeading = FindDirection(.Pos, UserList(UI).Pos)
                        Call MoveNPCChar(NPCIndex, tHeading)
                        Exit Sub
                    End If
                End If
            End If
        End If
    End With
    
    Call RestoreOldMovement(NPCIndex)
End Sub

Private Sub AiNpcAtacaNpc(ByVal NPCIndex As Integer)
    Dim tHeading As Byte
    Dim X As Long
    Dim Y As Long
    Dim NI As Integer
    Dim bNoEsta As Boolean
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    
    With NPCList(NPCIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For Y = .Pos.Y To .Pos.Y + SignoNS * NPC_RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .Pos.X To .Pos.X + SignoEO * NPC_RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, X, Y).NPCIndex
                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True
                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NPCIndex, NI)
                                    If NPCList(NI).NPCType = DRAGON Then
                                        NPCList(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NPCIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, NPCList(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NPCIndex, NI)
                                    End If
                                 End If
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        Else
            For Y = .Pos.Y - NPC_RANGO_VISION_Y To .Pos.Y + NPC_RANGO_VISION_Y
                For X = .Pos.X - NPC_RANGO_VISION_Y To .Pos.X + NPC_RANGO_VISION_Y
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                       NI = MapData(.Pos.Map, X, Y).NPCIndex
                       If NI > 0 Then
                            If .TargetNPC = NI Then
                                 bNoEsta = True
                                 If .Numero = ELEMENTALFUEGO Then
                                     Call NpcLanzaUnSpellSobreNpc(NPCIndex, NI)
                                     If NPCList(NI).NPCType = DRAGON Then
                                        NPCList(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NPCIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, NPCList(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NPCIndex, NI)
                                    End If
                                 End If
                                 If .flags.Inmovilizado = 1 Then Exit Sub
                                 If .TargetNPC = 0 Then Exit Sub
                                 tHeading = FindDirection(.Pos, NPCList(MapData(.Pos.Map, X, Y).NPCIndex).Pos)
                                 Call MoveNPCChar(NPCIndex, tHeading)
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NPCIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            End If
        End If
    End With
End Sub

Public Sub AiNpcObjeto(ByVal NPCIndex As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 14/09/2009 (ZaMa)
'***************************************************
    Dim UserIndex As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim UserProtected As Boolean
    
    With NPCList(NPCIndex)
        For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
            UserIndex = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
            
            'Is it in it's range of vision??
            If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= NPC_RANGO_VISION_X Then
                If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= NPC_RANGO_VISION_Y Then
                    
                    With UserList(UserIndex)
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                        
                        If .flags.Muerto = 0 And .flags.invisible = 0 And _
                            .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                            
                            ' No quiero que ataque siempre al primero
                            If RandomNumber(1, 3) < 3 Then
                                If NPCList(NPCIndex).flags.LanzaSpells > 0 Then
                                     Call NpcLanzaUnSpell(NPCIndex, UserIndex)
                                End If
                            
                                Exit Sub
                            End If
                        End If
                    End With
               End If
            End If
            
        Next i
    End With

End Sub

Sub NPCAI(ByVal NPCIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: ZaMa
'Last Modify Date: 15/11/2009
'08/16/2008: MarKoxX - Now pets that do melé attacks have to be near the enemy to attack.
'15/11/2009: ZaMa - Implementacion de npc objetos ai.
'**************************************************************
On Error GoTo ErrorHandler
    With NPCList(NPCIndex)
    
        'Irongete: Lista de NPC's a los que le desactivo la AI
        'Nota: Esto está para que el rey y la puerta del castillo no quieran atacar a jugadores y desparezcan.
        If .Numero = 514 Then Exit Sub 'Rey del castillo
        If .Numero = 603 Then Exit Sub 'Defensor de la Fortaleza
        If .Numero = 595 Then Exit Sub 'Puerta del Castillo
        If .Numero = 624 Then Exit Sub 'Dummy
        
        
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If .MaestroUser = 0 Then
            'Busca a alguien para atacar
            '¿Es un guardia?
            If .NPCType = eNPCType.GuardiaReal Then
                Call GuardiasAI(NPCIndex, False)
            ElseIf .NPCType = eNPCType.Guardiascaos Then
                Call GuardiasAI(NPCIndex, True)
            ElseIf .Hostile And .Stats.Alineacion <> 0 Then
                Call HostilMalvadoAI(NPCIndex)
            ElseIf .Hostile And .Stats.Alineacion = 0 Then
                Call HostilBuenoAI(NPCIndex)
            End If
        Else
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case .Movement
            Case TipoAI.MueveAlAzar
                If .flags.Inmovilizado = 1 Then Exit Sub
                If .NPCType = eNPCType.GuardiaReal Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))
                    End If
                    Call PersigueCriminal(NPCIndex)
                ElseIf .NPCType = eNPCType.Guardiascaos Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))
                    End If
                    Call PersigueCiudadano(NPCIndex)
                Else
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))
                    End If
                End If
            
            'Va hacia el usuario cercano
            Case TipoAI.NpcMaloAtacaUsersBuenos
                Call IrUsuarioCercano(NPCIndex)
            
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(NPCIndex)
            
            'Persigue criminales
            Case TipoAI.GuardiasAtacanCriminales
                Call PersigueCriminal(NPCIndex)
            
            Case TipoAI.SigueAmo
                If .flags.Inmovilizado = 1 Then Exit Sub
                Call SeguirAmo(NPCIndex)
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))
                End If
            
            Case TipoAI.NpcAtacaNpc
                Call AiNpcAtacaNpc(NPCIndex)
            
            Case TipoAI.NpcObjeto
                Call AiNpcObjeto(NPCIndex)
                
            Case TipoAI.NpcPathfinding
                If .flags.Inmovilizado = 1 Then Exit Sub
                If ReCalculatePath(NPCIndex) Then
                    Call PathFindingAI(NPCIndex)
                    'Existe el camino?
                    If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(NPCIndex, RandomNumber(eHeading.SOUTH, eHeading.EAST))
                    End If
                Else
                    If Not PathEnd(NPCIndex) Then
                        Call FollowPath(NPCIndex)
                    Else
                        .PFINFO.PathLenght = 0
                    End If
                End If
        End Select
    End With
Exit Sub

ErrorHandler:
    Call LogError("NPCAI " & NPCList(NPCIndex).Name & " " & NPCList(NPCIndex).MaestroUser & " " & NPCList(NPCIndex).MaestroNpc & " mapa:" & NPCList(NPCIndex).Pos.Map & " x:" & NPCList(NPCIndex).Pos.X & " y:" & NPCList(NPCIndex).Pos.Y & " Mov:" & NPCList(NPCIndex).Movement & " TargU:" & NPCList(NPCIndex).Target & " TargN:" & NPCList(NPCIndex).TargetNPC)
    Dim MiNPC As npc
    MiNPC = NPCList(NPCIndex)
    Call QuitarNPC(NPCIndex)
    Call ReSpawnNpc(MiNPC)
End Sub

Function UserNear(ByVal NPCIndex As Integer) As Boolean
'#################################################################
'Returns True if there is an user adjacent to the npc position.
'#################################################################
    UserNear = Not Int(Distance(NPCList(NPCIndex).Pos.X, NPCList(NPCIndex).Pos.Y, UserList(NPCList(NPCIndex).PFINFO.targetUser).Pos.X, UserList(NPCList(NPCIndex).PFINFO.targetUser).Pos.Y)) > 1
End Function

Function ReCalculatePath(ByVal NPCIndex As Integer) As Boolean
'#################################################################
'Returns true if we have to seek a new path
'#################################################################
    If NPCList(NPCIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NPCIndex) And NPCList(NPCIndex).PFINFO.PathLenght = NPCList(NPCIndex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True
    End If
End Function

Function PathEnd(ByVal NPCIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Returns if the npc has arrived to the end of its path
'#################################################################
    PathEnd = NPCList(NPCIndex).PFINFO.CurPos = NPCList(NPCIndex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal NPCIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Moves the npc.
'#################################################################
    Dim tmpPos As WorldPos
    Dim tHeading As Byte
    
    tmpPos.Map = NPCList(NPCIndex).Pos.Map
    tmpPos.X = NPCList(NPCIndex).PFINFO.Path(NPCList(NPCIndex).PFINFO.CurPos).Y ' invertí las coordenadas
    tmpPos.Y = NPCList(NPCIndex).PFINFO.Path(NPCList(NPCIndex).PFINFO.CurPos).X
    
    'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
    
    tHeading = FindDirection(NPCList(NPCIndex).Pos, tmpPos)
    
    MoveNPCChar NPCIndex, tHeading
    
    NPCList(NPCIndex).PFINFO.CurPos = NPCList(NPCIndex).PFINFO.CurPos + 1
End Function

Function PathFindingAI(ByVal NPCIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock / 11-07-02
'www.geocities.com/gmorgolock
'morgolock@speedy.com.ar
'This function seeks the shortest path from the Npc
'to the user's location.
'#################################################################
    Dim Y As Long
    Dim X As Long
    
    For Y = NPCList(NPCIndex).Pos.Y - 10 To NPCList(NPCIndex).Pos.Y + 10    'Makes a loop that looks at
         For X = NPCList(NPCIndex).Pos.X - 10 To NPCList(NPCIndex).Pos.X + 10   '5 tiles in every direction
            
             'Make sure tile is legal
             If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                
                 'look for a user
                 If MapData(NPCList(NPCIndex).Pos.Map, X, Y).UserIndex > 0 Then
                     'Move towards user
                      Dim tmpUserIndex As Integer
                      tmpUserIndex = MapData(NPCList(NPCIndex).Pos.Map, X, Y).UserIndex
                      With UserList(tmpUserIndex)
                        If .flags.Muerto = 0 Then
                            If (NPCList(NPCIndex).flags.VerInvi = 1) Or .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                                'We have to invert the coordinates, this is because
                                'ORE refers to maps in converse way of my pathfinding
                                'routines.
                                NPCList(NPCIndex).PFINFO.Target.X = UserList(tmpUserIndex).Pos.Y
                                NPCList(NPCIndex).PFINFO.Target.Y = UserList(tmpUserIndex).Pos.X 'ops!
                                NPCList(NPCIndex).PFINFO.targetUser = tmpUserIndex
                                Call SeekPath(NPCIndex)
                                
                                'Si es un MonsterDrag y se aleja 10 tiles de su OrigPos se le devuelve.
                                If NPCList(NPCIndex).NPCType = 11 Then
                                    If NPCList(NPCIndex).Pos.X = NPCList(NPCIndex).Orig.X - 10 Or NPCList(NPCIndex).Pos.X = NPCList(NPCIndex).Orig.X + 10 Or _
                                        NPCList(NPCIndex).Pos.Y = NPCList(NPCIndex).Orig.Y - 10 Or NPCList(NPCIndex).Pos.Y = NPCList(NPCIndex).Orig.Y + 10 Then
                                            Dim DragPos As WorldPos
                                            DragPos.Map = NPCList(NPCIndex).Orig.Map
                                            DragPos.X = NPCList(NPCIndex).Orig.X
                                            DragPos.Y = NPCList(NPCIndex).Orig.Y
                                            
                                            Call SpawnNpc(NPCList(NPCIndex).Numero, DragPos, True, False, True)
                                            Call SendData(SendTarget.ToPCArea, NPCIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
                                            Call QuitarNPC(NPCIndex)
                                        End If
                                End If
                                
                                Exit Function
                            End If
                        End If
                    End With
                End If
            End If
        Next X
    Next Y
End Function

Sub NpcLanzaUnSpell(ByVal NPCIndex As Integer, ByVal UserIndex As Integer)
    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 And (NPCList(NPCIndex).flags.VerInvi = 0) Then Exit Sub
    
    Dim k As Integer
    k = RandomNumber(1, NPCList(NPCIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NPCIndex, UserIndex, NPCList(NPCIndex).Spells(k))
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NPCIndex As Integer, ByVal TargetNPC As Integer)
    Dim k As Integer
    k = RandomNumber(1, NPCList(NPCIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NPCIndex, TargetNPC, NPCList(NPCIndex).Spells(k))
End Sub


