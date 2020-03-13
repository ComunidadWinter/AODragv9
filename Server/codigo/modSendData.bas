Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget
    ToAll = 1
    toMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToPartyArea
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToHigherAdmins
    ToGMsArea
    ToUsersAreaButGMs
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
'Last Modify Date: 01/08/2007
'Last modified by: (liquid)
'**************************************************************
On Error Resume Next
    Dim LoopC As Long
    Dim Map As Integer
    
    Select Case sndRoute
        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                   End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToAll
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToAllButIndex
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)
            Exit Sub
          
        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGuildMembers
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            Exit Sub
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPartyArea
            Call SendToUserPartyArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToDiosesYclan
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
                LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            Wend
            
            Exit Sub
        
        Case SendTarget.ToConsejo
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToConsejoCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToRolesMasters
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCiudadanos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If Not criminal(LoopC) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCriminales
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If criminal(LoopC) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToReal
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCiudadanosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If Not criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCriminalesYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToRealYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCaosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToHigherAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                   End If
                End If
            Next LoopC
            Exit Sub
            
        Case SendTarget.ToGMsArea
            Call SendToGMsArea(sndIndex, sndData)
            Exit Sub
            
        Case SendTarget.ToUsersAreaButGMs
            Call SendToUsersAreaButGMs(sndIndex, sndData)
            Exit Sub
            
    End Select
End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer
 
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
 
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
 
    If Not MapaValido(Map) Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
    
        If ModAreas.estoyAreaUser(TempIndex, UserIndex) >= 2 Then
                If UserList(TempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
    Next LoopC
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim TempIndex As Integer
 
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
 
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

    If Not MapaValido(Map) Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If ModAreas.estoyAreaUser(TempIndex, UserIndex) >= 2 Then
                If TempIndex <> UserIndex Then
                    If UserList(TempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(TempIndex, sdData)
                    End If
                End If
            End If
    Next LoopC
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer
 
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
 
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
 
    If Not MapaValido(Map) Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
    
            If ModAreas.estoyAreaUser(TempIndex, UserIndex) >= 2 Then
                'Dead and admins read
                If UserList(TempIndex).ConnIDValida = True And (UserList(TempIndex).flags.Muerto = 1 Or (UserList(TempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
                    Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
    Next LoopC
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer
 
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
 
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
 
    If Not MapaValido(Map) Then Exit Sub
 
    If UserList(UserIndex).GuildIndex = 0 Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
    
            If ModAreas.estoyAreaUser(TempIndex, UserIndex) >= 2 Then
                If UserList(TempIndex).ConnIDValida And (UserList(TempIndex).GuildIndex = UserList(UserIndex).GuildIndex Or ((UserList(TempIndex).flags.Privilegios And PlayerType.Dios) And (UserList(TempIndex).flags.Privilegios And PlayerType.RoleMaster) = 0)) Then
                    Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
    Next LoopC
End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer
 
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
 
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
 
    If Not MapaValido(Map) Then Exit Sub
 
    If UserList(UserIndex).PartyIndex = 0 Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
    
            If ModAreas.estoyAreaUser(TempIndex, UserIndex) >= 2 Then
                If UserList(TempIndex).ConnIDValida And UserList(TempIndex).PartyId = UserList(UserIndex).PartyId Then
                    Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
    Next LoopC
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer
 
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
 
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
 
    If Not MapaValido(Map) Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
    
            If ModAreas.estoyAreaUser(TempIndex, UserIndex) >= 2 Then
                If UserList(TempIndex).ConnIDValida Then
                    If UserList(TempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then _
                        Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
    Next LoopC
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim TempIndex As Integer
 
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
 
    Map = NPCList(NpcIndex).Pos.Map
    AreaX = NPCList(NpcIndex).AreasInfo.AreaPerteneceX
    AreaY = NPCList(NpcIndex).AreasInfo.AreaPerteneceY
 
    If Not MapaValido(Map) Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
    
            If ModAreas.npcAreaUser(NpcIndex, TempIndex) >= 2 Then
                If UserList(TempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
    Next LoopC
End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim TempIndex As Integer
 
    AreaX = idAreaX(AreaX)
    AreaY = idAreaY(AreaY)
 
    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
            If toUserArea(TempIndex, AreaX, AreaY) >= 2 Then
                If UserList(TempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(TempIndex, sdData)
                End If
            End If
    Next LoopC
End Sub


Public Sub SendToMap(ByVal Map As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 5/24/2007
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer
    
    If Not MapaValido(Map) Then Exit Sub

     For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(TempIndex).ConnIDValida Then
            Call EnviarDatosASlot(TempIndex, sdData)
        End If
    Next LoopC
End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 5/24/2007
'
'**************************************************************
    Dim LoopC As Long
    Dim Map As Integer
    Dim TempIndex As Integer
    
    Map = UserList(UserIndex).Pos.Map
    
    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If TempIndex <> UserIndex And UserList(TempIndex).ConnIDValida Then
            Call EnviarDatosASlot(TempIndex, sdData)
        End If
    Next LoopC
End Sub

Private Sub SendToGMsArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Torres Patricio(Pato)
'Last Modify Date: 10/17/2009
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer
 
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
 
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
 
    If Not MapaValido(Map) Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
    
        With UserList(TempIndex)
            If ModAreas.estoyAreaUser(TempIndex, UserIndex) >= 2 Then

                    If .ConnIDValida Then
                        ' Exclusivo para dioses, admins y gms
                        If (.flags.Privilegios And Not PlayerType.User And Not PlayerType.Consejero _
                            And Not PlayerType.RoleMaster) = .flags.Privilegios Then
                            Call EnviarDatosASlot(TempIndex, sdData)
                        End If
                End If
            End If
        End With
    Next LoopC
End Sub

Private Sub SendToUsersAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Torres Patricio(Pato)
'Last Modify Date: 10/17/2009
'
'**************************************************************
    Dim LoopC As Long
    Dim TempIndex As Integer
 
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
 
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
 
    If Not MapaValido(Map) Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        TempIndex = ConnGroups(Map).UserEntrys(LoopC)
    
            If ModAreas.estoyAreaUser(TempIndex, UserIndex) >= 2 Then
                If UserList(TempIndex).ConnIDValida Then
                    If UserList(TempIndex).flags.Privilegios And PlayerType.User Then
                        Call EnviarDatosASlot(TempIndex, sdData)
                    End If
                End If
            End If
    Next LoopC
End Sub
