Attribute VB_Name = "modClanes"
Option Explicit

Public Guilds As New Collection
Public Sub SubirLevelClan(ByVal UserIndex As Integer)

Select Case UserList(UserIndex).GuildRef.Nivel + 1

Case 1
If UserList(UserIndex).Stats.GLD > 999999 Then
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 1000000
Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
Call SendData(ToAll, 0, 0, "||¡¡¡ El Clan " & UserList(UserIndex).GuildInfo.GuildName & " ha subido al Nivel " & UserList(UserIndex).GuildRef.Nivel + 1 & " !!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
UserList(UserIndex).GuildRef.Nivel = UserList(UserIndex).GuildRef.Nivel + 1
Else
Call SendData(ToIndex, UserIndex, 0, "||Debes tener 1 Millon de Oro." & "´" & FontTypeNames.FONTTYPE_GUILD)
Exit Sub
    End If





Case 2

    If TieneObjetos(1095, 1, UserIndex) And UserList(UserIndex).Stats.GLD > 2999999 Then
Call QuitarObjetos(1095, 1, UserIndex)
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 3000000
Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
Call SendData(ToAll, 0, 0, "||¡¡¡ El Clan " & UserList(UserIndex).GuildInfo.GuildName & " ha subido al Nivel " & UserList(UserIndex).GuildRef.Nivel + 1 & " !!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
   UserList(UserIndex).GuildRef.Nivel = UserList(UserIndex).GuildRef.Nivel + 1
    Else
Call SendData(ToIndex, UserIndex, 0, "||Debes tener un Huevo Dragón y 3 Millones de Oro." & "´" & FontTypeNames.FONTTYPE_GUILD)
Exit Sub
    End If

Case 3

    If TieneObjetos(1095, 3, UserIndex) And TieneObjetos(1096, 1, UserIndex) And UserList(UserIndex).Stats.GLD > 14999999 Then
Call QuitarObjetos(1095, 3, UserIndex)
Call QuitarObjetos(1096, 1, UserIndex)
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 15000000
Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
Call SendData(ToAll, 0, 0, "||¡¡¡ El Clan " & UserList(UserIndex).GuildInfo.GuildName & " ha subido al Nivel " & UserList(UserIndex).GuildRef.Nivel + 1 & " !!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
   UserList(UserIndex).GuildRef.Nivel = UserList(UserIndex).GuildRef.Nivel + 1
   Else
Call SendData(ToIndex, UserIndex, 0, "||Debes tener Tres Huevo Dragón, Un Diamante de Sangre, 15 Millones de Oro." & "´" & FontTypeNames.FONTTYPE_GUILD)
Exit Sub
    End If


Case 4

    If TieneObjetos(1095, 5, UserIndex) And TieneObjetos(1096, 2, UserIndex) And UserList(UserIndex).Stats.GLD > 74999999 And UserList(UserIndex).GuildRef.Reputation > 2999 And UserList(UserIndex).GuildRef.Reputation > 3200 Then
Call QuitarObjetos(1095, 5, UserIndex)
Call QuitarObjetos(1096, 2, UserIndex)
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 75000000
Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
Call SendData(ToAll, 0, 0, "||¡¡¡ El Clan " & UserList(UserIndex).GuildInfo.GuildName & " ha subido al Nivel " & UserList(UserIndex).GuildRef.Nivel + 1 & " !!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
   UserList(UserIndex).GuildRef.Nivel = UserList(UserIndex).GuildRef.Nivel + 1
    Else
Call SendData(ToIndex, UserIndex, 0, "||Debes tener Cinco Huevos Dragón, Dos Diamante de Sangre, 3000 Puntos de Clan y 75 Millones de Oro." & "´" & FontTypeNames.FONTTYPE_GUILD)
Exit Sub
    End If

Case 5

    If TieneObjetos(1095, 10, UserIndex) And TieneObjetos(1096, 5, UserIndex) And UserList(UserIndex).Stats.GLD > 149999999 And UserList(UserIndex).GuildRef.Reputation > 4999 Then
Call QuitarObjetos(1095, 10, UserIndex)
Call QuitarObjetos(1096, 5, UserIndex)
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 150000000
Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
Call SendData(ToAll, 0, 0, "||¡¡¡ El Clan " & UserList(UserIndex).GuildInfo.GuildName & " ha subido al Nivel " & UserList(UserIndex).GuildRef.Nivel + 1 & " !!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
  UserList(UserIndex).GuildRef.Nivel = UserList(UserIndex).GuildRef.Nivel + 1
    Else
Call SendData(ToIndex, UserIndex, 0, "||Debes tener Diez Huevos Dragón, Cinco Diamante de Sangre, 5000 Puntos de Clan y 150 Millones de Oro." & "´" & FontTypeNames.FONTTYPE_GUILD)
Exit Sub
    End If
'pluto:6.8
Case 6
   If TieneObjetos(1095, 12, UserIndex) And TieneObjetos(1096, 8, UserIndex) And UserList(UserIndex).Stats.GLD > 199999999 And UserList(UserIndex).GuildRef.Reputation > 7200 Then
Call QuitarObjetos(1095, 12, UserIndex)
Call QuitarObjetos(1096, 8, UserIndex)
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 200000000
Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
Call SendData(ToAll, 0, 0, "||¡¡¡ El Clan " & UserList(UserIndex).GuildInfo.GuildName & " ha subido al Nivel " & UserList(UserIndex).GuildRef.Nivel + 1 & " !!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
  UserList(UserIndex).GuildRef.Nivel = UserList(UserIndex).GuildRef.Nivel + 1
    Else
Call SendData(ToIndex, UserIndex, 0, "||Debes tener Doce Huevos Dragón, Ocho Diamante de Sangre, 7200 Puntos de Clan y 200 Millones de Oro." & "´" & FontTypeNames.FONTTYPE_GUILD)

Exit Sub
End If


End Select

End Sub


Public Sub ComputeVote(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
Dim myGuild As cGuild

Set myGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If myGuild Is Nothing Then Exit Sub

If Not myGuild.Elections Then
   Call SendData(ToIndex, UserIndex, 0, "||Aun no es periodo de elecciones." & "´" & FontTypeNames.FONTTYPE_GUILD)
   Exit Sub
End If

If UserList(UserIndex).GuildInfo.YaVoto = 1 Then
   Call SendData(ToIndex, UserIndex, 0, "||Ya has votado!!! solo se permite un voto por miembro." & "´" & FontTypeNames.FONTTYPE_GUILD)
   Exit Sub
End If

If Not myGuild.IsMember(rdata) Then
   Call SendData(ToIndex, UserIndex, 0, "||No hay ningun miembro con ese nombre." & "´" & FontTypeNames.FONTTYPE_GUILD)
   Exit Sub
End If


Call myGuild.Votes.Add(rdata)
UserList(UserIndex).GuildInfo.YaVoto = 1
Call SendData(ToIndex, UserIndex, 0, "||Tu voto ha sido contabilizado." & "´" & FontTypeNames.FONTTYPE_GUILD)

Exit Sub
fallo:
Call LogError("computevote " & Err.number & " D: " & Err.Description)


End Sub

Public Sub ResetUserVotes(ByRef myGuild As cGuild)

On Error GoTo fallo

Dim k As Integer, index As Integer
Dim userfile As String
For k = 1 To myGuild.Members.Count
       
    index = DameUserIndexConNombre(myGuild.Members(k))
    If index <> 0 Then 'is online
        UserList(index).GuildInfo.YaVoto = 0
    Else
        userfile = CharPath & Left$(myGuild.Members(k), 1) & "\" & myGuild.Members(k) & ".chr"
        If PersonajeExiste(myGuild.Members(k)) Then
                Call WriteVar(userfile, "GUILD", "YaVoto", 0)
        End If
    End If
    
Next k

Exit Sub
fallo:
Call LogError("resetuservotes " & Err.number & " D: " & Err.Description)



End Sub


Public Sub DayElapsed()

On Error GoTo fallo
Exit Sub
Dim t%
Dim MemberIndex As Integer
Dim userfile As String

For t% = 1 To Guilds.Count
    If Guilds(t%).DaysSinceLastElection < Guilds(t%).ElectionPeriod Then
        Guilds(t%).DaysSinceLastElection = Guilds(t%).DaysSinceLastElection + 1
    Else
       If Guilds(t%).Elections = False Then
            Guilds(t%).ResetVotes
            Call ResetUserVotes(Guilds(t%))
            Guilds(t%).Elections = True
            
            MemberIndex = DameGuildMemberIndex(Guilds(t%).GuildName)
            
            If MemberIndex <> 0 Then
                Call SendData(ToGuildMembers, MemberIndex, 0, "||Hoy es la votacion para elegir un nuevo lider para el clan!!." & "´" & FontTypeNames.FONTTYPE_GUILD)
                Call SendData(ToGuildMembers, MemberIndex, 0, "||La eleccion durara 24 horas, se puede votar a cualquier miembro del clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
                Call SendData(ToGuildMembers, MemberIndex, 0, "||Para votar escribe /VOTO NICKNAME." & "´" & FontTypeNames.FONTTYPE_GUILD)
                Call SendData(ToGuildMembers, MemberIndex, 0, "||Solo se computara un voto por miembro." & "´" & FontTypeNames.FONTTYPE_GUILD)
            End If
        Else
            If Guilds(t%).Members.Count > 1 Then
                    'compute elections results
                    Dim leader$, newleaderindex As Integer, oldleaderindex As Integer
                    leader$ = Guilds(t%).NuevoLider
                    Guilds(t%).Elections = False
                    MemberIndex = DameGuildMemberIndex(Guilds(t%).GuildName)
                    newleaderindex = DameUserIndexConNombre(leader$)
                    oldleaderindex = DameUserIndexConNombre(Guilds(t%).leader)
                    If UCase$(leader$) <> UCase$(Guilds(t%).leader) Then
                        If oldleaderindex <> 0 Then
                            UserList(oldleaderindex).GuildInfo.EsGuildLeader = 0
                        Else
                            userfile = CharPath & Left$(Guilds(t%).leader, 1) & "\" & Guilds(t%).leader & ".chr"
                            If PersonajeExiste(Guilds(t%).leader) Then
                                    Call WriteVar(userfile, "GUILD", "EsGuildLeader", 0)
                            End If
                        End If
                        
                        If newleaderindex <> 0 Then
                            UserList(newleaderindex).GuildInfo.EsGuildLeader = 1
                            Call AddtoVar(UserList(newleaderindex).GuildInfo.VecesFueGuildLeader, 1, 10000)
                        Else
                            userfile = CharPath & Left$(leader$, 1) & "\" & leader$ & ".chr"
                            If PersonajeExiste(leader$) Then
                                    Call WriteVar(userfile, "GUILD", "EsGuildLeader", 1)
                            End If
                        End If
                        
                        Guilds(t%).leader = leader$
                    End If
                    
                    If MemberIndex <> 0 Then
                            Call SendData(ToGuildMembers, MemberIndex, 0, "||La elecciones han finalizado!!." & "´" & FontTypeNames.FONTTYPE_GUILD)
                            Call SendData(ToGuildMembers, MemberIndex, 0, "||El nuevo lider es " & leader$ & "´" & FontTypeNames.FONTTYPE_GUILD)
                    End If
                    
                    If newleaderindex <> 0 Then
                        Call SendData(ToIndex, newleaderindex, 0, "||¡¡¡Has ganado las elecciones, felicitaciones!!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
                        Call GiveGuildPoints(400, newleaderindex)
                    End If
                    Guilds(t%).DaysSinceLastElection = 0
            End If
        End If
    End If
    
Next t%

Exit Sub

fallo:
    Call LogError(Err.Description & " error en DayElapsed.")

End Sub

Public Sub GiveGuildPoints(ByVal Pts As Integer, ByVal UserIndex As Integer, Optional ByVal SendNotice As Boolean = True)
On Error GoTo fallo
If SendNotice Then _
   Call SendData(ToIndex, UserIndex, 0, "||¡¡¡Has recibido " & Pts & " guildpoints!!!" & "´" & FontTypeNames.FONTTYPE_GUILD)

Call AddtoVar(UserList(UserIndex).GuildInfo.GuildPoints, Pts, 9000000)
Exit Sub
fallo:
Call LogError("giveguildpoints " & Err.number & " D: " & Err.Description)


End Sub

Public Sub DropGuildPoints(ByVal Pts As Integer, ByVal UserIndex As Integer, Optional ByVal SendNotice As Boolean = True)
On Error GoTo fallo
UserList(UserIndex).GuildInfo.GuildPoints = UserList(UserIndex).GuildInfo.GuildPoints - Pts

'If UserList(UserIndex).GuildInfo.GuildPoints < (-5000) Then
'
'End If
Exit Sub
fallo:
Call LogError("dropguildpoints " & Err.number & " D: " & Err.Description)


End Sub


Public Sub AcceptPeaceOffer(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(rdata)

If oGuild Is Nothing Then Exit Sub

If Not oGuild.IsEnemy(UserList(UserIndex).GuildInfo.GuildName) Then
    Call SendData(ToIndex, UserIndex, 0, "||No estas en guerra con el clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Sub
End If

Call oGuild.RemoveEnemy(UserList(UserIndex).GuildInfo.GuildName)

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Call oGuild.RemoveEnemy(rdata)
Call oGuild.RemoveProposition(rdata)

Dim MemberIndex As Integer

MemberIndex = DameUserIndexConNombre(rdata)

If MemberIndex <> 0 Then _
    Call SendData(ToGuildMembers, MemberIndex, 0, "||El clan firmó la paz con " & UserList(UserIndex).GuildInfo.GuildName & "´" & FontTypeNames.FONTTYPE_GUILD)
    
Call SendData(ToGuildMembers, UserIndex, 0, "||El clan firmó la paz con " & rdata & "´" & FontTypeNames.FONTTYPE_GUILD)

Exit Sub
fallo:
Call LogError("acceptpeaceoffer " & Err.number & " D: " & Err.Description)




End Sub


Public Sub SendPeaceRequest(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim Soli As cSolicitud

Set Soli = oGuild.GetPeaceRequest(rdata)

If Soli Is Nothing Then Exit Sub

Call SendData2(ToIndex, UserIndex, 0, 60, Soli.Desc)
Exit Sub
fallo:
Call LogError("sendpeacerequest " & Err.number & " D: " & Err.Description)


End Sub


Public Sub RecievePeaceOffer(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim H$

H$ = UCase$(ReadField(1, rdata, 44))

If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(H$) Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(H$)

If oGuild Is Nothing Then Exit Sub

If Not oGuild.IsEnemy(UserList(UserIndex).GuildInfo.GuildName) Then
    Call SendData(ToIndex, UserIndex, 0, "||No estas en guerra con el clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Sub
End If

If oGuild.IsAllie(UserList(UserIndex).GuildInfo.GuildName) Then
    Call SendData(ToIndex, UserIndex, 0, "||Ya estas en paz con el clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Sub
End If

Dim peaceoffer As New cSolicitud

peaceoffer.Desc = ReadField(2, rdata, 44)
peaceoffer.UserName = UserList(UserIndex).GuildInfo.GuildName

If Not oGuild.IncludesPeaceOffer(peaceoffer.UserName) Then
    Call oGuild.PeacePropositions.Add(peaceoffer)
    Call SendData(ToIndex, UserIndex, 0, "||La propuesta de paz ha sido entregada." & "´" & FontTypeNames.FONTTYPE_GUILD)
Else
    Call SendData(ToIndex, UserIndex, 0, "||Ya has enviado una propuesta de paz." & "´" & FontTypeNames.FONTTYPE_GUILD)
End If

Exit Sub
fallo:
Call LogError("recibepeaceoffer " & Err.number & " D: " & Err.Description)

End Sub


Public Sub SendPeacePropositions(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim l%, k$

If oGuild.PeacePropositions.Count = 0 Then Exit Sub

k$ = oGuild.PeacePropositions.Count & ","

For l% = 1 To oGuild.PeacePropositions.Count
    k$ = k$ & oGuild.PeacePropositions(l%).UserName & ","
Next l%

Call SendData2(ToIndex, UserIndex, 0, 61, k$)
Exit Sub
fallo:
Call LogError("sendpeacepropositions " & Err.number & " D: " & Err.Description)

End Sub



Public Sub EacharMember(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim MemberIndex As Integer

MemberIndex = DameUserIndexConNombre(rdata)
'pluto:2-3-04
If MemberIndex = 0 Then
Dim userfile As String
userfile = CharPath & Left$(rdata, 1) & "\" & rdata & ".chr"
If Not PersonajeExiste(rdata) Then
    Call SendData(ToIndex, UserIndex, 0, "||No existe ese PJ." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Sub
End If
If val(GetVar(userfile, "GUILD", "EsGuildLeader")) = 1 Then
Call SendData(ToIndex, UserIndex, 0, "||Un lider no puede abandonar su clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
Exit Sub
End If
If GetVar(userfile, "GUILD", "GuildName") = UserList(UserIndex).GuildInfo.GuildName Then
Dim o As Integer
o = val(GetVar(userfile, "GUILD", "Echadas"))
o = o + 1
Call WriteVar(userfile, "GUILD", "Echadas", val(o))
Call WriteVar(userfile, "GUILD", "GuildPts", 0)
Call WriteVar(userfile, "GUILD", "GuildName", "")
'pluto:2.9.0
Call WriteVar(userfile, "STATS", "Pclan", 0)

Call SendData(ToGuildMembers, UserIndex, 0, "||" & rdata & " fue expulsado del clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
Call oGuild.RemoveMember(rdata)
Call LogClanMov("Expulsado " & rdata & " de " & UserList(UserIndex).GuildInfo.GuildName)
Exit Sub
Else
Call oGuild.RemoveMember(rdata)
Call LogClanMov("Expulsado " & rdata & " de " & UserList(UserIndex).GuildInfo.GuildName)
Exit Sub
End If
End If
'--fin pluto---

If UserList(MemberIndex).GuildInfo.EsGuildLeader = 1 Then
Call SendData(ToIndex, MemberIndex, 0, "||Un lider no puede abandonar su clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
Exit Sub
End If

If MemberIndex <> 0 Then 'esta online
    If UserList(MemberIndex).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName Then
    Call SendData(ToIndex, MemberIndex, 0, "||Has sido expulsado del clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Call AddtoVar(UserList(MemberIndex).GuildInfo.Echadas, 1, 1000)
    UserList(MemberIndex).GuildInfo.GuildPoints = 0
    UserList(MemberIndex).GuildInfo.GuildName = ""
    'pluto:2.9.0
    UserList(MemberIndex).Stats.PClan = 0
    
    Call SendData(ToGuildMembers, UserIndex, 0, "||" & rdata & " fue expulsado del clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
    End If

Else
    Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta ONLINE." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Sub
End If

Call oGuild.RemoveMember(UserList(MemberIndex).Name)
Call LogClanMov("Expulsado " & rdata & " de " & UserList(UserIndex).GuildInfo.GuildName)
Exit Sub
fallo:
Call LogError("echarmember " & Err.number & " D: " & Err.Description)

End Sub

Public Sub DenyRequest(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim Soli As cSolicitud

Set Soli = oGuild.GetSolicitud(rdata)

If Soli Is Nothing Then Exit Sub

Dim MemberIndex As Integer

MemberIndex = DameUserIndexConNombre(Soli.UserName)

If MemberIndex <> 0 Then 'esta online
    Call SendData(ToIndex, MemberIndex, 0, "||Tu solicitud ha sido rechazada." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Call AddtoVar(UserList(MemberIndex).GuildInfo.SolicitudesRechazadas, 1, 10000)
End If

Call oGuild.RemoveSolicitud(Soli.UserName)
Exit Sub
fallo:
Call LogError("denyrequest " & Err.number & " D: " & Err.Description)

End Sub


Public Sub AcceptClanMember(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
Dim oGuild As cGuild
Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
If oGuild Is Nothing Then Exit Sub
Dim Soli As cSolicitud
Set Soli = oGuild.GetSolicitud(rdata)
If Soli Is Nothing Then Exit Sub
Dim MemberIndex As Integer
MemberIndex = DameUserIndexConNombre(Soli.UserName)

'pluto:2.15

'Dim UserLider As String
'Dim LevelFounder As Byte
'UserLider = CharPath & Left$(oGuild.Founder, 1) & "\" & oGuild.Founder & ".chr"
'LevelFounder = val(GetVar(UserLider, "STATS", "ELV"))

'pluto:6.0A
Dim TopePjs As Byte

Select Case oGuild.Nivel
Case 1
TopePjs = 12
Case 2
TopePjs = 14
Case 3
TopePjs = 15
Case 4
TopePjs = 16
Case 5
TopePjs = 18
'pluto:6.9
Case 6
TopePjs = 20
End Select
If oGuild.Members.Count > TopePjs Then
    Call SendData(ToIndex, UserIndex, 0, "||El clan es demasiado numeroso." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Sub
End If
'-----------------------------

If MemberIndex <> 0 Then 'esta online
    If UserList(MemberIndex).GuildInfo.GuildName <> "" Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes aceptar esa solicitud, el pesonaje es lider de otro clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
'pluto:2.17
    If EsNewbie(MemberIndex) Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes aceptar Newbies." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
    'pluto:2.17
    'If UserList(MemberIndex).GuildInfo.ClanesParticipo > 10 Then
     
     If (10 + Int(UserList(MemberIndex).Mision.numero / 20) - UserList(MemberIndex).GuildInfo.ClanesParticipo) < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes aceptarle porque no tiene solicitudes disponibles." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Call SendData(ToIndex, MemberIndex, 0, "||Estuviste en más de 10 clanes, debes hacer NpcQuest para ganar más solicitudes." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
    
    UserList(MemberIndex).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName
    Call AddtoVar(UserList(MemberIndex).GuildInfo.ClanesParticipo, 1, 1000)
    Call SendData(ToIndex, MemberIndex, 0, "||Felicitaciones, tu solicitud ha sido aceptada." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Call SendData(ToIndex, MemberIndex, 0, "||Ahora sos un miembro activo del clan " & UserList(UserIndex).GuildInfo.GuildName & "´" & FontTypeNames.FONTTYPE_GUILD)
    Call SendData(ToIndex, MemberIndex, 0, "||Has participado en " & UserList(MemberIndex).GuildInfo.ClanesParticipo & ". Recuerda que el máximo es de 10 participaciones en Clanes." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Call GiveGuildPoints(25, MemberIndex)
Else
    'pluto:2-3-04
    Dim userfile As String
    userfile = CharPath & Left$(rdata, 1) & "\" & rdata & ".chr"
    If PersonajeExiste(rdata) = False Then
        Call SendData(ToIndex, UserIndex, 0, "||No existe ese PJ." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
    If GetVar(userfile, "GUILD", "GuildName") <> "" Then
        Call SendData(ToIndex, UserIndex, 0, "||Ya pertenece a otro Clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
'pluto:2.17
    'If val(GetVar(userfile, "GUILD", "ClanesParticipo")) > 10 Then
         If (10 + Int(val(GetVar(userfile, "QUEST", "Numero")) / 20)) - val(GetVar(userfile, "GUILD", "ClanesParticipo")) < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes aceptarle porque no tiene solicitudes." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
    If GetVar(userfile, "GUILD", "GuildName") = "" Then
        Dim o As Integer
        o = val(GetVar(userfile, "GUILD", "ClanesParticipo"))
        o = o + 1
        Call WriteVar(userfile, "GUILD", "ClanesParticipo", val(o))
        Call WriteVar(userfile, "GUILD", "GuildPts", 25)
        Call WriteVar(userfile, "GUILD", "GuildName", UserList(UserIndex).GuildInfo.GuildName)
        Call SendData(ToGuildMembers, UserIndex, 0, "||" & rdata & " fue aceptado por el clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Call LogClanMov("Aceptado " & rdata & " en " & UserList(UserIndex).GuildInfo.GuildName)
        Call oGuild.Members.Add(Soli.UserName)
        Call oGuild.RemoveSolicitud(Soli.UserName)
        Exit Sub
    End If
End If
'---------------fin pluto---------------

Call oGuild.Members.Add(Soli.UserName)
Call oGuild.RemoveSolicitud(Soli.UserName)
Call SendData(ToGuildMembers, UserIndex, 0, "TW" & SND_ACEPTADOCLAN)
Call SendData(ToGuildMembers, UserIndex, 0, "||" & rdata & " ha sido aceptado en el clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
Call LogClanMov("Aceptado " & rdata & " en " & UserList(UserIndex).GuildInfo.GuildName)
Exit Sub
fallo:
Call LogError("acceptclamember " & Err.number & " D: " & Err.Description)

End Sub


Public Sub SendPeticion(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
    
Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

  
Dim Soli As cSolicitud

Set Soli = oGuild.GetSolicitud(rdata)

If Soli Is Nothing Then Exit Sub

Call SendData2(ToIndex, UserIndex, 0, 69, Soli.Desc)

Exit Sub
fallo:
Call LogError("sendpeticion " & Err.number & " D: " & Err.Description)

End Sub


Public Sub SolicitudIngresoClan(ByVal UserIndex As Integer, ByVal Data As String)
On Error GoTo fallo
If EsNewbie(UserIndex) Then
   Call SendData(ToIndex, UserIndex, 0, "||Los newbies no pueden conformar clanes." & "´" & FontTypeNames.FONTTYPE_GUILD)
   Exit Sub
End If

Dim MiSol As New cSolicitud

MiSol.Desc = ReadField(2, Data, 44)
MiSol.UserName = UserList(UserIndex).Name

Dim clan$

clan$ = ReadField(1, Data, 44)


Dim oGuild As cGuild

Set oGuild = FetchGuild(clan$)

If oGuild Is Nothing Then Exit Sub

If oGuild.IsMember(UserList(UserIndex).Name) Then Exit Sub

If Not oGuild.SolicitudesIncludes(MiSol.UserName) Then
        Call AddtoVar(UserList(UserIndex).GuildInfo.Solicitudes, 1, 1000)
        
        Call oGuild.TestSolicitudBound
        Call oGuild.Solicitudes.Add(MiSol)
        
        Call SendData(ToIndex, UserIndex, 0, "||La solicitud fue recibida por el lider del clan, ahora debes esperar la respuesta." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
Else
        Call SendData(ToIndex, UserIndex, 0, "||Tu solicitud ya fue recibida por el lider del clan, ahora debes esperar la respuesta." & "´" & FontTypeNames.FONTTYPE_GUILD)
End If

Exit Sub
fallo:
Call LogError("solicitudingresoclan " & Err.number & " D: " & Err.Description)

End Sub


Public Sub SendCharInfo(ByVal UserName As String, ByVal UserIndex As Integer)
On Error GoTo fallo
'¿Existe el personaje?

If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub


Dim userfile As String
userfile = CharPath & Left$(UserName, 1) & "\" & UserName & ".chr"

If Not PersonajeExiste(UserName) Then Exit Sub

Dim MiUser As User

MiUser.Name = UserName
MiUser.raza = GetVar(userfile, "INIT", "Raza")
MiUser.clase = GetVar(userfile, "INIT", "Clase")
MiUser.Genero = GetVar(userfile, "INIT", "Genero")
MiUser.Stats.ELV = val(GetVar(userfile, "STATS", "ELV"))
MiUser.Stats.GLD = val(GetVar(userfile, "STATS", "GLD"))
MiUser.Stats.Banco = val(GetVar(userfile, "STATS", "BANCO"))
MiUser.Reputacion.Promedio = val(GetVar(userfile, "REP", "Promedio"))
'pluto:6.9
MiUser.Remort = val(GetVar(userfile, "STATS", "REMORT"))
Dim H$
H$ = UserName & ","
H$ = H$ & MiUser.raza & ","
H$ = H$ & MiUser.clase & ","
H$ = H$ & MiUser.Genero & ","
H$ = H$ & MiUser.Stats.ELV & ","
H$ = H$ & MiUser.Stats.GLD & ","
H$ = H$ & MiUser.Stats.Banco & ","
H$ = H$ & MiUser.Reputacion.Promedio & ","


MiUser.GuildInfo.FundoClan = val(GetVar(userfile, "Guild", "FundoClan"))
MiUser.GuildInfo.EsGuildLeader = val(GetVar(userfile, "Guild", "EsGuildLeader"))
MiUser.GuildInfo.Echadas = val(GetVar(userfile, "Guild", "Echadas"))
MiUser.GuildInfo.Solicitudes = val(GetVar(userfile, "Guild", "Solicitudes"))
MiUser.GuildInfo.SolicitudesRechazadas = val(GetVar(userfile, "Guild", "SolicitudesRechazadas"))
MiUser.GuildInfo.VecesFueGuildLeader = val(GetVar(userfile, "Guild", "VecesFueGuildLeader"))
'MiUser.GuildInfo.YaVoto = val(GetVar(UserFile, "Guild", "YaVoto"))
MiUser.GuildInfo.ClanesParticipo = val(GetVar(userfile, "Guild", "ClanesParticipo"))

H$ = H$ & MiUser.GuildInfo.FundoClan & ","
H$ = H$ & MiUser.GuildInfo.EsGuildLeader & ","
H$ = H$ & MiUser.GuildInfo.Echadas & ","
H$ = H$ & MiUser.GuildInfo.Solicitudes & ","
H$ = H$ & MiUser.GuildInfo.SolicitudesRechazadas & ","
H$ = H$ & MiUser.GuildInfo.VecesFueGuildLeader & ","
H$ = H$ & MiUser.GuildInfo.ClanesParticipo & ","


MiUser.GuildInfo.ClanFundado = GetVar(userfile, "Guild", "ClanFundado")
MiUser.GuildInfo.GuildName = GetVar(userfile, "Guild", "GuildName")


H$ = H$ & MiUser.GuildInfo.ClanFundado & ","
H$ = H$ & MiUser.GuildInfo.GuildName & ","


MiUser.Faccion.ArmadaReal = val(GetVar(userfile, "FACCIONES", "EjercitoReal"))
MiUser.Faccion.FuerzasCaos = val(GetVar(userfile, "FACCIONES", "EjercitoCaos"))
MiUser.Faccion.CiudadanosMatados = val(GetVar(userfile, "FACCIONES", "CiudMatados"))
MiUser.Faccion.CriminalesMatados = val(GetVar(userfile, "FACCIONES", "CrimMatados"))

H$ = H$ & MiUser.Faccion.ArmadaReal & ","
H$ = H$ & MiUser.Faccion.FuerzasCaos & ","
H$ = H$ & MiUser.Faccion.CiudadanosMatados & ","
'pluto:2.4
H$ = H$ & MiUser.Faccion.CriminalesMatados & ","
MiUser.Stats.PClan = val(GetVar(userfile, "STATS", "PCLAN"))
MiUser.GuildInfo.GuildPoints = val(GetVar(userfile, "Guild", "GuildPts"))
H$ = H$ & MiUser.Stats.PClan & ","
H$ = H$ & MiUser.GuildInfo.GuildPoints & ","
'Delzak) corregido por pluto (ya existe remort)
'MiUser.Stats.Remort = val(GetVar(userfile, "STATS", "REMORT"))
H$ = H$ & MiUser.Remort & ","

Call SendData2(ToIndex, UserIndex, 0, 62, H$)

Exit Sub
fallo:
Call LogError("sendcharinfo " & Err.number & " D: " & Err.Description)

End Sub



Public Sub UpdateGuildNews(ByVal rdata As String, ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

oGuild.GuildNews = rdata
Exit Sub
fallo:
Call LogError("updateguildnews " & Err.number & " D: " & Err.Description)

End Sub


Public Sub UpdateCodexAndDesc(ByVal rdata As String, ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Call oGuild.UpdateCodexAndDesc(rdata)
Exit Sub
fallo:
Call LogError("updatecodexanddesc " & Err.number & " D: " & Err.Description)

End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim cad$, t%

'<-------Lista de guilds ---------->

cad$ = Guilds.Count & "¬"
For t% = 1 To Guilds.Count
    cad$ = cad$ & Guilds(t%).GuildName & "¬"
Next t%

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Call SendData2(ToIndex, UserIndex, 0, 63, cad$)

'<-------Lista de miembros ---------->

cad$ = oGuild.Members.Count & "¬"

For t% = 1 To oGuild.Members.Count
    cad$ = cad$ & oGuild.Members.Item(t%) & "¬"
Next t%

Call SendData2(ToIndex, UserIndex, 0, 64, cad$)

'<------- Guild News -------->

cad$ = Replace(oGuild.GuildNews, vbCrLf, "º") & "¬"

'<------- Solicitudes ------->

cad$ = cad$ & oGuild.Solicitudes.Count & "¬"

For t% = 1 To oGuild.Solicitudes.Count
    cad$ = cad$ & oGuild.Solicitudes.Item(t%).UserName & "¬"
Next t%
Call SendData2(ToIndex, UserIndex, 0, 65, cad$)
Exit Sub
fallo:
Call LogError("sendguildleaderinfo " & Err.number & " D: " & Err.Description)

End Sub

Public Sub SetNewURL(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

oGuild.URL = rdata

Call SendData(ToIndex, UserIndex, 0, "||La direccion de la web ha sido actualizada" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
fallo:
Call LogError("setnewurl " & Err.number & " D: " & Err.Description)

End Sub

Public Sub SetNewEmblema(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

oGuild.Emblema = rdata

Call SendData(ToIndex, UserIndex, 0, "||La direccion de la web ha sido actualizada" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
fallo:
Call LogError("setnewemblema " & Err.number & " D: " & Err.Description)

End Sub

Public Sub DeclareAllie(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(rdata) Then Exit Sub


Dim LeaderGuild As cGuild, enemyGuild As cGuild

Set LeaderGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If LeaderGuild Is Nothing Then Exit Sub

Set enemyGuild = FetchGuild(rdata)

If enemyGuild Is Nothing Then Exit Sub

If LeaderGuild.IsEnemy(enemyGuild.GuildName) Then
        Call SendData(ToIndex, UserIndex, 0, "||Estas en guerra con éste clan, antes debes firmar la paz." & "´" & FontTypeNames.FONTTYPE_GUILD)
Else
   If Not LeaderGuild.IsAllie(enemyGuild.GuildName) Then
        Call LeaderGuild.AlliedGuilds.Add(enemyGuild.GuildName)
        Call enemyGuild.AlliedGuilds.Add(LeaderGuild.GuildName)
        
        Call SendData(ToGuildMembers, UserIndex, 0, "||Tu clan ha firmado una alianza con " & enemyGuild.GuildName & "´" & FontTypeNames.FONTTYPE_GUILD)
        Call SendData(ToGuildMembers, UserIndex, 0, "TW" & SND_DECLAREWAR)
        
        Dim index As Integer
        index = DameGuildMemberIndex(enemyGuild.GuildName)
        If index <> 0 Then
            Call SendData(ToGuildMembers, index, 0, "||" & LeaderGuild.GuildName & " firmo una alianza con tu clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
            Call SendData(ToGuildMembers, index, 0, "TW" & SND_DECLAREWAR)
        End If
   Else
        Call SendData(ToIndex, UserIndex, 0, "||Ya estas aliado con éste clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
   End If
End If

    
Exit Sub
fallo:
Call LogError("declareallie " & Err.number & " D: " & Err.Description)


End Sub


Public Sub DeclareWar(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then Exit Sub

If UCase$(UserList(UserIndex).GuildInfo.GuildName) = UCase$(rdata) Then Exit Sub


Dim LeaderGuild As cGuild, enemyGuild As cGuild

Set LeaderGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If LeaderGuild Is Nothing Then Exit Sub

Set enemyGuild = FetchGuild(rdata)

If enemyGuild Is Nothing Then Exit Sub

If Not LeaderGuild.IsEnemy(enemyGuild.GuildName) Then
        
        Call LeaderGuild.RemoveAllie(enemyGuild.GuildName)
        Call enemyGuild.RemoveAllie(LeaderGuild.GuildName)
        
        Call LeaderGuild.EnemyGuilds.Add(enemyGuild.GuildName)
        Call enemyGuild.EnemyGuilds.Add(LeaderGuild.GuildName)
        
        
        Call SendData(ToGuildMembers, UserIndex, 0, "||Tu clan le ha declarado la guerra al clan " & enemyGuild.GuildName & "´" & FontTypeNames.FONTTYPE_GUILD)
        Call SendData(ToGuildMembers, UserIndex, 0, "TW" & SND_DECLAREWAR)
        
        Dim index As Integer
        index = DameGuildMemberIndex(enemyGuild.GuildName)
        If index <> 0 Then
            Call SendData(ToGuildMembers, index, 0, "||" & LeaderGuild.GuildName & " ha declarado la guerra a tu clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
            Call SendData(ToGuildMembers, index, 0, "TW" & SND_DECLAREWAR)
        End If
Else
   Call SendData(ToIndex, UserIndex, 0, "||Tu clan ya esta en guerra con " & enemyGuild.GuildName & "´" & FontTypeNames.FONTTYPE_GUILD)
End If

Exit Sub
fallo:
Call LogError("declarewar " & Err.number & " D: " & Err.Description)

End Sub

Public Function DameGuildMemberIndex(ByVal GuildName As String) As Integer
On Error GoTo fallo
Dim loopc As Integer
  
loopc = 1
  
GuildName = UCase$(GuildName)
  
Do Until UCase$(UserList(loopc).GuildInfo.GuildName) = GuildName

    loopc = loopc + 1
    
    If loopc > MaxUsers Then
        DameGuildMemberIndex = 0
        Exit Function
    End If
    
Loop
  
DameGuildMemberIndex = loopc


Exit Function
fallo:
Call LogError("dameguildmemberindex " & Err.number & " D: " & Err.Description)

End Function


Public Sub SendGuildNews(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).GuildInfo.GuildName = "" Then Exit Sub


Dim oGuild As cGuild

Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

If oGuild Is Nothing Then Exit Sub

Dim k$

k$ = oGuild.GuildNews & "¬"

Dim t%

k$ = k$ & oGuild.EnemyGuilds.Count & "¬"

For t% = 1 To oGuild.EnemyGuilds.Count

    k$ = k$ & oGuild.EnemyGuilds(t%) & "¬"
    
Next t%

k$ = k$ & oGuild.AlliedGuilds.Count & "¬"

For t% = 1 To oGuild.AlliedGuilds.Count

    k$ = k$ & oGuild.AlliedGuilds(t%) & "¬"
    
Next t%



Call SendData2(ToIndex, UserIndex, 0, 59, k$)

If oGuild.Elections Then
    Call SendData(ToIndex, UserIndex, 0, "||Hoy es la votacion para elegir un nuevo lider para el clan!!." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Call SendData(ToIndex, UserIndex, 0, "||La eleccion durara 24 horas, se puede votar a cualquier miembro del clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Call SendData(ToIndex, UserIndex, 0, "||Para votar escribe /VOTO NICKNAME." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Call SendData(ToIndex, UserIndex, 0, "||Solo se computara un voto por miembro." & "´" & FontTypeNames.FONTTYPE_GUILD)
End If

Exit Sub
fallo:
Call LogError("sendguildnews " & Err.number & " D: " & Err.Description)

End Sub

Public Sub SendGuildsList(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim cad$, t%

cad$ = "GL" & Guilds.Count & ","

For t% = 1 To Guilds.Count
    cad$ = cad$ & Guilds(t%).GuildName & ","
Next t%

Call SendData(ToIndex, UserIndex, 0, cad$)
Exit Sub
fallo:
Call LogError("sendguildlist " & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.4
Public Sub SendGuildsPuntos(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim cad As String
Dim t As Byte

cad = "GX" & Guilds.Count & ","

For t = 1 To Guilds.Count
    cad = cad & Guilds(t).GuildName & ","
Next t
For t = 1 To Guilds.Count
    cad = cad & Guilds(t).Reputation & ","
Next t
'pluto:6.0A--------------------------------
For t = 1 To Guilds.Count
    cad = cad & Guilds(t).Nivel & ","
Next t
'-------------------------------------------
Call SendData(ToIndex, UserIndex, 0, cad)

Exit Sub
fallo:
Call LogError("sendguildspuntos " & Err.number & " D: " & Err.Description)

End Sub

Public Function FetchGuild(ByVal GuildName As String) As Object
On Error GoTo fallo
Dim k As Integer
For k = 1 To Guilds.Count


    If UCase$(Guilds.Item(k).GuildName) = UCase$(GuildName) Then
            Set FetchGuild = Guilds.Item(k)
            Exit Function
    End If
Next k
Exit Function
fallo:
Call LogError("fetchguild " & Err.number & " D: " & Err.Description)

End Function

Public Sub LoadGuildsDB()
On Error GoTo fallo
Dim file As String
Dim Cant As Integer

file = App.Path & "\Guilds\" & "GuildsInfo.inf"

If Not FileExist(file, vbNormal) Then Exit Sub

Cant = val(GetVar(file, "INIT", "NroGuilds"))
'pluto:6.9---------
ReDim PuntClan(1 To Cant) As Integer
ReDim NomClan(1 To Cant) As String
'--------------------
Dim NewGuild As cGuild
Dim k%

For k% = 1 To Cant
    Set NewGuild = New cGuild
    Call NewGuild.InitializeGuildFromDisk(k%)
    Call Guilds.Add(NewGuild)
  
Next k%

'pluto:6.9
  'ordenamos puntostorneos
Dim E As Integer
Dim i As Integer
Dim dniaux As Integer
Dim nomaux1 As String

For E = 1 To Guilds.Count
NomClan(E) = Guilds(E).GuildName
PuntClan(E) = Guilds(E).PuntosTorneos
Next
            
            
            For E = 1 To Guilds.Count
                For i = 1 To Guilds.Count
    
           If PuntClan(i) < PuntClan(E) Then
              nomaux1 = NomClan(i)
              NomClan(i) = NomClan(E)
              NomClan(E) = nomaux1
            
           
              dniaux = PuntClan(i)
              PuntClan(i) = PuntClan(E)
              PuntClan(E) = dniaux
           End If
   
                Next i
            Next E
'--------------------


Exit Sub
fallo:
Call LogError("loadguildsdb " & Err.number & " D: " & Err.Description)

End Sub

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String)
On Error GoTo fallo

Dim oGuild As cGuild

If Guilds.Count = 0 Then Exit Sub

Set oGuild = FetchGuild(GuildName)

If oGuild Is Nothing Then Exit Sub

Dim cad$

cad$ = cad$ & oGuild.GuildName
cad$ = cad$ & "¬" & oGuild.Founder
cad$ = cad$ & "¬" & oGuild.FundationDate
cad$ = cad$ & "¬" & oGuild.leader
cad$ = cad$ & "¬" & oGuild.URL
cad$ = cad$ & "¬" & oGuild.Members.Count
cad$ = cad$ & "¬" & oGuild.DaysToNextElection
cad$ = cad$ & "¬" & oGuild.Nivel
cad$ = cad$ & "¬" & oGuild.EnemyGuilds.Count
cad$ = cad$ & "¬" & oGuild.AlliedGuilds.Count

Dim codex$

codex$ = oGuild.CodexLenght()

Dim k%

For k% = 0 To oGuild.CodexLenght()
    codex$ = codex$ & "¬" & oGuild.GetCodex(k%)
Next k%


cad$ = cad$ & "¬" & codex$ & oGuild.Description


Call SendData2(ToIndex, UserIndex, 0, 66, cad$)

Exit Sub
fallo:
Call LogError("sendguilddetails " & Err.number & " D: " & Err.Description)


End Sub


Public Function CanCreateGuild(ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
If UserList(UserIndex).Stats.ELV < 30 Then
    CanCreateGuild = False
    Call SendData(ToIndex, UserIndex, 0, "||Para fundar un clan debes de ser nivel 30 o superior" & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Function
End If
 
 If UserList(UserIndex).Stats.UserAtributosBackUP(Carisma) < 19 Then
    CanCreateGuild = False
    Call SendData(ToIndex, UserIndex, 0, "||Para fundar un clan debes tener carisma 19 o superior" & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Function
End If

If UserList(UserIndex).Stats.UserSkills(Liderazgo) < 200 Then
    CanCreateGuild = False
    Call SendData(ToIndex, UserIndex, 0, "||Para fundar un clan necesitas al menos 200 pts en liderazgo" & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Function
End If
'PLUTO
If UserList(UserIndex).Stats.GLD < 1000000 Then
    CanCreateGuild = False
    Call SendData(ToIndex, UserIndex, 0, "||Para fundar un clan necesitas 1.000.000 Oros." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Function
End If
'PLUTO:6.0A
If UserList(UserIndex).GuildInfo.GuildName <> "" Then
    CanCreateGuild = False
    Call SendData(ToIndex, UserIndex, 0, "||Para fundar un clan necesitas salir antes del clan actual." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Function
End If

CanCreateGuild = True
Exit Function
fallo:
Call LogError("cancreateguild " & Err.number & " D: " & Err.Description)

End Function

Public Function ExisteGuild(ByVal Name As String) As Boolean
On Error GoTo fallo
Dim k As Integer
Name = UCase$(Name)

For k = 1 To Guilds.Count
    If UCase$(Guilds(k).GuildName) = Name Then
            ExisteGuild = True
            Exit Function
    End If
Next k
Exit Function
fallo:
Call LogError("existeguild " & Err.number & " D: " & Err.Description)

End Function

Public Function CreateGuild(ByVal Name As String, ByVal Rep As Long, ByVal index As Integer, ByVal GuildInfo As String) As Boolean
On Error GoTo fallo
If Not CanCreateGuild(index) Then
    CreateGuild = False
    Exit Function
End If

Dim miClan As New cGuild

If Not miClan.Initialize(GuildInfo, Name, Rep) Then
    CreateGuild = False
    Call SendData(ToIndex, index, 0, "||Los datos del clan son invalidos, asegurate que no contiene caracteres invalidos." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Function
End If

If ExisteGuild(miClan.GuildName) Then
    CreateGuild = False
    Call SendData(ToIndex, index, 0, "||Ya exíste un clan con ese nombre." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Function
End If

Call miClan.Members.Add(UCase$(UserList(index).Name))

Call Guilds.Add(miClan, miClan.GuildName)

UserList(index).GuildInfo.FundoClan = 1
UserList(index).GuildInfo.EsGuildLeader = 1

Call AddtoVar(UserList(index).GuildInfo.VecesFueGuildLeader, 1, 10000)
Call AddtoVar(UserList(index).GuildInfo.ClanesParticipo, 1, 10000)

UserList(index).GuildInfo.ClanFundado = miClan.GuildName
UserList(index).GuildInfo.GuildName = UserList(index).GuildInfo.ClanFundado


Call GiveGuildPoints(5000, index)

Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
Call SendData(ToAll, 0, 0, "||¡¡¡" & UserList(index).Name & " fundo el clan '" & UserList(index).GuildInfo.GuildName & "'!!!" & "´" & FontTypeNames.FONTTYPE_GUILD)

CreateGuild = True
UserList(index).Stats.GLD = UserList(index).Stats.GLD - 1000000


Exit Function
fallo:
Call LogError("createguild " & Err.number & " D: " & Err.Description)

End Function


Public Sub SaveGuildsDB()
On Error GoTo fallo
Dim j As Integer
Dim file As String

file = App.Path & "\Guilds\" & "GuildsInfo.inf"

If FileExist(file, vbNormal) Then Kill file

Call WriteVar(file, "INIT", "NroGuilds", str(Guilds.Count))
For j = 1 To Guilds.Count
    
    Call Guilds(j).SaveGuild(file, j)
    
Next j
Exit Sub
fallo:
Call LogError("saveguildsdb " & Err.number & " D: " & Err.Description)

End Sub
