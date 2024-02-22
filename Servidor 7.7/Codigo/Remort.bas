Attribute VB_Name = "Remort"
Public Sub DoRemort(raza As String, UserIndex As Integer)
    On Error GoTo fallo
    Dim X As Integer
    'pluto:6.0A
    If UserList(UserIndex).flags.Navegando > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Deja de Navegar!." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If UserList(UserIndex).flags.TomoPocion = True Or UserList(UserIndex).flags.DuracionEfecto = True Then
        Call SendData(ToIndex, UserIndex, 0, "||Espera que se pase el efecto del dope." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Or UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Or UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Or UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||Desequipate todo." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    If (UserList(UserIndex).Stats.ELV < 55) Then
        Call SendData(ToIndex, UserIndex, 0, "||No eres nivel 55" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If (UserList(UserIndex).flags.Privilegios > 0) Then
        Call SendData(ToIndex, UserIndex, 0, "||Dejate de coñas, y atiende los SOS." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If (UserList(UserIndex).Remort = 1) Then
        Call SendData(ToIndex, UserIndex, 0, "||Ya has hecho remort ;)" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If (UserList(UserIndex).GuildInfo.EsGuildLeader = 1) Then
        Call SendData(ToIndex, UserIndex, 0, "||Un lider no puede abandonar su Clan" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    'pluto:2.17
    If (UserList(UserIndex).GuildInfo.GuildName <> "") Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes salir del Clan." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    'pluto:6.9
     If UserList(UserIndex).Stats.GLD > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Deja tu oro en el Banco antes de hacer remort!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
   
    
    
    Dim Name As String
    Name = UserList(UserIndex).Name
    Select Case UCase$(raza)
        Case "ELIAN-LAL"
            UserList(UserIndex).Remort = 1
            UserList(UserIndex).Remorted = "Elian-LAL"
            UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 8
            UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 1
        Case "GORK-ROR"
            UserList(UserIndex).Remort = 1
            UserList(UserIndex).Remorted = "Gork-RoR"
            UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 1
            UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 8
        Case "DRAKON"
            UserList(UserIndex).Remort = 1
            UserList(UserIndex).Remorted = "Drakon"
            UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 4
            UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 4
      Case Else
            Call SendData(ToIndex, UserIndex, 0, "||Raza desconocida, las razas posibles son: ELIAN-LAL, GORK-ROR, DRAKON" & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
    End Select
    
    If (UserList(UserIndex).Remort = 1) Then
        
       'pluto:2-3-04
        Call QuitarObjetos(882, 1, UserIndex)

        For X = 1 To NUMSKILLS
            UserList(UserIndex).Stats.UserSkills(X) = 0
        Next X
'pluto:2-3-04 -----------------------------------
        For loopc = 1 To MAXUSERHECHIZOS
            UserList(UserIndex).Stats.UserHechizos(loopc) = 0
        Next loopc
        Call LimpiarInventario(UserIndex)
        Call DarCuerpoDesnudo(UserIndex)
 '------------------------------------------------
        UserList(UserIndex).Stats.MaxHP = 5 + UserList(UserIndex).Stats.UserAtributos(Constitucion)
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        UserList(UserIndex).Stats.MaxAGU = 200
        UserList(UserIndex).Stats.MaxHam = 200
        UserList(UserIndex).Stats.MaxSta = 5 + UserList(UserIndex).Stats.UserAtributos(Agilidad)
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
        If UserList(UserIndex).clase = "Mago" Then
            UserList(UserIndex).Stats.MaxMAN = 50 + UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            UserList(UserIndex).Stats.MinMAN = 50 + UserList(UserIndex).Stats.UserAtributos(Inteligencia)
        ElseIf UserList(UserIndex).clase = "Clerigo" Or UserList(UserIndex).clase = "Druida" _
            Or UserList(UserIndex).clase = "Bardo" Or UserList(UserIndex).clase = "Asesino" Or UserList(UserIndex).clase = "Pirata" Then
                UserList(UserIndex).Stats.MaxMAN = 30
                UserList(UserIndex).Stats.MinMAN = 30
        Else
            UserList(UserIndex).Stats.MaxMAN = 0
            UserList(UserIndex).Stats.MinMAN = 0
        End If
        UserList(UserIndex).Stats.GLD = 0
        UserList(UserIndex).Stats.Puntos = 0
        UserList(UserIndex).Stats.MaxHIT = 3
        UserList(UserIndex).Stats.MinHIT = 2
        UserList(UserIndex).Stats.exp = 0
        'pluto:2.9.0
        UserList(UserIndex).Stats.PClan = 0
        UserList(UserIndex).GuildInfo.GuildPoints = 0
        'pluto:6.0
        UserList(UserIndex).flags.Minotauro = 0
'pluto:2.17
UserList(UserIndex).Stats.Elu = 900
UserList(UserIndex).Stats.LibrosUsados = 0

'UserList(UserIndex).Stats.Elu = 1200 - ((UserList(UserIndex).Stats.ELV - 45) * 40)
        
        UserList(UserIndex).Stats.ELV = 1
        UserList(UserIndex).Stats.SkillPts = 10
        Call ResetFacciones(UserIndex)
        UserList(UserIndex).Reputacion.AsesinoRep = 0
        UserList(UserIndex).Reputacion.BandidoRep = 0
        UserList(UserIndex).Reputacion.BurguesRep = 0
        UserList(UserIndex).Reputacion.LadronesRep = 0
        UserList(UserIndex).Reputacion.NobleRep = 1000
        UserList(UserIndex).Reputacion.PlebeRep = 30
        UserList(UserIndex).Reputacion.Promedio = 30 / 6

        Call ResetGuildInfo(UserIndex)
        Call ResetUserMision(UserIndex)
        
        
        
        
Select Case UCase$(UserList(UserIndex).raza)
Case "ORCO"
Call WarpUserChar(UserIndex, Pobladoorco.Map, Pobladoorco.X, Pobladoorco.Y, True)
Case "HUMANO"
Call WarpUserChar(UserIndex, Pobladohumano.Map, Pobladohumano.X, Pobladohumano.Y, True)
Case "CICLOPE"
Call WarpUserChar(UserIndex, Pobladohumano.Map, Pobladohumano.X, Pobladohumano.Y, True)
Case "ELFO"
Call WarpUserChar(UserIndex, Pobladoelfo.Map, Pobladoelfo.X, Pobladoelfo.Y, True)
Case "ELFO OSCURO"
Call WarpUserChar(UserIndex, Pobladoelfo.Map, Pobladoelfo.X, Pobladoelfo.Y, True)
Case "VAMPIRO"
Call WarpUserChar(UserIndex, Pobladovampiro.Map, Pobladovampiro.X, Pobladovampiro.Y, True)
Case "ENANO"
Call WarpUserChar(UserIndex, Pobladoenano.Map, Pobladoenano.X, Pobladoenano.Y, True)
Case "GNOMO"
Call WarpUserChar(UserIndex, Pobladoenano.Map, Pobladoenano.X, Pobladoenano.Y, True)
Case "GOBLIN"
Call WarpUserChar(UserIndex, Pobladoenano.Map, Pobladoenano.X, Pobladoenano.Y, True)

End Select
        
        
        
        
        Call SendData(ToIndex, UserIndex, 0, "!!Te has convertido en un REMORT, cuando vuelvas a entrar se habrán realizado los cambios necesarios en tu Pj.")
        
        Call CloseUser(UserIndex)
    End If

Exit Sub
fallo:
Call LogError("doremort " & Err.number & " D: " & Err.Description)

End Sub
