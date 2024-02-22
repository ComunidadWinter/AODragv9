Attribute VB_Name = "ModViajes"
Sub SistemaViajes(ByVal UserIndex As Integer, rdata As String)
'On Error GoTo fallo
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_VIAJERO Then Exit Sub
        '¿Esta en NIX?
        If UserList(UserIndex).Pos.Map = 34 Then
        If rdata = "ULLA" Then
        Viaje = 3000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 3000
            Call WarpUserChar(UserIndex, 1, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "CAOS" Then
        Viaje = 4000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 4000
            Call WarpUserChar(UserIndex, 170, 23, 78, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "DESCANSO" Then
        Viaje = 6000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6000
            Call WarpUserChar(UserIndex, 81, 85, 58, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "ATLANTIS" Then
        Viaje = 6500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 85, 70, 43, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        End If
        
                '¿Esta en Ulla?
        If UserList(UserIndex).Pos.Map = 1 Then
        If rdata = "NIX" Then
        Viaje = 3000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 3000
            Call WarpUserChar(UserIndex, 34, 57, 79, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "CAOS" Then
        Viaje = 3000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 3000
            Call WarpUserChar(UserIndex, 170, 23, 78, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "DESCANSO" Then
        Viaje = 3500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 3500
            Call WarpUserChar(UserIndex, 81, 85, 58, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "BANDER" Then
        Viaje = 5500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 5500
            Call WarpUserChar(UserIndex, 59, 51, 44, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "RINKEL" Then
        Viaje = 3500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 3500
            Call WarpUserChar(UserIndex, 20, 29, 92, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        End If
        
        '¿Esta en Descanso?
        If UserList(UserIndex).Pos.Map = 81 Then
        If rdata = "NIX" Then
        Viaje = 6000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6000
            Call WarpUserChar(UserIndex, 34, 57, 79, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "BANDER" Then
        Viaje = 3500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 3500
            Call WarpUserChar(UserIndex, 59, 51, 44, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "ULLA" Then
        Viaje = 3500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 3500
            Call WarpUserChar(UserIndex, 1, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "CAOS" Then
        Viaje = 5500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 5500
            Call WarpUserChar(UserIndex, 170, 23, 78, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "ARGHAL" Then
        Viaje = 12000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 12000
            Call WarpUserChar(UserIndex, 150, 35, 29, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        End If
        
       '¿Esta en Rinkel?
        If UserList(UserIndex).Pos.Map = 20 Then
        If rdata = "ULLA" Then
        Viaje = 3500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 3500
            Call WarpUserChar(UserIndex, 1, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "LINDOS" Then
        Viaje = 5500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 5500
            Call WarpUserChar(UserIndex, 63, 54, 14, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "ATLANTIS" Then
        Viaje = 9500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 9500
            Call WarpUserChar(UserIndex, 85, 70, 43, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "ESPERANZA" Then
        Viaje = 12500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 12500
            Call WarpUserChar(UserIndex, 111, 86, 76, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
         Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        End If
        
        '¿Esta en CAOS?
        If UserList(UserIndex).Pos.Map = 170 Then
        If rdata = "NIX" Then
        Viaje = 4000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 4000
            Call WarpUserChar(UserIndex, 34, 57, 79, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "ULLA" Then
        Viaje = 2500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 2500
            Call WarpUserChar(UserIndex, 1, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        If rdata = "LINDOS" Then
        Viaje = 6500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
    If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 63, 54, 14, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "DESCANSO" Then
        Viaje = 5500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 5500
            Call WarpUserChar(UserIndex, 81, 85, 58, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
        End If
        End If
        End If
        
                '¿Esta en ARGHAL?
        If UserList(UserIndex).Pos.Map = 151 Then
        If rdata = "DESCANSO" Then
        Viaje = 12000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 81, 36, 86, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "BANDER" Then
        Viaje = 12000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 59, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        End If
        
        '¿Esta en ATLANTIS?
        If UserList(UserIndex).Pos.Map = 85 Then
        If rdata = "NIX" Then
        Viaje = 6500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 34, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "BANDER" Then
        Viaje = 10000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 59, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "RINKEL" Then
        Viaje = 9500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 20, 16, 86, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        End If
        
                '¿Esta en LINDOS?
        If UserList(UserIndex).Pos.Map = 63 Then
        If rdata = "RINKEL" Then
        Viaje = 5500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 20, 16, 86, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "ESPERANZA" Then
        Viaje = 7500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 7500
            Call WarpUserChar(UserIndex, 111, 86, 76, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
          Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "CAOS" Then
        Viaje = 3500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 170, 24, 78, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        End If
        
        '¿Esta en ESPERANZA?
        If UserList(UserIndex).Pos.Map = 111 Then
        If rdata = "LINDOS" Then
        Viaje = 7500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 7500
            Call WarpUserChar(UserIndex, 63, 54, 14, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "RINKEL" Then
        Viaje = 12500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 20, 16, 86, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        End If
        '¿Esta en BANDER?
        If UserList(UserIndex).Pos.Map = 59 Then
        If rdata = "ULLA" Then
        Viaje = 5500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 1, 50, 50, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "DESCANSO" Then
        Viaje = 3500
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 6500
            Call WarpUserChar(UserIndex, 81, 38, 86, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "ATLANTIS" Then
        Viaje = 10000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 10000
            Call WarpUserChar(UserIndex, 85, 70, 43, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        If rdata = "ARGHAL" Then
        Viaje = 12000
        If UserList(UserIndex).Stats.GLD < Viaje Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes " & Viaje & " oros para viajar a " & rdata & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        Else
        If UserList(UserIndex).Stats.MinAGU < 1 Or UserList(UserIndex).Stats.MinHam < 1 Or UserList(UserIndex).Stats.MinSta < 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Apenas tienes energias, vuelve cuando estes preparado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If TieneObjetos(474, 1, UserIndex) Or TieneObjetos(475, 1, UserIndex) Or TieneObjetos(476, 1, UserIndex) And UserList(UserIndex).Stats.UserSkills(Navegacion) > 40 Then
            'le quitamos la stamina
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - UserList(UserIndex).Stats.MinSta
            'le ponemos el hambre y la sed a 0
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).Stats.MinHam = 0
            Viaje = 12000
            Call WarpUserChar(UserIndex, 150, 35, 29, True)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Viaje
            Else
            Call SendData(ToIndex, UserIndex, 0, "||Para viajar a una isla, necesitaras una embarcación y 40 puntos en navegación  " & "´" & FontTypeNames.FONTTYPE_info)
        End If
        End If
        End If
        End If
Call EnviarHambreYsed(UserIndex)

'fallo:
'Call LogError("sistemaviajes" & Err.number & " D: " & Err.Description)

End Sub
