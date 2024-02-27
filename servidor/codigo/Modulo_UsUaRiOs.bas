Attribute VB_Name = "UsUaRiOs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)
    Dim DaExp As Integer
    Dim EraCriminal As Boolean
    
    DaExp = CInt(UserList(VictimIndex).Stats.ELV) * 2
    
    With UserList(attackerIndex)
        .Stats.Exp = .Stats.Exp + DaExp
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        'Lo mata
        Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).Name & "!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(attackerIndex, "Has ganado " & DaExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_exp)
              
        Call WriteConsoleMsg(VictimIndex, "¡" & .Name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
        
        '28/02/2016 Irongete: Mapas 16 y 20 los ciudadanos no pierden el status
        If .Pos.Map = 20 Or .Pos.Map = 16 Then Exit Sub
        
        '14/02/2016 Lorwik: Añado el Gran Poder como excepcion para atacar entre ciudadanos.
        If TriggerZonaPelea(VictimIndex, attackerIndex) <> TRIGGER6_PERMITE Or Not VictimIndex = GranPoder Then
            EraCriminal = criminal(attackerIndex)
            
            With .Reputacion
                If Not criminal(VictimIndex) Then
                    .AsesinoRep = .AsesinoRep + vlASESINO * 2
                    If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
                    .BurguesRep = 0
                    .NobleRep = 0
                    .PlebeRep = 0
                Else
                    .NobleRep = .NobleRep + vlNoble
                    If .NobleRep > MAXREP Then .NobleRep = MAXREP
                End If
            End With
            
            If criminal(attackerIndex) Then
                If Not EraCriminal Then Call RefreshCharStatus(attackerIndex)
            Else
                If EraCriminal Then Call RefreshCharStatus(attackerIndex)
            End If
        End If
        
        'Call UserDie(VictimIndex)
        
        Call FlushBuffer(VictimIndex)
        
        'Log
        Call LogAsesinato(.Name & " asesino a " & UserList(VictimIndex).Name)
    End With
End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .flags.Muerto = 0
        .Stats.MinHP = .Stats.UserAtributos(eAtributos.Constitucion)
        
        If .Stats.MinHP > VidaMaxima(UserIndex) Then
            .Stats.MinHP = VidaMaxima(UserIndex)
        End If
        
        If .flags.Navegando = 1 Then
            Dim Barco As ObjData
            Barco = ObjData(.Invent.BarcoObjIndex)
            .Char.Head = 0
            
            If .Faccion.ArmadaReal = 1 Then
                .Char.body = iFragataReal
            ElseIf .Faccion.FuerzasCaos = 1 Then
                .Char.body = iFragataCaos
            Else
                If criminal(UserIndex) Then
                    Select Case Barco.Ropaje
                        Case iBarca
                            .Char.body = iBarcaPk
                        
                        Case iGalera
                            .Char.body = iGaleraPk
                        
                        Case iGaleon
                            .Char.body = iGaleonPk
                    End Select
                Else
                    Select Case Barco.Ropaje
                        Case iBarca
                            .Char.body = iBarcaCiuda
                        
                        Case iGalera
                            .Char.body = iGaleraCiuda
                        
                        Case iGaleon
                            .Char.body = iGaleonCiuda
                    End Select
                End If
            End If
            
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            Call DarCuerpoDesnudo(UserIndex)
            
            .Char.Head = .OrigChar.Head
        End If

        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteUpdateUserStats(UserIndex)
    End With
End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)

    With UserList(UserIndex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, Arma, Escudo, .FX, .loops, casco))
    End With
End Sub

Sub EnviarFama(ByVal UserIndex As Integer)
    Dim L As Long
    
    With UserList(UserIndex).Reputacion
        L = (-.AsesinoRep) + _
            (-.BandidoRep) + _
            .BurguesRep + _
            (-.LadronesRep) + _
            .NobleRep + _
            .PlebeRep
        L = Round(L / 6)
        
        .Promedio = L
    End With
    
    Call WriteFame(UserIndex)
End Sub

Public Sub EraseUserChar(ByVal UserIndex As Integer, ByVal IsAdminInvisible As Boolean)
'*************************************************
'Author: Unknownn
'Last Modification: 27/07/2012 - ^[GS]^
'*************************************************

On Error GoTo ErrorHandler
    
    With UserList(UserIndex)
        If .Char.CharIndex > 0 And .Char.CharIndex <= LastChar Then
    
            CharList(.Char.CharIndex) = 0
            
            If .Char.CharIndex = LastChar Then
                Do Until CharList(LastChar) > 0
                    LastChar = LastChar - 1
                    If LastChar <= 1 Then Exit Do
                Loop
            End If
            
            ' Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
            If IsAdminInvisible Then
                Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
            Else
                'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
            End If
        End If
        
        If MapaValido(.Pos.Map) Then
            Call QuitarUser(UserIndex, .Pos.Map)
            
            MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
        End If
        
        .Char.CharIndex = 0
    End With
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:

    Dim UserName As String
    Dim CharIndex As Integer
    
    If UserIndex > 0 Then
        UserName = UserList(UserIndex).Name
        CharIndex = UserList(UserIndex).Char.CharIndex
    End If

    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description & _
                  ". User: " & UserName & "(UI: " & UserIndex & " - CI: " & CharIndex & ")")
                  
End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 04/07/2009
'Refreshes the status and tag of UserIndex.
'04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
'*************************************************
    Dim klan As String
    Dim Barco As ObjData
    Dim esCriminal As Boolean
    
    With UserList(UserIndex)
        If .GuildIndex > 0 Then
            klan = modGuilds.GuildName(.GuildIndex)
            klan = " <" & klan & ">"
        End If
        
        esCriminal = criminal(UserIndex)
        
        If .showName Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, esCriminal, .Name))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, esCriminal, vbNullString))
        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            If .flags.Muerto = 1 Then
                .Char.body = iFragataFantasmal
            Else
                Barco = ObjData(.Invent.Object(.Invent.BarcoSlot).ObjIndex)
                
                If .Faccion.ArmadaReal = 1 Or .Faccion.Legion = 1 Then
                    .Char.body = iFragataReal
                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                    .Char.body = iFragataCaos
                Else
                    If esCriminal Then
                        Select Case Barco.Ropaje
                            Case iBarca
                                .Char.body = iBarcaPk
                            
                            Case iGalera
                                .Char.body = iGaleraPk
                            
                            Case iGaleon
                                .Char.body = iGaleonPk
                        End Select
                    Else
                        Select Case Barco.Ropaje
                            Case iBarca
                                .Char.body = iBarcaCiuda
                            
                            Case iGalera
                                .Char.body = iGaleraCiuda
                            
                            Case iGaleon
                                .Char.body = iGaleonCiuda
                        End Select
                    End If
                End If
            End If
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

Public Function MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 23/07/2009
'
'23/07/2009: Budi - Ahora se envía el nick
'*************************************************

On Error GoTo hayerror
    Dim CharIndex As Integer
    
    With UserList(UserIndex)
    
        If InMapBounds(Map, X, Y) Then
            'If needed make a new character in list
            If .Char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex) = UserIndex
            End If
            
            'Place character on map if needed
            If toMap Then MapData(Map, X, Y).UserIndex = UserIndex
            
            'Send make character command to clients
            Dim klan As String
            If .GuildIndex > 0 Then
                klan = modGuilds.GuildName(.GuildIndex)
            End If
            
            Dim bCr As Byte
            Dim bNick As String
            Dim bPriv As Byte
            
            If EsNewbie(UserIndex) = True Then 'Es newbie?
                bCr = 10
            ElseIf EsNewbie(UserIndex) = False Then
                If criminal(UserIndex) = True Then 'Es criminal?
                    bCr = 1
                Else 'Entonces sera ciudadano...
                    bCr = 0
                End If
            End If
            
            bPriv = .flags.Privilegios
            'Preparo el nick
            If .showName Then
                If UserList(sndIndex).flags.Privilegios And PlayerType.User Then
                    bNick = .Name
'                    bPriv = .flags.Privilegios
                Else
                    If .flags.invisible Or .flags.Oculto Then
                        bNick = .Name & " " & TAG_USER_INVISIBLE
                    Else
                        bNick = .Name
                    End If
'                    bPriv = .flags.Privilegios
                End If
            Else
                bNick = vbNullString
'                bPriv = PlayerType.User
            End If
            
            If Not toMap Then
                Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.heading, _
                            .Char.CharIndex, X, Y, _
                            .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, _
                            bNick, klan, bCr, bPriv, 2, 0, .flags.Speed, .PartyId, 0, 0, 0)
            Else
                'Hide the name and clan - set privs as normal user
                 Call AgregarUser(UserIndex, .Pos.Map)
            End If
            
        End If
    End With
    
    MakeUserChar = True
    
Exit Function

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
    'Resume Next
    Call CloseSocket(UserIndex)
End Function

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Sub CheckUserLevel(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 12/09/2007
'Chequea que el usuario no halla alcanzado el siguiente nivel,
'de lo contrario le da la vida, mana, etc, correspodiente.
'07/08/2006 Integer - Modificacion de los valores
'01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
'13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
'09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
'12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y está en un clan, se lo expulsa para no sumar antifacción
'02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
'*************************************************
    Dim Pts As Integer
    Dim AumentoHIT As Integer
    Dim AumentoMANA As Integer
    Dim AumentoSTA As Integer
    Dim AumentoHP As Integer
    Dim WasNewbie As Boolean
    Dim Promedio As Double
    Dim aux As Integer
    Dim DistVida(1 To 5) As Integer
    Dim GI As Integer 'Guild Index
    
On Error GoTo errHandler
    
    WasNewbie = EsNewbie(UserIndex)
    
    With UserList(UserIndex)
        If .Stats.Exp >= .Stats.ELU Then
           
            'Store it!
            Call Statistics.UserLevelUp(UserIndex)
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
            Call WriteConsoleMsg(UserIndex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
            
 
            .Stats.ELV = .Stats.ELV + 1
            
            .Stats.Exp = .Stats.Exp - .Stats.ELU
            
            If .Stats.ELV < 20 Then
                .Stats.ELU = .Stats.ELU * 1.1
            ElseIf .Stats.ELV < 30 Then
                .Stats.ELU = .Stats.ELU * 1.2
            ElseIf .Stats.ELV < 40 Then
                .Stats.ELU = .Stats.ELU * 1.3
            ElseIf .Stats.ELV < 50 Then
                .Stats.ELU = .Stats.ELU * 1.4
            ElseIf .Stats.ELV < 60 Then
                .Stats.ELU = .Stats.ELU * 1.5
            ElseIf .Stats.ELV < 70 Then
                .Stats.ELU = .Stats.ELU * 1.6
            ElseIf .Stats.ELV < 80 Then
                .Stats.ELU = .Stats.ELU * 1.7
            ElseIf .Stats.ELV < 90 Then
                .Stats.ELU = .Stats.ELU * 1.8
            ElseIf .Stats.ELV < 100 Then
                .Stats.ELU = .Stats.ELU * 1.9
            ElseIf .Stats.ELV < 110 Then
                .Stats.ELU = .Stats.ELU * 2
            ElseIf .Stats.ELV < 120 Then
                .Stats.ELU = .Stats.ELU * 2.1
            Else
                .Stats.ELU = .Stats.ELU * 2.2
            End If
            
  'Calculo subida de vida
            Promedio = ModVida(.clase) - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
            aux = RandomNumber(0, 100)

            If Promedio - Int(Promedio) = 0.5 Then
                'Es promedio semientero
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 0.5
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 0.5
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio - 0.5
                Else
                    AumentoHP = Promedio - 0.5
                End If
                
            Else
                'Es promedio entero
                
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 1
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio
                Else
                    AumentoHP = Promedio - 1
                End If
                
            End If
            

        
            Select Case .clase
                Case eClass.Warrior
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = 2.8 * .Stats.UserAtributos(eAtributos.Energia)
                
                Case eClass.Hunter
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = 2.8 * .Stats.UserAtributos(eAtributos.Energia)
                
                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = 2 * .Stats.UserAtributos(eAtributos.Energia)
                
                Case eClass.Mage
                    AumentoHIT = 1
                    AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = 1.5 * .Stats.UserAtributos(eAtributos.Energia)
                
                
                Case eClass.Cleric
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = 1.8 * .Stats.UserAtributos(eAtributos.Energia)
                
                Case eClass.Druid
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = 1.8 * .Stats.UserAtributos(eAtributos.Energia)
                
                Case eClass.Assasin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = 2.2 * .Stats.UserAtributos(eAtributos.Energia)
                
                Case eClass.Bard
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = 2 * .Stats.UserAtributos(eAtributos.Energia)
                
                Case Else
                    AumentoHIT = 2
                    AumentoSTA = 1.5 * .Stats.UserAtributos(eAtributos.Energia)
            End Select
            
            'Actualizamos HitPoints
            .Stats.MaxHP = .Stats.MaxHP + AumentoHP
            If .Stats.MaxHP > STAT_MAXHP Then .Stats.MaxHP = STAT_MAXHP
            
            'Actualizamos Stamina
            .Stats.MaxSta = .Stats.MaxSta + AumentoSTA
            If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
            
            'Actualizamos Mana
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA
            If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
            
            'Actualizamos Golpe Máximo
            .Stats.MaxHIT = .Stats.MaxHIT + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MaxHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
                    .Stats.MaxHIT = STAT_MAXHIT_OVER36
            End If
            
            'Actualizamos Golpe Mínimo
            .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MinHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
                    .Stats.MinHIT = STAT_MAXHIT_OVER36
            End If
            
            'Notificamos al user
            If AumentoHP > 0 Then
                Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoSTA > 0 Then
                Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de energia.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoMANA > 0 Then
                Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoHIT > 0 Then
                Call WriteConsoleMsg(UserIndex, "Tu golpe máximo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Tu golpe minimo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 12, 0))
            
            Call LogDesarrollo(.Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)
            

                'If user is in a party, we modify the variable p_sumaniveleselevados
                'Call mdParty.ActualizarSumaNivelesElevados(UserIndex)
                    'If user reaches lvl 25 and he is in a guild, we check the guild's alignment and expulses the user if guild has factionary alignment
        
            If .Stats.ELV = 25 Then
                GI = .GuildIndex
                If GI > 0 Then
                    If modGuilds.GuildAlignment(GI) = "Legión oscura" Or modGuilds.GuildAlignment(GI) = "Armada Real" Then
                        'We get here, so guild has factionary alignment, we have to expulse the user
                        Call modGuilds.m_EcharMiembroDeClan(-1, .Name)
                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                        Call WriteConsoleMsg(UserIndex, "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la Facción bajo la cual tu clan está alineado, estarás excluído del mismo.", FontTypeNames.FONTTYPE_GUILD)
                    End If
                End If
            End If

        End If
        
        'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
        If Not EsNewbie(UserIndex) And WasNewbie Then
            Call QuitarNewbieObj(UserIndex)
            If UCase$(MapInfo(.Pos.Map).Restringir) = "NEWBIE" Then
                Call WarpUserChar(UserIndex, 1, 50, 50, True)
                Call WriteConsoleMsg(UserIndex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        Call RefreshCharStatus(UserIndex)

    End With
    
    Call WriteUpdateUserStats(UserIndex)
Exit Sub

errHandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.Description)
    Debug.Print "Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.Description
End Sub

Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1 _
                    Or UserList(UserIndex).flags.Vuela = 1
End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)
'***************************************************
'Author: Unknown
'Revisión: Edurne
'Last Modification: 10/09/2015
'***************************************************
    Dim nPos As worldPos
    Dim sailing As Boolean
    Dim CasperIndex As Integer
    Dim CasperHeading As eHeading
    Dim isAdminInvi As Boolean
    
    
    With UserList(UserIndex)
        sailing = PuedeAtravesarAgua(UserIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
            
        isAdminInvi = (.flags.AdminInvisible = 1)
            
        If MoveToLegalPos(.Pos.Map, nPos.X, nPos.Y, sailing, Not sailing) Then
            'si no estoy solo en el mapa...
            If MapInfo(.Pos.Map).NumUsers > 1 Then
                   
                CasperIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex
                'Si hay un usuario, y paso la validacion, entonces es un casper
                If CasperIndex > 0 Then
                    ' Los admins invisibles no pueden patear caspers
                    If Not isAdminInvi Then
                        
                        If TriggerZonaPelea(UserIndex, CasperIndex) = TRIGGER6_PROHIBE Then
                            If UserList(CasperIndex).flags.SeguroResu = False Then
                                UserList(CasperIndex).flags.SeguroResu = True
                                Call WriteConsoleMsg(UserIndex, "Seguro de resurrección activado.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
        
                        With UserList(CasperIndex)
                            CasperHeading = InvertHeading(nHeading)
                            Call HeadtoPos(CasperHeading, .Pos)
                            
                            ' Si es un admin invisible, no se avisa a los demas clientes
                            If Not .flags.AdminInvisible = 1 Then _
                                Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y))
                            
                            Call WriteForceCharMove(CasperIndex, CasperHeading)
                                
                            'Update map and char
                            .Char.heading = CasperHeading
                            MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = CasperIndex
                        
                        End With
                    
                        'Actualizamos las áreas de ser necesario
                        Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
                        
                        '11/12/2018 Irongete: Actualizo las zonas
                        Call Drag_Zonas.comprobar_zona(CasperIndex)
                    End If
                End If
    
                
                ' Si es un admin invisible, no se avisa a los demas clientes
                If Not isAdminInvi Then _
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))
                
            End If
            
            ' Los admins invisibles no pueden patear caspers
            If Not (isAdminInvi And (CasperIndex <> 0)) Then
                Dim oldUserIndex As Integer
                
                oldUserIndex = MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex
                
                ' Si no hay intercambio de pos con nadie
                If oldUserIndex = UserIndex Then
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
                End If
                
                .Pos = nPos
                .Char.heading = nHeading
                MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                
                Call DoTileEvents(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
                
                'Actualizamos las áreas de ser necesario
                Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
                
                '11/12/2018 Irongete: Actualizo las zonas
                Call Drag_Zonas.comprobar_zona(UserIndex)
            Else
                Call WritePosUpdate(UserIndex)
            End If
    
        Else
            Call WritePosUpdate(UserIndex)
        End If
        
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
    
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
        
        
        If .flags.Muerto = 0 Then
            If .flags.Privilegios = PlayerType.User Then Call VigilarEventosTrampas(UserIndex)
            If .Pos.Map = MapInvocacion Then Call VigilarEventosInvocacion(UserIndex)
        End If
        
    End With
        
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
'*************************************************
'Author: ZaMa
'Last modified: 30/03/2009
'Returns the heading opposite to the one passed by val.
'*************************************************
    Select Case nHeading
        Case eHeading.EAST
            InvertHeading = WEST
        Case eHeading.WEST
            InvertHeading = EAST
        Case eHeading.SOUTH
            InvertHeading = NORTH
        Case eHeading.NORTH
            InvertHeading = SOUTH
    End Select
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserObj)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    UserList(UserIndex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(UserIndex, Slot)
End Sub

Function NextOpenCharIndex() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then _
                LastChar = LoopC
            
            Exit Function
        End If
    Next LoopC
End Function

Function NextOpenUser() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC
End Function

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
    Dim GuildI As Integer
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHP & "/" & VidaMaxima(UserIndex) & "  Mana: " & .Stats.MinMAN & "/" & ManaMaxima(UserIndex) & "  Energia: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT & " (" & ObjData(.Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        GuildI = .GuildIndex
        If GuildI > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
            If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.Name) Then
                Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_INFO)
            End If
            'guildpts no tienen objeto
        End If
        
#If ConUpTime Then
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim TempStr As String
        TempDate = Now - .LogOnTime
        TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
#End If
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & .Stats.GLD & "  Posicion: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.Map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Dados: " & FuerzaMaxima(UserIndex) & ", " & AgilidadMaxima(UserIndex) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Energia) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
        
        If .Faccion.ArmadaReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Armada Real Desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en Nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " Ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.Legion = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legion Desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en Nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " Ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
            
        ElseIf .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legion Oscura Desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en Nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Armada Real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        'ElseIf .Faccion.RecibioExpInicialFaccion = 1 Then
        '    Call WriteConsoleMsg(sendIndex, "Fue Legion", FontTypeNames.FONTTYPE_INFO)
        '    Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
            
        ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Caos", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
        
        If .GuildIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is offline.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
    Dim CharFile As String
    Dim Ban As String
    Dim BanDetailPath As String
    
    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(sendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        
        If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Armada Real Desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en Nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " Ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legion Oscura Desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en Nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Armada Real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Legionario", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        End If

        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)
        
        If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)
        End If
        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        
        If Ban = "1" Then
            Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(sendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next

    Dim j As Long
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            If .Invent.Object(j).ObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(.Invent.Object(j).ObjIndex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    End With
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
    Dim j As Integer
    
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
    Next j
    
End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

    If NPCList(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not criminal(NPCList(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(NPCList(NpcIndex).MaestroUser, "¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 06/28/2008
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
'06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
'**********************************************
    Dim EraCriminal As Boolean
    
    'Guardamos el usuario que ataco el npc.
    NPCList(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name
    
    'Npc que estabas atacando.
    Dim LastNpcHit As Integer
    LastNpcHit = UserList(UserIndex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(UserIndex).flags.NPCAtacado = NpcIndex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If NPCList(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If NPCList(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                NPCList(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
        NPCList(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).Name
    ElseIf NPCList(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).Name Then
        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If NPCList(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                NPCList(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
    End If
    
    If NPCList(NpcIndex).MaestroUser > 0 Then
        If NPCList(NpcIndex).MaestroUser <> UserIndex Then
            Call AllMascotasAtacanUser(UserIndex, NPCList(NpcIndex).MaestroUser)
        End If
    End If
    
    If EsMascotaCiudadano(NpcIndex, UserIndex) Then
        Call VolverCriminal(UserIndex)
        NPCList(NpcIndex).Movement = TipoAI.NPCDEFENSA
        NPCList(NpcIndex).Hostile = 1
    Else
        EraCriminal = criminal(UserIndex)
        
        'Reputacion
        If NPCList(NpcIndex).Stats.Alineacion = 0 Then
           If NPCList(NpcIndex).NPCType = eNPCType.GuardiaReal Then
                Call VolverCriminal(UserIndex)
           Else
                If Not NPCList(NpcIndex).MaestroUser > 0 Then   'mascotas nooo!
                
                    '28/02/2016 Irongete: A los dummys tampoco
                    If NPCList(NpcIndex).Numero = 624 Then Exit Sub
                    
                    Call VolverCriminal(UserIndex)
                End If
           End If
        
        ElseIf NPCList(NpcIndex).Stats.Alineacion = 1 Then
            If esCaos(UserIndex) = False Then
                UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR / 2
                If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
                 UserList(UserIndex).Reputacion.PlebeRep = MAXREP
            End If
        End If
        
        If NPCList(NpcIndex).MaestroUser <> UserIndex Then
            'hacemos que el npc se defienda
            NPCList(NpcIndex).Movement = TipoAI.NPCDEFENSA
            NPCList(NpcIndex).Hostile = 1
        End If
        
        If EraCriminal And Not criminal(UserIndex) Then
            Call VolverCiudadano(UserIndex)
        End If
    End If
End Sub
Public Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1 Then
            PuedeApuñalar = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR _
                        Or UserList(UserIndex).clase = eClass.Assasin
        End If
    End If
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)

    With UserList(UserIndex)
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            
            If .Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
            
            Dim Lvl As Integer
            Lvl = .Stats.ELV
            
            If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
            
            If .Stats.UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
            
            Dim Prob As Byte
            
            Prob = 10 'Probabilidad para que suba skill

            If RandomNumber(1, Prob) <= 4 Then
                .Stats.UserSkills(Skill) = .Stats.UserSkills(Skill) + 1
                Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & .Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                
                .Stats.Exp = .Stats.Exp + 5
                If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
                
                Call WriteConsoleMsg(UserIndex, "¡Has ganado 5 puntos de experiencia!", FontTypeNames.FONTTYPE_exp)
                
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
            End If
        End If
    End With
End Sub

Sub Desequipar_Todo(ByVal UserIndex As Integer)
'***************************************************
'Author: Edurne
'Last Modification: 10/09/2015
'Info: DESEQUIPA TODOS LOS OBJETOS
'***************************************************
    With UserList(UserIndex).Invent
        'desequipar armadura
        If .ArmourEqpObjIndex > 0 Then
            Debug.Print .ArmourEqpSlot
            Call Desequipar(UserIndex, .ArmourEqpSlot)
        End If
        
        'desequipar arma
        If .WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .WeaponEqpSlot)
        End If
        
        'desequipar casco
        If .CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .CascoEqpSlot)
        End If
        
        'desequipar herramienta
        If .AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, .AnilloEqpSlot)
        End If
        
        'desequipar municiones
        If .MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .MunicionEqpSlot)
        End If
        
        'desequipar escudo
        If .EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .EscudoEqpSlot)
        End If
    End With

End Sub

Sub Reset_flags_muerte(ByVal UserIndex As Integer)
'***************************************************
'Author: Edurne
'Last Modification: 10/09/2015
'Info: Reinicia todos los flags a su valor base tras morir
'***************************************************
    Dim b As Byte
    With UserList(UserIndex)
        With .flags
            .AtacadoPorUser = 0
            .Envenenado = 0
            .Muerto = 1
        
            If .AtacadoPorNpc > 0 Then
                NPCList(.AtacadoPorNpc).Movement = NPCList(.AtacadoPorNpc).flags.OldMovement
                NPCList(.AtacadoPorNpc).Hostile = NPCList(.AtacadoPorNpc).flags.OldHostil
                NPCList(.AtacadoPorNpc).flags.AttackedBy = vbNullString
            End If
            
            If .NPCAtacado > 0 Then
                If NPCList(.NPCAtacado).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                    NPCList(.NPCAtacado).flags.AttackedFirstBy = vbNullString
                End If
            End If
            .AtacadoPorNpc = 0
            .NPCAtacado = 0
            
            'Quitamos lo relativo a la montura, no hace falta más que estos dos flags.
            .QueMontura = 0
            
            '<<<< Paralisis >>>>
            If .Paralizado = 1 Then
                .Paralizado = 0
                Call WriteParalizeOK(UserIndex)
            End If
            
            '<<< Estupidez >>>
            If .Estupidez = 1 Then
                .Estupidez = 0
                Call WriteDumbNoMore(UserIndex)
            End If
            
            '<<<< Descansando >>>>
            If .Descansar Then
                .Descansar = False
                Call WriteRestOK(UserIndex)
            End If
            
            '<<<< Meditando >>>>
            If .Meditando Then
                .Meditando = False
                Call WriteMeditateToggle(UserIndex)
            End If
            
            '<<<< Invisible >>>>
            If .invisible = 1 Or .Oculto = 1 Then
                .Oculto = 0
                .invisible = 0
                
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
            End If
        End With
        
        With .Counters
            .TiempoOculto = 0
            .Invisibilidad = 0
            .Morph = 0
            .Ceguera = 0
            .Estupidez = 0
            .Mimetismo = 0
            .Veneno = 0
        End With
        
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 0
        End If
        
        ' << Restauramos el mimetismo
        If .flags.Mimetizado = 1 Then
            .Char.body = .CharMimetizado.body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
        End If
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then
            For b = 1 To 5
                .Stats.UserAtributos(b) = .Stats.UserAtributosBackUP(b)
            Next b
        End If
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
        
        
        For b = 1 To MAXMASCOTAS
            If .MascotasIndex(b) > 0 Then
                Call MuereNpc(.MascotasIndex(b), 0)
            ' Si estan en agua o zona segura
            Else
                .MascotasType(b) = 0
            End If
        Next b
        
        .NroMascotas = 0

    End With

End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal UserIndex As Integer, Optional ByVal Atacante As Integer = 0)
'************************************************
'Author: Uknown
'Last Modified: 21/07/2009
'04/15/2008: NicoNZ - Ahora se resetea el counter del invi
'13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
'27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
'21/07/2009: Marco - Al morir se desactiva el comercio seguro.
'10/09/2015: Edurne - Dejamos el control de desequipamiento y reseteo de flags a otros subs
'************************************************
On Error GoTo ErrorHandler
    Dim i As Long
    Dim aN As Integer
    
    With UserList(UserIndex)
    
        If .Name = "magognomos" Then
          Exit Sub
        End If
        
        
    
        'Sonido
        If .genero = eGenero.Mujer Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_MUJER)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_HOMBRE)
        End If
        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
    
        '28/10/2015 Irongete: Si el jugador esta en un duelo no morirá
        '¡FALTA! -> Ya que no muere, hay que añadir en la funcion TerminarDuelosSet que si es el segundo duelo sí que muera y se le caigan los objetos
        If UserList(UserIndex).flags.EstaDueleandoSet = True Then
            UserList(UserIndex).flags.PerdioRondaSet = UserList(UserIndex).flags.PerdioRondaSet + 1
            Call TerminarDueloSet(UserList(UserIndex).flags.OponenteSet, UserIndex)
            Exit Sub
        End If
        
        .Stats.MinHP = 0
        .Stats.MinSta = 0
    
        '<Edurne>
        Call Reset_flags_muerte(UserIndex)
        '</Edurne>
        
        If TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE Then _
            Call TirarTodo(UserIndex)

        
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            .Char.body = iFragataFantasmal
        End If
        
        '<< Actualizamos clientes >>
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
        Call WriteUpdateUserStats(UserIndex)
        

        '<<Cerramos comercio seguro>>
        Call LimpiarComercioSeguro(UserIndex)

        If .Pos.Map = 20 Then _
            Call WarpUserChar(UserIndex, 1, 44, 88, True)

        '15/02/2016 Lorwik: Traslado esto aqui, si el usuario muere pierde el gran poder.
        If UserIndex = GranPoder Then
            If Atacante = 0 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha perdido el Gran Poder", FontTypeNames.FONTTYPE_WARNING))
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(Atacante).Name & "le ha quitado el Gran poder a " & .Name, FontTypeNames.FONTTYPE_WARNING))
            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FxGranPoder, 0))
            Call OtorgarGranPoder(Atacante)
        End If

    End With
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    
    With UserList(Atacante)
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        
        If criminal(Muerto) Then
            If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                .flags.LastCrimMatado = UserList(Muerto).Name
                If .Faccion.CriminalesMatados < MAXUSERMATADOS Then _
                    .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + 1
            End If
            
            If .Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
                .Faccion.Reenlistadas = 200  'jaja que trucho
                
                'con esto evitamos que se vuelva a reenlistar
            End If
        Else
            If .flags.LastCiudMatado <> UserList(Muerto).Name Then
                .flags.LastCiudMatado = UserList(Muerto).Name
                If .Faccion.CiudadanosMatados < MAXUSERMATADOS Then _
                    .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
            End If
        End If
        
        If .Stats.UsuariosMatados < MAXUSERMATADOS Then _
            .Stats.UsuariosMatados = .Stats.UsuariosMatados + 1
            
    End With
End Sub

Sub Tilelibre(ByRef Pos As worldPos, ByRef nPos As worldPos, ByRef Obj As Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'**************************************************************
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    Dim hayobj As Boolean
    
    hayobj = False
    nPos.Map = Pos.Map
    nPos.X = 0
    nPos.Y = 0
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
        If LoopC > 15 Then
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
                
                If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        
                        'break both fors
                        tX = Pos.X + LoopC
                        tY = Pos.Y + LoopC
                    End If
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
    Loop
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal FX As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 15/07/2009
'15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
'**************************************************************
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    With UserList(UserIndex)
    
        '¿Está trabajando?
        If .flags.Makro <> 0 Then
            Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
            .flags.Makro = 0
        End If

        If Not IntervaloPermiteCambiardeMapa(UserIndex) And .flags.PuedeCambiarMapa Then
            If Not UserList(UserIndex).flags.UltimoMensaje = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡Estas en combate! No puedes pasar de mapa en estos momentos.", FontTypeNames.FONTTYPE_INFOBOLD)
                UserList(UserIndex).flags.UltimoMensaje = 1
            End If
            Exit Sub
        End If
        
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        Call WriteRemoveAllDialogs(UserIndex)
        
        OldMap = .Pos.Map
        OldX = .Pos.X
        OldY = .Pos.Y

        Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)
        
        If OldMap <> Map Then
            Call WriteChangeMap(UserIndex, Map, MapInfo(.Pos.Map).MapVersion)
            
            'Update new Map Users
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
            
            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If
            
            '15/12/2018 Irongete: Le mando las zonas del mapa
            Call EnviarZonas(UserIndex, Map)
            
        End If
        
        .Pos.X = X
        .Pos.Y = Y
        .Pos.Map = Map
        
        Call MakeUserChar(True, Map, UserIndex, Map, X, Y)
        Call WriteUserCharIndexInServer(UserIndex)
        
        'Force a flush, so user index is in there before it's destroyed for teleporting
        Call FlushBuffer(UserIndex)
        
        'Seguis invisible al pasar de mapa
        If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            If permiso_en_zona(UserIndex) And permiso_zona.no_invisibilidad Then
                .flags.Oculto = 0
                .flags.invisible = 0
                Call SetInvisible(UserIndex, .Char.CharIndex, False)
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SetInvisible(UserIndex, .Char.CharIndex, True)
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            End If
        End If
        
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
        End If
        
        If .NroMascotas Then Call WarpMascotas(UserIndex)
        
        ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
        Call IntervaloPermiteSerAtacado(UserIndex, True)
        
        ' Automatic toogle navigate
        If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) = 0 Then
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                If .flags.Navegando = 0 Then
                    .flags.Navegando = 1
                        
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(UserIndex)
                End If
            Else
                If .flags.Navegando = 1 Then
                    .flags.Navegando = 0
                            
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(UserIndex)
                End If
            End If
        End If
      
    End With
End Sub

Private Sub WarpMascotas(ByVal UserIndex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 11/05/2009
'13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
'13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
'11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
'************************************************
    Dim i As Integer
    Dim petType As Integer
    Dim PetRespawn As Boolean
    Dim PetTiempoDeVida As Integer
    Dim NroPets As Integer
    Dim InvocadosMatados As Integer
    Dim canWarp As Boolean
    Dim Index As Integer
    Dim iMinHP As Integer
    
    NroPets = UserList(UserIndex).NroMascotas
    canWarp = (MapInfo(UserList(UserIndex).Pos.Map).Pk = True)
    
    For i = 1 To MAXMASCOTAS
        Index = UserList(UserIndex).MascotasIndex(i)
        
        If Index > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
            If NPCList(Index).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(Index)
                UserList(UserIndex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
                
                petType = 0
            Else
                'Store data and remove NPC to recreate it after warp
                'PetRespawn = NPCList(index).flags.Respawn = 0
                petType = UserList(UserIndex).MascotasType(i)
                'PetTiempoDeVida = NPCList(index).Contadores.TiempoExistencia
                
                ' Guardamos el hp, para restaurarlo uando se cree el npc
                iMinHP = NPCList(Index).Stats.MinHP
                
                Call QuitarNPC(Index)
                
                ' Restauramos el valor de la variable
                UserList(UserIndex).MascotasType(i) = petType

            End If
        ElseIf UserList(UserIndex).MascotasType(i) > 0 Then
            'Store data and remove NPC to recreate it after warp
            PetRespawn = True
            petType = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida = 0
        Else
            petType = 0
        End If
        
        If petType > 0 And canWarp Then
            Index = SpawnNpc(petType, UserList(UserIndex).Pos, False, PetRespawn)
            
            'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
            ' Exception: Pets don't spawn in water if they can't swim
            If Index = 0 Then
                Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            Else
                UserList(UserIndex).MascotasIndex(i) = Index

                ' Nos aseguramos de que conserve el hp, si estaba dañado
                NPCList(Index).Stats.MinHP = IIf(iMinHP = 0, NPCList(Index).Stats.MinHP, iMinHP)
            
                NPCList(Index).MaestroUser = UserIndex
                NPCList(Index).Movement = TipoAI.SigueAmo
                NPCList(Index).Target = 0
                NPCList(Index).TargetNPC = 0
                NPCList(Index).Contadores.TiempoExistencia = PetTiempoDeVida
                Call FollowAmo(Index)
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If Not canWarp Then
        Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    UserList(UserIndex).NroMascotas = NroPets
End Sub

''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 09/04/08 (NicoNZ)
'
'***************************************************
    Dim isNotVisible As Boolean
    Dim i As Byte
    
    With UserList(UserIndex)
    If .flags.UserLogged And Not .Counters.Saliendo Then
    
        'Si esta en arenas no puede salir
        If .flags.ArenaRinkel = True Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir mientras estas en arenas..", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        .Counters.Saliendo = True
        .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.Map).Pk, IntervaloCerrarConexion, 0)
        
        isNotVisible = (.flags.Oculto Or .flags.invisible)
        If isNotVisible Then
            .flags.Oculto = 0
            .flags.invisible = 0
            Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call SetInvisible(UserIndex, .Char.CharIndex, False)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
        End If
        
        For i = 0 To 1
            If Oponente(i) = UserIndex Then
                Oponente(0) = 0
                Oponente(1) = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El oponente salio del juego, la cola se reseteara.", FontTypeNames.FONTTYPE_DIOS))
            End If
        Next i
        
        Call WriteConsoleMsg(UserIndex, "Cerrando...Se cerrará el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If UserList(UserIndex).flags.EstaDueleandoSet = True Then
        Call DesconectarDueloSet(UserList(UserIndex).flags.OponenteSet, UserIndex)
    End If
    
    '14/02/2016 Lorwik: Si esta en el mapa de duelos se le lleva a nix.
    Call DesconectarDuelos(UserIndex)
    End With
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/08
'
'***************************************************
    If UserList(UserIndex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Pk, IntervaloCerrarConexion, 0)
        End If
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
    Dim ViejoNick As String
    Dim ViejoCharBackup As String
    
    If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
    ViejoNick = UserList(UserIndexDestino).Name
    
    If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
    End If
End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal nombre As String)
    If FileExist(CharPath & nombre & ".chr", vbArchive) = False Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Energia: " & GetVar(CharPath & nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        
#If ConUpTime Then
        Dim TempSecs As Long
        Dim TempStr As String
        TempSecs = GetVar(CharPath & nombre & ".chr", "INIT", "UpTime")
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
#End If
    
    End If
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
    Dim CharFile As String
    
On Error Resume Next
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente
'**************************************************************
    With UserList(UserIndex)
        'Si es newbie no puede ser criminal (Es chico y no sabe lo que hace xD)
        If EsNewbie(UserIndex) = True Then Exit Sub
    
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            .Reputacion.BurguesRep = 0
            .Reputacion.NobleRep = 0
            .Reputacion.PlebeRep = 0
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO
            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
            'If .Faccion.Legion = 1 Then Call ExpulsarFaccionLegion(UserIndex)
        End If
    End With
    
    Call RefreshCharStatus(UserIndex)
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente.
'**************************************************************
    With UserList(UserIndex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        .Reputacion.LadronesRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.AsesinoRep = 0
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO
        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
    End With
    
    Call RefreshCharStatus(UserIndex)
End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/07/2008
'Checks if a given body index is a boat
'**************************************************************
'TODO : This should be checked somehow else. This is nasty....
    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or _
            body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or _
            body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function

Public Sub SetInvisible(ByVal UserIndex As Integer, _
                        ByVal userCharIndex As Integer, _
                        ByVal invisible As Boolean)
    Dim sndNick As String
    Dim klan    As String

    Call SendData(SendTarget.ToUsersAreaButGMs, UserIndex, PrepareMessageSetInvisible(userCharIndex, invisible))

    If invisible Then
        sndNick = UserList(UserIndex).Name & " " & TAG_USER_INVISIBLE
    Else
        sndNick = UserList(UserIndex).Name

        If UserList(UserIndex).GuildIndex > 0 Then
            sndNick = sndNick & " <" & modGuilds.GuildName(UserList(UserIndex).GuildIndex) & ">"
        End If
    End If

    Call SendData(SendTarget.ToGMsArea, UserIndex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))
End Sub

Public Function VidaMaxima(ByVal UserIndex As Integer) As Integer
    'Devuelve la vida maxima sumado al modificador de aumento de vida
    VidaMaxima = UserList(UserIndex).Stats.MaxHP + UserList(UserIndex).flags.AumentodeVida
End Function

Public Function ManaMaxima(ByVal UserIndex As Integer) As Integer
    'Devuelve la maná maxima sumado al modificador de aumento de mana
    ManaMaxima = UserList(UserIndex).Stats.MaxMAN + UserList(UserIndex).flags.AumentodeMana
End Function

Public Function FuerzaMaxima(ByVal UserIndex As Integer) As Integer
    'Devuelve la Fuerza maxima sumado al modificador de aumento de Fuerza
    FuerzaMaxima = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + UserList(UserIndex).flags.AumentodeFuerza
End Function

Public Function AgilidadMaxima(ByVal UserIndex As Integer) As Integer
    'Devuelve la Fuerza maxima sumado al modificador de aumento de Fuerza
    AgilidadMaxima = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + UserList(UserIndex).flags.AumentodeAgilidad
End Function
