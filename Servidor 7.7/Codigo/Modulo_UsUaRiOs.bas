Attribute VB_Name = "UsUaRiOs"
Option Explicit

Public Sub BorrarUsuario(ByVal UserName As String)
'on error Resume Next
'If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
'    Kill CharPath & UCase$(UserName) & ".chr"
'End If
End Sub

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
On Error GoTo fallo
Dim DaExp As Long

'pluto:7.0------------------------------------------
If UserList(AttackerIndex).Mision.Level > 0 Then
        If UserList(VictimIndex).Stats.ELV >= UserList(AttackerIndex).Mision.Level Then
        UserList(AttackerIndex).Mision.PjConseguidos = UserList(AttackerIndex).Mision.PjConseguidos + 1
        End If
End If
'------------------------------------------------
'PLUTO:2.4.1
Dim aa As Integer
aa = RandomNumber(1, 30)
DaExp = CInt((UserList(VictimIndex).Stats.ELV * 2) + aa)
'pluto:6.2
'If UserList(VictimIndex).Name = "Jaba" Then DaExp = 1000000

Call AddtoVar(UserList(AttackerIndex).Stats.exp, DaExp, MAXEXP)
         
'Lo mata
Call SendData(ToIndex, AttackerIndex, 0, "||Has matado " & UserList(VictimIndex).Name & "!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendData(ToIndex, AttackerIndex, 0, "||Has ganado " & DaExp & " puntos de experiencia." & "´" & FontTypeNames.FONTTYPE_FIGHT)
          
Call SendData(ToIndex, VictimIndex, 0, "||" & UserList(AttackerIndex).Name & " te ha matado!" & "´" & FontTypeNames.FONTTYPE_FIGHT)

'pluto:2.6.0 añade fortaleza
'pluto:6.8 añade torneo2 y ciudades y salas clan
If MapInfo(UserList(AttackerIndex).Pos.Map).Pk = True And UserList(AttackerIndex).Pos.Map <> 185 And MapInfo(UserList(AttackerIndex).Pos.Map).Terreno <> "TORNEO" And MapInfo(UserList(AttackerIndex).Pos.Map).Terreno <> "CASTILLO" And MapInfo(UserList(AttackerIndex).Pos.Map).Zona <> "CLAN" Then

If Not Criminal(VictimIndex) Then
         Call AddtoVar(UserList(AttackerIndex).Reputacion.AsesinoRep, vlASESINO * 2, MAXREP)
         UserList(AttackerIndex).Reputacion.BurguesRep = 0
         UserList(AttackerIndex).Reputacion.NobleRep = 0
         UserList(AttackerIndex).Reputacion.PlebeRep = 0
Else
         Call AddtoVar(UserList(AttackerIndex).Reputacion.NobleRep, vlNoble, MAXREP)
End If

End If

If UserList(AttackerIndex).MuertesTime > 6 Then
'pluto:2.10
Call SendData(ToAdmins, AttackerIndex, 0, "|| Posible Puente de Armadas en " & UserList(AttackerIndex).Name & " Mata a " & UserList(VictimIndex).Name & " --> " & UserList(AttackerIndex).MuertesTime & " Muertes/Minuto" & "´" & FontTypeNames.FONTTYPE_TALK)
End If

Call UserDie(VictimIndex)

Call AddtoVar(UserList(AttackerIndex).Stats.UsuariosMatados, 1, 31000)

'Log
Call LogAsesinato(UserList(AttackerIndex).Name & " asesino a " & UserList(VictimIndex).Name)
Exit Sub
fallo:
Call LogError("accstats " & Err.number & " D: " & Err.Description)

End Sub


Sub RevivirUsuario(ByVal UserIndex As Integer)
On Error GoTo fallo
'pluto:6.2-------------- aca tenes q fijarte para el tema del resumata
UserList(UserIndex).flags.Incor = True
UserList(UserIndex).Counters.Incor = 0
'-----------------------
UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = 10

Call DarCuerpoDesnudo(UserIndex)
'[GAU] Agregamo UserList(UserIndex).Char.Botas
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendUserStatsVida(UserIndex)
Exit Sub
fallo:
Call LogError("revivirusuario " & Err.number & " D: " & Err.Description)

End Sub
Sub RevivirUsuarioangel(ByVal UserIndex As Integer)
On Error GoTo fallo
'pluto:3-2-04
If Criminal(UserIndex) Then Exit Sub
'pluto:6.0A
If UserList(UserIndex).flags.Navegando > 0 Then Exit Sub
'pluto:6.2 - quito el aura (nati
'UserList(UserIndex).flags.Incor = True
'UserList(UserIndex).Counters.Incor = 0
'-----------
UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
Call SendData(ToIndex, UserIndex, 0, "||¡PODER DIVINO has ganado 500 puntos de nobleza!." & "´" & FontTypeNames.FONTTYPE_info)

Call DarCuerpoDesnudo(UserIndex)
'[GAU] Agregamo UserList(UserIndex).Char.Botas
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendUserStatsVida(UserIndex)
Exit Sub
fallo:
Call LogError("revivirusuarioangel " & Err.number & " D: " & Err.Description)

End Sub
'[GAU] Agregamo botas
'[GAU] Agregamo botas
Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer, ByVal Botas As Integer)

On Error GoTo fallo
'pluto:6.5
'If UserIndex = 0 Then Exit Sub

'[GAU]
UserList(UserIndex).Char.Botas = Botas
'[GAU]
UserList(UserIndex).Char.Body = Body
UserList(UserIndex).Char.Head = Head
UserList(UserIndex).Char.Heading = Heading
UserList(UserIndex).Char.WeaponAnim = Arma
UserList(UserIndex).Char.ShieldAnim = Escudo
UserList(UserIndex).Char.CascoAnim = Casco
'[GAU] Agregamo las botas
'Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & Casco & "," & Botas)
'pluto.6.2 quitamos el fx y el loop
Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & 0 & "," & 0 & "," & Casco & "," & Botas)

Exit Sub
fallo:
Call LogError("changeuserchar " & Err.number & " D: " & Err.Description)

End Sub



Sub EnviarSubirNivel(ByVal UserIndex As Integer, ByVal Puntos As Integer)
On Error GoTo fallo
Call SendData2(ToIndex, UserIndex, 0, 48, Puntos)
Exit Sub
fallo:
Call LogError("enviarsubirnivel " & Err.number & " D: " & Err.Description)


End Sub
Sub EnviaUnSkills(ByVal UserIndex As Integer, ByVal Skill As Integer)
On Error GoTo fallo

Call SendData(ToIndex, UserIndex, 0, "J1" & UserList(UserIndex).Stats.UserSkills(Skill) & "," & Skill)

Exit Sub
fallo:
Call LogError("enviaUnskills " & Err.number & " D: " & Err.Description)

End Sub
Sub EnviarSkills(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim i As Integer
Dim cad$
For i = 1 To NUMSKILLS
   cad$ = cad$ & UserList(UserIndex).Stats.UserSkills(i) & ","
Next
cad$ = cad$ + str$(UserList(UserIndex).Stats.SkillPts)
SendData2 ToIndex, UserIndex, 0, 57, cad$

Exit Sub
fallo:
Call LogError("enviarsubirskills " & Err.number & " D: " & Err.Description)

End Sub

Sub EnviarFama(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim cad$
cad$ = cad$ & UserList(UserIndex).Reputacion.AsesinoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BandidoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BurguesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.LadronesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.NobleRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.PlebeRep & ","

Dim l As Long
l = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
        (-UserList(UserIndex).Reputacion.BandidoRep) + _
        UserList(UserIndex).Reputacion.BurguesRep + _
        (-UserList(UserIndex).Reputacion.LadronesRep) + _
        UserList(UserIndex).Reputacion.NobleRep + _
        UserList(UserIndex).Reputacion.PlebeRep
l = l / 6

UserList(UserIndex).Reputacion.Promedio = l

cad$ = cad$ & UserList(UserIndex).Reputacion.Promedio & ","
'cad$ = cad$ & UserList(UserIndex).clase
SendData2 ToIndex, UserIndex, 0, 47, cad$
Exit Sub
fallo:
Call LogError("enviarfama " & Err.number & " D: " & Err.Description)


End Sub

Sub EnviarAtrib(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim i As Integer
Dim cad$
For i = 1 To NUMATRIBUTOS
  cad$ = cad$ & UserList(UserIndex).Stats.UserAtributos(i) & ","
Next
Call SendData2(ToIndex, UserIndex, 0, 36, cad$)
Exit Sub
fallo:
Call LogError("enviaratrib " & Err.number & " D: " & Err.Description)

End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer)

On Error GoTo ErrorHandler
   'pluto:6.5
   If UserList(UserIndex).Char.CharIndex = 0 Then Exit Sub
        
        'Debug.Print (UserList(UserIndex).Name)
        'Exit Sub
   ' End If
        
        
        CharList(UserList(UserIndex).Char.CharIndex) = 0
        
        If UserList(UserIndex).Char.CharIndex = LastChar Then
                Do Until CharList(LastChar) > 0
                        LastChar = LastChar - 1
                        If LastChar = 0 Then Exit Do
                Loop
        End If
        
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        
        'Le mandamos el mensaje para que borre el personaje a los clientes que este en el mismo mapa
        Call SendData(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "BP" & UserList(UserIndex).Char.CharIndex)
        
        UserList(UserIndex).Char.CharIndex = 0
        
   ' NumChars = NumChars - 1
        
        Exit Sub
        
ErrorHandler:
                Call LogError("Error en EraseUserchar")

End Sub
Sub EraseUserCharMismoIndex(ByVal UserIndex As Integer)

On Error GoTo ErrorHandler
Dim Fallito As Byte
   'pluto:6.5
   If UserList(UserIndex).Char.CharIndex = 0 Then Exit Sub
        Fallito = 1
        'Debug.Print (UserList(UserIndex).Name)
        'Exit Sub
   ' End If
        
        
        'CharList(UserList(UserIndex).Char.CharIndex) = 0
        
        'If UserList(UserIndex).Char.CharIndex = LastChar Then
         '   Do Until CharList(LastChar) > 0
          '      LastChar = LastChar - 1
           '     If LastChar = 0 Then Exit Do
                'Loop
 '   End If
        
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        Fallito = 2
        'Le mandamos el mensaje para que borre el personaje a los clientes que este en el mismo mapa
        Call SendData(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "BP" & UserList(UserIndex).Char.CharIndex)
        Fallito = 3
        'UserList(UserIndex).Char.CharIndex = 0
        
        'NumChars = NumChars - 1
        Fallito = 4
        Exit Sub
        
ErrorHandler:
                Call LogError("Error en EraseUsercharMismoIndex Name: " & UserList(UserIndex).Name & "Pos: " & UserList(UserIndex).Pos.Map & " X: " & UserList(UserIndex).Pos.X & " Y: " & UserList(UserIndex).Pos.Y & " F: " & Fallito & " Charindex: " & UserList(UserIndex).Char.CharIndex & " " & Err.Description)

End Sub
Sub MakeUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo fallo

Dim CharIndex As Integer

If InMapBounds(Map, X, Y) Then

           'If needed make a new character in list
           If UserList(UserIndex).Char.CharIndex = 0 Then
                   CharIndex = NextOpenCharIndex
                   UserList(UserIndex).Char.CharIndex = CharIndex
                   CharList(CharIndex) = UserIndex
           End If
           
           'Place character on map
           MapData(Map, X, Y).UserIndex = UserIndex
           
           'Send make character command to clients
           Dim klan$
           klan$ = UserList(UserIndex).GuildInfo.GuildName
           Dim bCr As Byte
           If (Criminal(UserIndex)) Then bCr = 1
           'If (UserList(UserIndex).Faccion.ArmadaReal = 2) Then bCr = 2
           
           'bCr = Criminal(UserIndex)
                'If klan$ <> "" Then
                   '[GAU] Agregamo las Botas
                        'Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr & "," & UserList(UserIndex).flags.Privilegios & "," & UserList(UserIndex).Char.Botas)
                'Else
                        'Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios & "," & UserList(UserIndex).Char.Botas)
                'End If
         
   'pluto:7.0

   Dim EsGoblin As Byte

   Dim rReal As Byte
   If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
        If UserList(UserIndex).raza = "Goblin" Then EsGoblin = 1 Else EsGoblin = 0
        rReal = 1
        Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & klan$ & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios & "," & UserList(UserIndex).Char.Botas & "," & UserList(UserIndex).flags.partyNum & "," & UserList(UserIndex).flags.DragCredito4 & "," & EsGoblin & "," & rReal)
   ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
        If UserList(UserIndex).raza = "Goblin" Then EsGoblin = 1 Else EsGoblin = 0
        rReal = 2
        Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & klan$ & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios & "," & UserList(UserIndex).Char.Botas & "," & UserList(UserIndex).flags.partyNum & "," & UserList(UserIndex).flags.DragCredito4 & "," & EsGoblin & "," & rReal)
    Else
        If UserList(UserIndex).raza = "Goblin" Then EsGoblin = 1 Else EsGoblin = 0
        rReal = 0
        Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & klan$ & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios & "," & UserList(UserIndex).Char.Botas & "," & UserList(UserIndex).flags.partyNum & "," & UserList(UserIndex).flags.DragCredito4 & "," & EsGoblin & "," & rReal)
    End If

   
   
  'Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & klan$ & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios & "," & UserList(UserIndex).Char.Botas & "," & UserList(UserIndex).flags.partyNum & "," & UserList(UserIndex).flags.DragCredito4 & "," & EsGoblin & "," & Caos)
 
End If
Exit Sub
fallo:
Call LogError("makeuserchar " & Err.number & " D: " & Err.Description)

End Sub
Sub CheckUserLevel(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim Pts As Integer
Dim AumentoHIT As Integer
Dim AumentoST As Integer
Dim AumentoMANA As Integer
Dim AumentoHP As Integer
Dim WasNewbie As Boolean
Call SendUserStatsEXP(UserIndex)
'¿Alcanzo el maximo nivel?



If UserList(UserIndex).Stats.ELV = STAT_MAXELV Then '1
        UserList(UserIndex).Stats.exp = 0
        UserList(UserIndex).Stats.Elu = 0
        Exit Sub
End If '1

WasNewbie = EsNewbie(UserIndex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
If UserList(UserIndex).Stats.exp >= UserList(UserIndex).Stats.Elu Then '2
         
        
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_NIVEL)
Call SendData(ToIndex, UserIndex, 0, "||¡Has subido de nivel!" & "´" & FontTypeNames.FONTTYPE_info)
'nati: agrego MensajesQuest!
'Call MensajesQuest(UserIndex)
        
                'pluto:2.15--------SUBE LEVEL NIÑO---------------
                If UserList(UserIndex).Bebe > 0 Then '3
                UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
                UserList(UserIndex).Stats.exp = 0
                UserList(UserIndex).Stats.Elu = UserList(UserIndex).Stats.Elu * 2.5
                AumentoHP = RandomNumber(2, UserList(UserIndex).Stats.UserAtributos(Constitucion) / 2) + Int(UserList(UserIndex).Bebe / 3)
                AumentoST = Int(UserList(UserIndex).Bebe / 2)
                AumentoHIT = Int(UserList(UserIndex).Bebe / 3)
                 
                Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
                Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
                Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
                Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
                Call SendData(ToIndex, UserIndex, 0, "||Mejora de tus Atributos: " & "´" & FontTypeNames.FONTTYPE_info)
                Dim Incre As Byte
                Dim probi As Byte
                Dim n As Byte

                For n = 1 To 5
                probi = RandomNumber(1, 30) + UserList(UserIndex).Bebe
                Incre = 0
                If probi > 6 Then Incre = 1
                If probi > 22 Then Incre = RandomNumber(1, CInt(UserList(UserIndex).Bebe / 5))
                If probi > 30 Then Incre = 2
                UserList(UserIndex).Stats.UserAtributosBackUP(n) = UserList(UserIndex).Stats.UserAtributosBackUP(n) + Incre
                UserList(UserIndex).Stats.UserAtributos(n) = UserList(UserIndex).Stats.UserAtributosBackUP(n)
                If n = 1 Then Call SendData(ToIndex, UserIndex, 0, "||Fuerza: " & Incre & "´" & FontTypeNames.FONTTYPE_info)
                If n = 2 Then Call SendData(ToIndex, UserIndex, 0, "||Agilidad: " & Incre & "´" & FontTypeNames.FONTTYPE_info)
                If n = 3 Then Call SendData(ToIndex, UserIndex, 0, "||Inteligencia: " & Incre & "´" & FontTypeNames.FONTTYPE_info)
                If n = 4 Then Call SendData(ToIndex, UserIndex, 0, "||Carisma: " & Incre & "´" & FontTypeNames.FONTTYPE_info)
                If n = 5 Then Call SendData(ToIndex, UserIndex, 0, "||Constitución: " & Incre & "´" & FontTypeNames.FONTTYPE_info)

                Next
                UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + (UserList(UserIndex).Bebe * 2)
                Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & UserList(UserIndex).Bebe * 2 & " SkillPoints." & "´" & FontTypeNames.FONTTYPE_info)
                'pluto:6.0A
                UserList(UserIndex).Stats.Fama = UserList(UserIndex).Stats.Fama + 25

                                        '------------deja de ser niño-----------------
                        If UserList(UserIndex).Stats.ELV >= 5 Then '4
                        
                        Call DarCuerpoYCabeza(UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).raza, UserList(UserIndex).Genero)
                        UserList(UserIndex).OrigChar = UserList(UserIndex).Char
                        UserList(UserIndex).Char.WeaponAnim = NingunArma
                        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                        UserList(UserIndex).Char.CascoAnim = NingunCasco
                        UserList(UserIndex).Stats.MET = 1
                        UserList(UserIndex).Stats.FIT = 1
                        'clase aleatoria
                        Dim Oiu As Byte
                        Dim UserClase As String
                        Oiu = RandomNumber(1, 19)
                        UserList(UserIndex).clase = ListaClases(Oiu)
                        UserClase = UserList(UserIndex).clase
                        'pluto:2.15
                        Dim ains As Integer

                         
                                If UserClase = "Mago" Or UserClase = "Clerigo" Or _
                                UserClase = "Druida" Or UserClase = "Bardo" Or _
                                UserClase = "Pirata" Or UserClase = "Asesino" Then '5
                                ains = 18
                                If UserList(UserIndex).raza = "Gnomo" Then ains = 3 + 18
                                If UserList(UserIndex).raza = "Humano" Then ains = 1 + 18
                                If UserList(UserIndex).raza = "Elfo" Then ains = 2 + 18

                                        If UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) < ains Then
                                        UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia) = ains
                                        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributosBackUP(Inteligencia)
                                        End If
                                End If 'clase mago or cler.... '5

                                If UserClase = "Guerrero" Or UserClase = "Cazador" Or _
                                UserClase = "Arquero" Or UserClase = "Paladin" Then '6
                                ains = 18
                                If UserList(UserIndex).raza = "Orco" Or UserList(UserIndex).raza = "Enano" Then ains = 3 + 18
                                If UserList(UserIndex).raza = "Humano" Then ains = 2 + 18
                                If UserList(UserIndex).raza = "Elfo" Or UserList(UserIndex).raza = "Elfo Oscuro" Or UserList(UserIndex).raza = "Gnomo" Then ains = 1 + 18
                                If UserList(UserIndex).raza = "Vampiro" Then ains = 2 + 18

        
                                        If UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) < ains Then
                                        UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion) = ains
                                        UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributosBackUP(Constitucion)
                                        End If
                                End If 'clase guerrero... '6

                                If UserClase = "Ladron" Or UserClase = "Bandido" Then '7
                                ains = 18
                                If UserList(UserIndex).raza = "Gnomo" Then ains = 3 + 18
                                If UserList(UserIndex).raza = "Elfo" Or UserList(UserIndex).raza = "Elfo Oscuro" Or UserList(UserIndex).raza = "Humano" Or UserList(UserIndex).raza = "Vampiro" Then ains = 2 + 18
 
                                If UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) < ains Then
                                UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) = ains
                                UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad)
                                End If
                                End If 'clase ladron '7

                                '---------------
                        'Mana
                        Dim MiInt As Byte
                                If UserClase = "Mago" Then '8
                                MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Inteligencia)) / 3
                                UserList(UserIndex).Stats.MaxMAN = 100 + MiInt
                                UserList(UserIndex).Stats.MinMAN = 100 + MiInt
                                ElseIf UserClase = "Clerigo" Or UserClase = "Druida" _
                                 Or UserClase = "Bardo" Or UserClase = "Pirata" Or UserClase = "Asesino" Then
                                MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Inteligencia)) / 4
                                UserList(UserIndex).Stats.MaxMAN = 50
                                UserList(UserIndex).Stats.MinMAN = 50
                                Else
                                UserList(UserIndex).Stats.MaxMAN = 0
                                UserList(UserIndex).Stats.MinMAN = 0
                                End If '8

                                If UserClase = "Mago" Or UserClase = "Clerigo" Or _
                                UserClase = "Druida" Or UserClase = "Bardo" Or _
                                UserClase = "Pirata" Or UserClase = "Asesino" Then '9
                                UserList(UserIndex).Stats.UserHechizos(1) = 2
                
                                End If '9
UserList(UserIndex).Stats.exp = 0
UserList(UserIndex).Stats.Elu = 200
UserList(UserIndex).Stats.ELV = 0
UserList(UserIndex).Bebe = 0
Call SendData(ToIndex, UserIndex, 0, "!! Ya eres adulto y has decidido que tu futuro es llegar a ser el mejor " & UserClase & " de estas tierras.")
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)

'----------fin deja ser niño---------------
Else
GoTo yap
End If '4

'GoTo yap
End If '3
'-----------------------------------------------

   ' If UserList(UserIndex).Stats.ELV = 1 Then
          'Pts = 10
          
        'Else
          Pts = 5
        'End If
        
        'pluto:2.17
        If UserList(UserIndex).clase = "Minero" Or UserList(UserIndex).clase = "Leñador" Or UserList(UserIndex).clase = "Pescador" Or UserList(UserIndex).clase = "Herrero" Or UserList(UserIndex).clase = "Ermitaño" Or UserList(UserIndex).clase = "Carpintero" Or UserList(UserIndex).clase = "Domador" Then Pts = Pts * 2
        If UserList(UserIndex).Remort > 0 Then Pts = Pts + 1
        '-------------------------------
        UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts
        
        Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & Pts & " skillpoints." & "´" & FontTypeNames.FONTTYPE_info)
'pluto:6.0A
UserList(UserIndex).Stats.Fama = UserList(UserIndex).Stats.Fama + 25
        
        UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
        
        UserList(UserIndex).Stats.exp = 0
        
        'sacar del newbie
        If Not EsNewbie(UserIndex) And WasNewbie Then
        If UserList(UserIndex).Pos.Map = 37 Then Call WarpUserChar(UserIndex, Nix.Map, Nix.X, Nix.Y, True)
        Call WarpUserChar(UserIndex, 34, 34, 37, True)
        Call SendData(ToIndex, UserIndex, 0, "!! Has dejado de ser Newbie y ya no estás protegido por los Dioses. Todos los objetos de Newbie serán borrados de tu inventario y a partir de ahora todos los objetos que consigas se te caerán al morir (incluido el oro). Cuando tengas 100.000 oros en tu casillero el oro ya no se caerá mientrás tanto deberás ir soltándolos en el banco para no perderlos. Suerte Aodraguero!! ")
        End If
        
        'pluto:6.5--
        'End If
        '-------------
        If Not EsNewbie(UserIndex) And WasNewbie Then Call QuitarNewbieObj(UserIndex)
        
        If UserList(UserIndex).Stats.ELV < 11 Then
                UserList(UserIndex).Stats.Elu = UserList(UserIndex).Stats.Elu * 1.5
        ElseIf UserList(UserIndex).Stats.ELV < 25 Then
                UserList(UserIndex).Stats.Elu = UserList(UserIndex).Stats.Elu * 1.3
        Else
                UserList(UserIndex).Stats.Elu = UserList(UserIndex).Stats.Elu * 1.2
        End If


'pluto:6.5
 Dim Elixir As Byte
  Elixir = UserList(UserIndex).flags.Elixir
 'pluto:6.9
 Dim ManaEquipado As Integer
 ManaEquipado = ObjetosConMana(UserIndex)
 '------------------------

'nati: nuevo diseño de subida de vida.
Select Case UserList(UserIndex).clase

                Case "Guerrero"
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 11) + Elixir
                Else
                AumentoHP = RandomNumber(6, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 12) + Elixir
                Else
                AumentoHP = RandomNumber(8, 12)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(9, 12) + Elixir
                Else
                AumentoHP = RandomNumber(9, 12)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(9, 13) + Elixir
                Else
                AumentoHP = RandomNumber(9, 13)
                End If
                End If
                'nico bug hijo
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 22 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(9, 14) + Elixir
                Else
                AumentoHP = RandomNumber(9, 14)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 23 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(9, 14) + Elixir
                Else
                AumentoHP = RandomNumber(9, 14)
                End If
                End If
                'fin bug hijo guerrero
                
                
                        'pluto:6.5-----------------------------
                        'If Elixir = 10 Then
                        'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 1
                        'Else
                        'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 1
                        'End If
                        '-------------------------------------
                        
                        AumentoST = 15
                        AumentoHIT = 3
                        'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWnomagico" & UserList(UserIndex).Stats.ELV)
                        
                        If (UserList(UserIndex).Remort = 1) Then
                        'pluto:6.5-----------------------------
                        If Elixir = 10 Then
                        AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 4
                        Else
                        AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 4
                        End If
                        '----------------------------------------------------
                                        
                                        AumentoHIT = 4
                                        AumentoST = 25
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 800)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1500)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 200)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 200)
                                        GoTo yap
                        End If
                        '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
                        
                        '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
                        
                        '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 120)
                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 120)
                
                Case "Cazador"
                   'pluto:6.5-----------------------------
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 11) + Elixir
                Else
                AumentoHP = RandomNumber(8, 11)
                End If
                End If
            'nico bug hijo
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 22 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 12) + Elixir
                Else
                AumentoHP = RandomNumber(8, 12)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 23 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 13) + Elixir
                Else
                AumentoHP = RandomNumber(8, 13)
                End If
                End If
              'fin bug hijo cazador
                
                        'If Elixir = 10 Then
                   'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
                        'Else
                        'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
                   'End If
                   '------------------------------------
                   
                   AumentoST = 15
                   AumentoHIT = 3
                         'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWnomagico" & UserList(UserIndex).Stats.ELV)

                   If (UserList(UserIndex).Remort = 1) Then
                                        
                                        'pluto:6.5-----------------------------
                                        If Elixir = 10 Then
                                        AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 2
                                        Else
                                        AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 2
                                        End If
                                        '----------------------------------------------------
                                        
                                        AumentoST = 25
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 650)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1000)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 120)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 120)
                                        GoTo yap
                        End If
                        '¿?¿?¿?¿?¿?¿?¿ HitPoints ¿?¿?¿?¿?¿?¿?¿
                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
                        
                        '¿?¿?¿?¿?¿?¿?¿ Stamina ¿?¿?¿?¿?¿?¿?¿
                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
                        
                        '¿?¿?¿?¿?¿?¿?¿ Golpe ¿?¿?¿?¿?¿?¿?¿
                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
                        
          'pluto:2.17
  
           Case "Arquero"
                
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(3, 7) + Elixir
                Else
                AumentoHP = RandomNumber(3, 7)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(4, 7) + Elixir
                Else
                AumentoHP = RandomNumber(4, 7)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(4, 8) + Elixir
                Else
                AumentoHP = RandomNumber(4, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 8) + Elixir
                Else
                AumentoHP = RandomNumber(5, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                                'nico bug hijo
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 22 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 23 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                'fin bug hijo arquero
                        'pluto:6.5-----------------------------
                        'If Elixir = 10 Then
                        'AumentoHP = UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2
                        'Else
                        'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                        'End If
                        '---------------------------------------------
                        
                        AumentoST = 15
                        AumentoHIT = 2
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWnomagico" & UserList(UserIndex).Stats.ELV)

                        If (UserList(UserIndex).Remort = 1) Then
                        
                        'pluto:6.5-----------------------------
                        If Elixir = 10 Then
                        AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                        Else
                        AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                        End If
                        '-------------------------------------------------------
        
        
        'pluto:6.0A------------------
                        If UserList(UserIndex).Stats.ELV Mod (2) = 0 Then
                        AumentoHIT = 3
                        Else
                        AumentoHIT = 3
                        End If
                        '-------------------------
                                   AumentoST = 20
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 500)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1500)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 130)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 130)
                                        GoTo yap
                        End If

                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        
                           'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT

          
          
                Case "Pirata"
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 11) + Elixir
                Else
                AumentoHP = RandomNumber(8, 11)
                End If
                End If
                        'pluto:6.5-----------------------------
                        'If Elixir = 10 Then
                        'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
                        'Else
                        'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
                        'End If
                        '-------------------------------------------------
                        
                        
                        AumentoST = 15
                        AumentoHIT = 3
                        AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWsemi" & UserList(UserIndex).Stats.ELV)

                        If (UserList(UserIndex).Remort = 1) Then
                        
                        'pluto:6.5-----------------------------
                        If Elixir = 10 Then
                        AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 1
                        Else
                        AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 1
                        End If
                        '---------------------------------------
                        
                                        AumentoST = 20
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 750)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1300)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 120)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 120)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 3000 + ManaEquipado)
                                        GoTo yap
                        End If
                        'If CInt(UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1)) > STAT_MAXMAN Then AumentoMANA = 0

                        'HP
                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
                        'Mana
                        If CInt((UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 2)) + 0) < STAT_MAXMAN Then AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 2000 + ManaEquipado

                        'STA
                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
                        
                        'Golpe
                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
                        
                Case "Paladin"
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 11) + Elixir
                Else
                AumentoHP = RandomNumber(8, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 22 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 12) + Elixir
                Else
                AumentoHP = RandomNumber(8, 12)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 23 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 13) + Elixir
                Else
                AumentoHP = RandomNumber(8, 13)
                End If
                End If
                        'pluto:6.5-----------------------------
                        'If Elixir = 10 Then
                        'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
                        'Else
                        'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
                        'End If
                        '-------------------------------------------------
                        
                        
                        AumentoST = 15
                        AumentoHIT = 3
                        AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWsemi" & UserList(UserIndex).Stats.ELV)

                        If (UserList(UserIndex).Remort = 1) Then
                        
                        'pluto:6.5-----------------------------
                        If Elixir = 10 Then
                        AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 1
                        Else
                        AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero + 1
                        End If
                        '---------------------------------------
                        
                                        AumentoST = 20
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 650)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1300)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 120)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 120)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 3000 + ManaEquipado)
                                        GoTo yap
                        End If
                        'If CInt(UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1)) > STAT_MAXMAN Then AumentoMANA = 0

                        'HP
                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
                        'Mana
                        If CInt((UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 2)) + 0) < STAT_MAXMAN Then AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 2000 + ManaEquipado

                        'STA
                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
                        
                        'Golpe
                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
                                                
                Case "Ladron"
                        
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 11) + Elixir
                Else
                AumentoHP = RandomNumber(8, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 12) + Elixir
                Else
                AumentoHP = RandomNumber(8, 12)
                End If
                End If
                        'pluto:6.5-----------------------------
                        'If Elixir = 10 Then
                        'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                        'Else
                        'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                        'End If
                        '-------------------------------------------
                        
                        AumentoST = 15 + AdicionalSTLadron
                        AumentoHIT = 1
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWnomagico" & UserList(UserIndex).Stats.ELV)

                        If (UserList(UserIndex).Remort = 1) Then
                                        'AumentoHP = RandomNumber(6, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                        'pluto:6.5-----------------------------
                        If Elixir = 10 Then
                        AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                        Else
                        AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                        End If
                        '-------------------------------------
                                        AumentoST = 20
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 625)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1500)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 110)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 110)
                                        GoTo yap
                        End If
                        
                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
                        
                Case "Mago"
                
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(3, 7) + Elixir
                Else
                AumentoHP = RandomNumber(3, 7)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(4, 7) + Elixir
                Else
                AumentoHP = RandomNumber(4, 7)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(4, 8) + Elixir
                Else
                AumentoHP = RandomNumber(4, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 8) + Elixir
                Else
                AumentoHP = RandomNumber(5, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                                        'pluto:6.5-----------------------------
                'If Elixir = 10 Then
                 'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 1
                 'Else
                 'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 1
                'End If
                   '-----------------------------------------------------------------------
                                                
                        If AumentoHP < 1 Then AumentoHP = 4
                        AumentoST = 15 - AdicionalSTLadron / 2
                        If AumentoST < 1 Then AumentoST = 5
                        AumentoHIT = 1
                        AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWmagico" & UserList(UserIndex).Stats.ELV)

                        If (UserList(UserIndex).Remort = 1) Then
                                         'pluto:6.5-----------------------------
                                        If Elixir = 10 Then
                                        AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                                        Else
                                        AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                                        End If
                                        '-----------------------------------------------
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 475)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1500)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 99)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 99)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 5000 + ManaEquipado)
                                        GoTo yap
                        End If
                        
                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Mana
                        If CInt((UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 2) * 3) + 107) < STAT_MAXMAN Then AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 2000 + ManaEquipado
                                                'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
                Case "Leñador"
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 11) + Elixir
                Else
                AumentoHP = RandomNumber(8, 11)
                End If
                End If
                        'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)), UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                        AumentoST = 15 + AdicionalSTLeñador
                        AumentoHIT = 2
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWcurro" & UserList(UserIndex).Stats.ELV)

                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
                Case "Minero"
                                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 11) + Elixir
                Else
                AumentoHP = RandomNumber(8, 11)
                End If
                End If
                        'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)), UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                        AumentoST = 15 + AdicionalSTMinero
                        AumentoHIT = 2
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWcurro" & UserList(UserIndex).Stats.ELV)

                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
                Case "Pescador"
                                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 11) + Elixir
                Else
                AumentoHP = RandomNumber(8, 11)
                End If
                End If
                        'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)), UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                        AumentoST = 15 + AdicionalSTPescador
                        AumentoHIT = 1
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWcurro" & UserList(UserIndex).Stats.ELV)

                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
                                   
                Case "Clerigo"
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(4, 8) + Elixir
                Else
                AumentoHP = RandomNumber(4, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 8) + Elixir
                Else
                AumentoHP = RandomNumber(5, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                ''''''''''bug hijos nico clero
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 22 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 11) + Elixir
                Else
                AumentoHP = RandomNumber(8, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 23 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 12) + Elixir
                Else
                AumentoHP = RandomNumber(8, 12)
                End If
                End If
                'bug hijos clero fin
                'pluto:6.5-----------------------------
                'If Elixir = 10 Then
                'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                'Else
                'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                'End If
                '--------------------------------
                        
                        AumentoST = 15
                        AumentoHIT = 2
                        AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
  'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWsemi" & UserList(UserIndex).Stats.ELV)

                         If (UserList(UserIndex).Remort = 1) Then
                                        AumentoST = 20
                        'pluto:6.0A------------------
                        If UserList(UserIndex).Stats.ELV Mod (2) = 0 Then
                        AumentoHIT = 3
                        Else
                        AumentoHIT = 2
                        End If
                        '-------------------------
                'pluto:6.5-----------------------------
                If Elixir = 10 Then
                AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 2
                Else
                AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 2
                End If
                '----------------------------------
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 600)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1500)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 110)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 110)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 4000)
                                        GoTo yap
                        End If
                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Mana
                         If CInt((UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 2) * 2) + 50) < STAT_MAXMAN Then AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 3000
                   
                        'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT

                Case "Druida"
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(4, 8) + Elixir
                Else
                AumentoHP = RandomNumber(4, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 8) + Elixir
                Else
                AumentoHP = RandomNumber(5, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 22 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 23 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 12) + Elixir
                Else
                AumentoHP = RandomNumber(7, 12)
                End If
                End If
                'pluto:6.5-----------------------------
                'If Elixir = 10 Then
                'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                'Else
                'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                'End If
                 '---------------------------------------
                        AumentoST = 15
                        AumentoHIT = 2
                        AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWmagico" & UserList(UserIndex).Stats.ELV)

                         If (UserList(UserIndex).Remort = 1) Then
                'pluto:6.5---------------------------------------------
                If Elixir = 10 Then
                AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 2
                Else
                AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 2
                End If
                '-----------------------------------------------------
                                        
                                        
                                        AumentoST = 20
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 525)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1000)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 99)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 99)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 4000)
                                        GoTo yap
                        End If
                        If CInt(UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 1) * 2) > STAT_MAXMAN Then AumentoMANA = 0
                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Mana
                          If CInt((UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 2) * 2) + 50) < STAT_MAXMAN Then AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 3000
                  
                                                'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT

                Case "Asesino"
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 8) + Elixir
                Else
                AumentoHP = RandomNumber(5, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 8) + Elixir
                Else
                AumentoHP = RandomNumber(6, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 22 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 12) + Elixir
                Else
                AumentoHP = RandomNumber(7, 12)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 23 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 12) + Elixir
                Else
                AumentoHP = RandomNumber(8, 12)
                End If
                End If
                                        'pluto:6.5-----------------------------
                'If Elixir = 10 Then
                ' AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                'Else
                'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                'End If
                        
                        
                        AumentoST = 15
                        AumentoHIT = 3
                        AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                          'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWsemi" & UserList(UserIndex).Stats.ELV)

                         If (UserList(UserIndex).Remort = 1) Then
                'pluto:6.5-----------------------------
                If Elixir = 10 Then
                AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 2
                Else
                AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 2
                 End If
                   '------------------------------------------
                                        
                                        
                                        
                                        AumentoST = 20
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 625)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1500)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 120)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 120)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 4000)
                                        GoTo yap
                        End If

                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Mana
                         If CInt((UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 2)) + 50) < STAT_MAXMAN Then AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 3000
                   
                                                'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
                        
                Case "Bardo"
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 8) + Elixir
                Else
                AumentoHP = RandomNumber(5, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 8) + Elixir
                Else
                AumentoHP = RandomNumber(6, 8)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 22 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 12) + Elixir
                Else
                AumentoHP = RandomNumber(7, 12)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 23 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 12) + Elixir
                Else
                AumentoHP = RandomNumber(8, 12)
                End If
                End If
                        
                                        'pluto:6.5-----------------------------
                'If Elixir = 10 Then
                'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                'Else
                'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                'End If
                        
                        
                        AumentoST = 15
                        AumentoHIT = 2
                        AumentoMANA = CInt(1.5 * UserList(UserIndex).Stats.UserAtributos(Inteligencia))
                           'pluto:6.0A
                        Call SendData(ToIndex, UserIndex, 0, "AWsemi" & UserList(UserIndex).Stats.ELV)

                         If (UserList(UserIndex).Remort = 1) Then
                 'pluto:6.5-----------------------------
                If Elixir = 10 Then
                AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 2
                Else
                AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 2
                End If
                '----------------------------------------
                                        AumentoST = 18
                                        AumentoHIT = 3
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 625)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1500)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 120)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 120)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 4000)
                                        GoTo yap
                        End If

                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Mana
                           If CInt((UserList(UserIndex).Stats.UserAtributos(Inteligencia) * (UserList(UserIndex).Stats.ELV - 2) * 1.5) + 50) < STAT_MAXMAN Then AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, 3000

                           'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
                Case Else
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 16 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(5, 9) + Elixir
                Else
                AumentoHP = RandomNumber(5, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 17 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 9) + Elixir
                Else
                AumentoHP = RandomNumber(6, 9)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 18 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(6, 10) + Elixir
                Else
                AumentoHP = RandomNumber(6, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 19 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 10) + Elixir
                Else
                AumentoHP = RandomNumber(7, 10)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 20 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(7, 11) + Elixir
                Else
                AumentoHP = RandomNumber(7, 11)
                End If
                End If
                If UserList(UserIndex).Stats.UserAtributos(Constitucion) = 21 Then
                If Elixir = 10 Then
                AumentoHP = RandomNumber(8, 11) + Elixir
                Else
                AumentoHP = RandomNumber(8, 11)
                End If
                End If
            
                'pluto:6.5-----------------------------
                'If Elixir = 10 Then
                'AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                'Else
                'AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 6 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
                'End If
                 '--------------------------------------------
                        
                        AumentoST = 15
                        AumentoHIT = 2
                        'pluto:6.0A------------------------
                   If UserList(UserIndex).clase = "Bandido" Or UserList(UserIndex).clase = "Domador" Then
                        Call SendData(ToIndex, UserIndex, 0, "AWnomagico" & UserList(UserIndex).Stats.ELV)
                        Else
                        Call SendData(ToIndex, UserIndex, 0, "AWcurro" & UserList(UserIndex).Stats.ELV)
                        End If
                        '-----------------------------------
                        If (UserList(UserIndex).Remort = 1) Then
                'pluto:6.5-----------------------------
                If Elixir = 10 Then
                AumentoHP = (UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                Else
                AumentoHP = RandomNumber((UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - 4 + (UserList(UserIndex).Stats.UserAtributos(Constitucion) Mod (2)) + Elixir, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + 1
                End If
                '-----------------------------------------
                                        AumentoST = 18
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, 600)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, 1000)
                                        Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, 99)
                                        Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, 99)
                        GoTo yap
                End If
                        
                        'HP
                        AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
                        'STA
                        AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
                        'Golpe
                        AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
                        AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        End Select
        
  
                        
 
yap:
        If AumentoHP > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & "´" & FontTypeNames.FONTTYPE_info
        'pluto:6.5
        If Elixir > 0 Then
                If Elixir = 10 Then
                SendData ToIndex, UserIndex, 0, "||Has ganado el Máximo de puntos de vida gracias al Elixir de Vida." & "´" & FontTypeNames.FONTTYPE_info
                Else
                SendData ToIndex, UserIndex, 0, "||De los " & AumentoHP & " Puntos de vida " & Elixir & " han sido gracias al Elixir de Vida." & "´" & FontTypeNames.FONTTYPE_info
                End If
        Elixir = 0
        UserList(UserIndex).flags.Elixir = 0
        End If
        '-----------------------------
        
        If AumentoST > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoST & " puntos de vitalidad." & "´" & FontTypeNames.FONTTYPE_info
        If AumentoMANA > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoMANA & " puntos de magia." & "´" & FontTypeNames.FONTTYPE_info
        If AumentoHIT > 0 Then
                SendData ToIndex, UserIndex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & "´" & FontTypeNames.FONTTYPE_info
                SendData ToIndex, UserIndex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & "´" & FontTypeNames.FONTTYPE_info
        End If
        
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        Call EnviarSkills(UserIndex)
        'Call EnviarSubirNivel(UserIndex, Pts)
'[Tite]Party
If UserList(UserIndex).flags.party = True Then
        If partylist(UserList(UserIndex).flags.partyNum).reparto = 1 Then
                Call BalanceaPrivisLVL(UserList(UserIndex).flags.partyNum)
        End If
End If
Call SendData(toParty, UserIndex, 0, "DD27" & UserList(UserIndex).Name)

'[\Tite]
        senduserstatsbox UserIndex
        If Not Criminal(UserIndex) And UserList(UserIndex).Stats.ELV = 50 Then Call AgregarHechizoangel(UserIndex, 37)
        If Not Criminal(UserIndex) And UserList(UserIndex).Stats.ELV = 50 Then Call AgregarHechizoangel(UserIndex, 38)
If Criminal(UserIndex) And UserList(UserIndex).Stats.ELV = 50 Then Call AgregarHechizoangel(UserIndex, 53)
If Criminal(UserIndex) And UserList(UserIndex).Stats.ELV = 50 Then Call AgregarHechizoangel(UserIndex, 52)

        
End If


Exit Sub

errhandler:
Call LogError("Error CHECKUSERLEVEL --> " & Err.number & " D: " & Err.Description & "--> " & UserList(UserIndex).Name & " -- " & UserList(UserIndex).Stats.ELV)

        'LogError ("Error en la subrutina CheckUserLevel")
End Sub

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
PuedeAtravesarAgua = _
  UserList(UserIndex).flags.Navegando = 1 Or _
  UserList(UserIndex).flags.Vuela = 1 'Or _
  UserList(UserIndex).Flags.Angel = 1 Or _
  UserList(UserIndex).Flags.Demonio = 1
Exit Function
fallo:
Call LogError("puedeatravesaragua " & Err.number & " D: " & Err.Description)

End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As Byte)

On Error GoTo fallo
'¿Tiene un indece valido?
If UserIndex <= 0 Then
Call CloseSocket(UserIndex)
Exit Sub
End If



'pluto:2.17
If UserList(UserIndex).Char.Heading <> nHeading Then
UserList(UserIndex).Char.Heading = nHeading
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
End If
'-------------------

Dim nPos As WorldPos
Dim AdminHide As Integer

'pluto:2.8.0
If UserList(UserIndex).Pos.Map <> 192 Then GoTo ppp 'dragfutbol
If nHeading = 4 Then
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y).NpcIndex > 0 Then
If Npclist(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y).NpcIndex).NPCtype = 21 Then
'Call MoveNPCChar(MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y + 1).NpcIndex, nHeading)
Call MoveNPCChar(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y).NpcIndex, nHeading)
End If
End If
End If

'---------------
ppp:
'Move
nPos = UserList(UserIndex).Pos
Call HeadtoPos(nHeading, nPos)

'Delzak) Triger auto-resu, hay que editar los mapas que tengan curas y ponerles triger 6 alrededor
'If MapData(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y).trigger = 6 And UserList(UserIndex).flags.Muerto = 1 Then
'Call RevivirUsuario(UserIndex)
'UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
'Call SendUserStatsVida(val(UserIndex))
'Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido restaurado!!" & "´" & FontTypeNames.FONTTYPE_info)
'End If



If LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(UserIndex)) Then
        AdminHide = 0
        If ((UserList(UserIndex).flags.AdminInvisible = 1) And (UserList(UserIndex).flags.Privilegios > 0)) Then AdminHide = 1
        Call SendData(ToMapButIndex, UserIndex, UserList(UserIndex).Pos.Map, "MP" & UserList(UserIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y & "," & AdminHide)
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        UserList(UserIndex).Pos = nPos
        UserList(UserIndex).Char.Heading = nHeading
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
        
Else
        'else correct user's pos
        Call SendData2(ToIndex, UserIndex, 0, 15, UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
End If


'----pluto:6.5 --------------controlamos si hay salida-------
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).TileExit.Map > 0 Then
Call ControlaSalidas(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
End If

If UserList(UserIndex).flags.Privilegios > 0 Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
'pluto:6.0A-----Eventos de las casas encantadas-------
If UserList(UserIndex).Pos.Map = 171 Or UserList(UserIndex).Pos.Map = 177 Then
Call VigilarEventosCasas(UserIndex)
End If
'-------------Eventos sala invocación---------------
If UserList(UserIndex).Pos.Map = mapi Then Call VigilarEventosInvocacion(UserIndex)

'--------------Eventos Trampas------------------------
If UserList(UserIndex).Pos.Map = 178 Or UserList(UserIndex).Pos.Map = 179 Then
Call VigilarEventosTrampas(UserIndex)
End If
'--------------------------------------------



Exit Sub
fallo:
Call LogError("moveuserchar " & UserIndex & " D: " & Err.Description & " name: " & UserList(UserIndex).Name & " mapa: " & UserList(UserIndex).Pos.Map & " X: " & UserList(UserIndex).Pos.X & " Y: " & UserList(UserIndex).Pos.Y)

End Sub

Sub ChangeUserInv(UserIndex As Integer, Slot As Byte, Object As UserOBJ)

On Error GoTo fallo
UserList(UserIndex).Invent.Object(Slot) = Object

If Object.ObjIndex > 0 Then
'pluto:6.0A
        Call SendData2(ToIndex, UserIndex, 0, 32, Slot & "#" & Object.ObjIndex & "#" & Object.Amount & "#" & Object.Equipped)
   
Else
        Call SendData2(ToIndex, UserIndex, 0, 32, Slot & "#" & "0") ' & "#" & "(None)" & "#" & "0" & "#" & "0")
End If

Exit Sub
fallo:
Call LogError("changeuserinv " & Err.number & " D: " & Err.Description)

End Sub

Function NextOpenCharIndex() As Integer
On Error GoTo fallo
Dim loopc As Integer
Dim n As Integer
For loopc = 1 To LastChar + 1
        If CharList(loopc) = 0 Then
        'pluto:6.6 ANTICLONES--------------------
        For n = 1 To MaxUsers
        If UserList(n).Char.CharIndex = loopc Then
        CharList(loopc) = UserList(n).Char.CharIndex
        GoTo otro
        End If
        Next
        '-----------------------------------------
                NextOpenCharIndex = loopc
                'NumChars = NumChars + 1
                If loopc > LastChar Then LastChar = loopc
                Exit Function
        End If
otro:
        
Next loopc

Exit Function
fallo:
Call LogError("nextopencharindex " & Err.number & " D: " & Err.Description)
End Function

Function NextOpenUser() As Integer
On Error GoTo fallo
Dim loopc As Integer
  
For loopc = 1 To MaxUsers + 1
  If loopc > MaxUsers Then Exit For
  'If (UserList(LoopC).ConnID = -1) Then Exit For
'pluto:2.22-------------
If (UserList(loopc).ConnID = -1 And UserList(loopc).flags.UserLogged = False) Then Exit For
'-------------------------
Next loopc
  
NextOpenUser = loopc
Exit Function
fallo:
Call LogError("nextopenuser " & Err.number & " D: " & Err.Description)

End Function
'pluto:2.9.0
Sub SendUserClase(ByVal UserIndex As Integer)
 On Error GoTo fallo
 'pluto:7.0 añado raza
   Call SendData2(ToIndex, UserIndex, 0, 93, UserList(UserIndex).clase & "," & UserList(UserIndex).raza)
Exit Sub
fallo:
Call LogError("senduserclase " & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.9.0
Sub SendUserMuertos(ByVal UserIndex As Integer)
 On Error GoTo fallo
 Call SendData(ToIndex, UserIndex, 0, "K2" & UserList(UserIndex).Faccion.CiudadanosMatados & "," & UserList(UserIndex).Faccion.CriminalesMatados & "," & UserList(UserIndex).Stats.NPCsMuertos)
Exit Sub
fallo:
Call LogError("sendusermuertos " & Err.number & " D: " & Err.Description)

End Sub

Sub senduserstatsbox(ByVal UserIndex As Integer)
  On Error GoTo fallo
   Call SendData2(ToIndex, UserIndex, 0, 23, UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.Elu & "," & UserList(UserIndex).Stats.exp)
Exit Sub
fallo:
Call LogError("senduserstatsbox " & Err.number & " D: " & Err.Description)

End Sub

'Delzak sistema premios

Sub SendUserPremios(ByVal UserIndex As Integer)
Dim n As Integer
Dim ELogros As String
  On Error GoTo fallo
  
  For n = 1 To 34
  ELogros = ELogros & UserList(UserIndex).Stats.PremioNPC(n) & ","
  Next
  
  
'Mata NPCS1
Call SendData(ToIndex, UserIndex, 0, "D1" & ELogros)

'Mata NPCS2
'Call SendData(ToIndex, UserIndex, 0, "D2" & UserList(UserIndex).Stats.PremioNPC.MataMedusas & "," & UserList(UserIndex).Stats.PremioNPC.MataCiclopes & "," & UserList(UserIndex).Stats.PremioNPC.MataPolares & "," & UserList(UserIndex).Stats.PremioNPC.MataDevastadores & "," & UserList(UserIndex).Stats.PremioNPC.MataGigantes & "," & UserList(UserIndex).Stats.PremioNPC.MataPiratas & "," & UserList(UserIndex).Stats.PremioNPC.MataUruks & "," & UserList(UserIndex).Stats.PremioNPC.MataDemonios & "," & UserList(UserIndex).Stats.PremioNPC.Matadevir & "," & UserList(UserIndex).Stats.PremioNPC.MataGollums & "," & UserList(UserIndex).Stats.PremioNPC.MataDragones & "," & UserList(UserIndex).Stats.PremioNPC.Mataettin & "," & UserList(UserIndex).Stats.PremioNPC.MataPuertas & "," & UserList(UserIndex).Stats.PremioNPC.MataReyes & "," & UserList(UserIndex).Stats.PremioNPC.MataDefensores & "," & UserList(UserIndex).Stats.PremioNPC.MataRaids & "," & UserList(UserIndex).Stats.PremioNPC.MataNavidad)

 
Exit Sub
fallo:
Call LogError("senduserpremios " & Err.number & " D: " & Err.Description)

End Sub
Sub SendUserRazaClase(ByVal UserIndex As Integer)
  On Error GoTo fallo

   Call SendData(ToIndex, UserIndex, 0, "J3" & UserList(UserIndex).raza & "," & UserList(UserIndex).clase)
Exit Sub
fallo:
Call LogError("senduserrazaclase " & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.3
Sub SendUserStatsVida(ByVal UserIndex As Integer)
   On Error GoTo fallo
        If UserList(UserIndex).Stats.MinHP < 0 Then UserList(UserIndex).Stats.MinHP = 0
        Call SendData2(ToIndex, UserIndex, 0, 24, UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP)
Exit Sub
fallo:
Call LogError("senduserstatsvida " & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.3
Sub SendUserStatsMana(ByVal UserIndex As Integer)
 On Error GoTo fallo
        Call SendData2(ToIndex, UserIndex, 0, 25, UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN)
Exit Sub
fallo:
Call LogError("senduserstatsmana " & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.3
Sub SendUserStatsEnergia(ByVal UserIndex As Integer)
On Error GoTo fallo
Call SendData2(ToIndex, UserIndex, 0, 26, UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta)
Exit Sub
fallo:
Call LogError("senduserstatsenergia " & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.3
Sub SendUserStatsOro(ByVal UserIndex As Integer)
On Error GoTo fallo
Call SendData2(ToIndex, UserIndex, 0, 27, UserList(UserIndex).Stats.GLD)
Exit Sub
fallo:
Call LogError("senduserstatsoro" & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.3
Sub SendUserStatsFama(ByVal UserIndex As Integer)
On Error GoTo fallo
Call SendData(ToIndex, UserIndex, 0, "H2" & UserList(UserIndex).Stats.Fama)

Exit Sub
fallo:
Call LogError("senduserstatsFama" & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.3
Sub SendUserStatsEXP(ByVal UserIndex As Integer)
On Error GoTo fallo
Call SendData2(ToIndex, UserIndex, 0, 28, UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.Elu & "," & UserList(UserIndex).Stats.exp & "," & UserList(UserIndex).Stats.Fama)
Exit Sub
fallo:
Call LogError("senduserstatsexp " & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.3
Sub SendUserStatsPeso(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).Stats.Peso < 0.001 Then UserList(UserIndex).Stats.Peso = 0
Call SendData2(ToIndex, UserIndex, 0, 29, Round(UserList(UserIndex).Stats.Peso, 3) & "#" & UserList(UserIndex).Stats.PesoMax)
Exit Sub
fallo:
Call LogError("senduserstatspeso " & Err.number & " D: " & Err.Description)

End Sub
Sub EnviarHambreYsed(ByVal UserIndex As Integer)
  On Error GoTo fallo
        Call SendData2(ToIndex, UserIndex, 0, 46, UserList(UserIndex).Stats.MaxAGU & "," & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MaxHam & "," & UserList(UserIndex).Stats.MinHam)
Exit Sub
fallo:
Call LogError("enviarhambreysed " & Err.number & " D: " & Err.Description)

End Sub
Sub SendUserMuertes(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
 On Error GoTo fallo
        Call SendData(ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, sendIndex, 0, "||Ciudadanos asesinados: " & UserList(UserIndex).Faccion.CiudadanosMatados & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, sendIndex, 0, "||Criminales asesinados: " & UserList(UserIndex).Faccion.CriminalesMatados & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, sendIndex, 0, "||Total Npcs matados: " & UserList(UserIndex).Stats.NPCsMuertos & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
fallo:
Call LogError("sendusermuertes " & Err.number & " D: " & Err.Description)

End Sub
Sub SendUserStatstxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
  On Error GoTo fallo
        Call SendData(ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, sendIndex, 0, "||Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.exp & "/" & UserList(UserIndex).Stats.Elu & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, sendIndex, 0, "||Clase: " & UserList(UserIndex).clase & "´" & FontTypeNames.FONTTYPE_info)
        'Call SendData(ToIndex, sendIndex, 0, "||Vitalidad: " & UserList(UserIndex).Stats.FIT & FONTTYPENAMES.FONTTYPE_INFO)
        Call SendData(ToIndex, sendIndex, 0, "||Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta & "´" & FontTypeNames.FONTTYPE_info)
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & "´" & FontTypeNames.FONTTYPE_info)
        Else
                Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & "´" & FontTypeNames.FONTTYPE_info)
        End If

        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef & "´" & FontTypeNames.FONTTYPE_info)
        Else
                Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & "´" & FontTypeNames.FONTTYPE_info)
        End If
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef & "´" & FontTypeNames.FONTTYPE_info)
        Else
                Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & "´" & FontTypeNames.FONTTYPE_info)
        End If
        '[GAU]
If UserList(UserIndex).Invent.BotaEqpObjIndex > 0 Then
        Call SendData(ToIndex, sendIndex, 0, "||(PIES) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.BotaEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.BotaEqpObjIndex).MaxDef & "´" & FontTypeNames.FONTTYPE_info)
Else
        Call SendData(ToIndex, sendIndex, 0, "||(PIES) Min Def/Max Def: 0" & "´" & FontTypeNames.FONTTYPE_info)
End If
'[GAU]
        
        If UserList(UserIndex).GuildInfo.GuildName <> "" Then
                Call SendData(ToIndex, sendIndex, 0, "||Clan: " & UserList(UserIndex).GuildInfo.GuildName & "´" & FontTypeNames.FONTTYPE_info)
                If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
                        If UserList(UserIndex).GuildInfo.ClanFundado = UserList(UserIndex).GuildInfo.GuildName Then
                                Call SendData(ToIndex, sendIndex, 0, "||Status:" & "Fundador/Lider" & "´" & FontTypeNames.FONTTYPE_info)
                        Else
                                Call SendData(ToIndex, sendIndex, 0, "||Status:" & "Lider" & "´" & FontTypeNames.FONTTYPE_info)
                        End If
                Else
                        Call SendData(ToIndex, sendIndex, 0, "||Status:" & UserList(UserIndex).GuildInfo.GuildPoints & "´" & FontTypeNames.FONTTYPE_info)
                End If
                Call SendData(ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & "´" & FontTypeNames.FONTTYPE_info)
        End If
Call SendData(ToIndex, sendIndex, 0, "||Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map & "´" & FontTypeNames.FONTTYPE_info)
'pluto:2.15
Call SendData(ToIndex, sendIndex, 0, "||Muertes Ciudas: " & UserList(UserIndex).Faccion.CiudadanosMatados & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, sendIndex, 0, "||Muertes Crimis: " & UserList(UserIndex).Faccion.CriminalesMatados & "´" & FontTypeNames.FONTTYPE_info)

If UserList(UserIndex).Embarazada > 0 Then
Call SendData(ToIndex, sendIndex, 0, "||Min. Embarazo:" & UserList(UserIndex).Embarazada & "´" & FontTypeNames.FONTTYPE_info)
End If
'PLUTO:2-3-04
'Call SendData(ToIndex, sendIndex, 0, "||DragPuntos: " & UserList(UserIndex).Stats.Puntos & FONTTYPENAMES.FONTTYPE_INFO)
Exit Sub
fallo:
Call LogError("senduserstatstxt " & Err.number & " D: " & Err.Description)

End Sub
Sub SendESTADISTICAS(ByVal UserIndex As Integer)
  On Error GoTo fallo
        
 Dim ci As String
 Dim ww1 As Integer
 Dim ww2 As Integer
 Dim ww3 As Byte
 Dim ww4 As Byte
 Dim ww5 As Byte
 Dim ww6 As Byte
 Dim ww7 As Byte
 Dim ww8 As Byte
 Dim ww9 As Byte
 Dim ww10 As Byte
 'pluto:7.0
 Dim AciertoArmas As Integer
 Dim AciertoProyectiles As Integer
 Dim DañoArmas As Integer
 Dim DañoProyectiles As Integer
 Dim Evasion As Integer
 Dim EvasionProyec As Integer
 Dim Escudos As Integer
 Dim ResisMagia As Integer
 Dim DañoMagia As Integer
 Dim DefensaFisica As Integer
 
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        ww1 = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT
        ww2 = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT
Else
        ww1 = 0
        ww2 = 0
End If
If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        ww3 = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef
        ww4 = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef
Else
        ww3 = 0
        ww4 = 0
End If
If UserList(UserIndex).Invent.BotaEqpObjIndex > 0 Then
        ww5 = ObjData(UserList(UserIndex).Invent.BotaEqpObjIndex).MinDef
        ww6 = ObjData(UserList(UserIndex).Invent.BotaEqpObjIndex).MaxDef
Else
        ww5 = 0
        ww6 = 0
End If
If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        ww7 = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef
        ww8 = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef
Else
        ww7 = 0
        ww8 = 0
End If
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        ww9 = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MinDef
        ww10 = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MaxDef
Else
        ww9 = 0
        ww10 = 0
End If
ci = UserList(UserIndex).Stats.MinHIT & "," & UserList(UserIndex).Stats.MaxHIT & "," & ww1 & "," & ww2 & "," & ww7 & "," & ww8 & "," & ww3 & "," & ww4 & "," & ww9 & "," & ww10 & "," & ww5 & "," & ww6 & "," & UserList(UserIndex).GuildInfo.GuildPoints
'pluto:2.22
Dim Solicit As Integer
Solicit = (10 + Int(UserList(UserIndex).Mision.numero / 20) - UserList(UserIndex).GuildInfo.ClanesParticipo)
ci = ci & "," & UserList(UserIndex).Stats.PClan & "," & UserList(UserIndex).Stats.Puntos & "," & UserList(UserIndex).Stats.GTorneo & "," & UserList(UserIndex).GuildInfo.ClanesParticipo & "," & Solicit & "," & UserList(UserIndex).Mision.numero
'pluto:7.0
'acierto armas
AciertoArmas = PoderAtaqueArma(UserIndex)
DañoArmas = PoderDañoArma(UserIndex)
AciertoProyectiles = PoderAtaqueProyectil(UserIndex)
DañoProyectiles = PoderDañoProyectiles(UserIndex)
Escudos = PoderEvasionEscudo(UserIndex)
Evasion = PoderEvasion(UserIndex, Tacticas)
EvasionProyec = PoderEvasion(UserIndex, EvitarProyec)
ResisMagia = PoderResistenciaMagias(UserIndex)
DañoMagia = PoderDañoMagias(UserIndex)
DefensaFisica = PoderDefensaFisica(UserIndex)
ci = ci & "," & AciertoArmas & "," & DañoArmas & "," & AciertoProyectiles & "," & DañoProyectiles & "," & Escudos & "," & Evasion & "," & EvasionProyec & "," & ResisMagia & "," & DañoMagia & "," & DefensaFisica
Call SendData(ToIndex, UserIndex, 0, "J2" & ci)
Exit Sub

fallo:
Call LogError("sendESTADISTICAS " & Err.number & " D: " & Err.Description)

End Sub
Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo fallo
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).Invent.NroItems & " objetos." & "´" & FontTypeNames.FONTTYPE_info)
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
                Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount & "´" & FontTypeNames.FONTTYPE_info)
        End If
Next
Exit Sub
fallo:
Call LogError("senduserinvtxt " & Err.number & " D: " & Err.Description)

End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo fallo
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
For j = 1 To NUMSKILLS
        Call SendData(ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j) & "´" & FontTypeNames.FONTTYPE_info)
Next
Exit Sub
fallo:
Call LogError("senduserskillstxt " & Err.number & " D: " & Err.Description)

End Sub


Sub UpdateUserMap(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim Map As Integer
Dim X As Integer
Dim Y As Integer

Map = UserList(UserIndex).Pos.Map
'pluto:2.17 añade ciudades sin invi
If MapInfo(UserList(UserIndex).Pos.Map).Invisible = 1 Then
UserList(UserIndex).flags.Invisible = 0
UserList(UserIndex).Counters.Invisibilidad = 0
UserList(UserIndex).flags.Oculto = 0
End If

For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

                If MapData(Map, X, Y).UserIndex > 0 And UserIndex <> MapData(Map, X, Y).UserIndex Then
                        Call MakeUserChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)
                        If UserList(MapData(Map, X, Y).UserIndex).flags.Invisible = 1 Then Call SendData2(ToIndex, UserIndex, 0, 17, UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
                End If

                If MapData(Map, X, Y).NpcIndex > 0 Then
                        Call MakeNPCChar(ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
'pluto:6.0A-----------------------------
If Npclist(MapData(Map, X, Y).NpcIndex).Raid > 0 Then
Call SendData(ToMapButIndex, UserIndex, UserList(UserIndex).Pos.Map, "H4" & Npclist(MapData(Map, X, Y).NpcIndex).Char.CharIndex & "," & Npclist(MapData(Map, X, Y).NpcIndex).Stats.MinHP)
End If
'---------------------------------------------------
                End If

                If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
                        Call MakeObj(ToIndex, UserIndex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                        
                        If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = OBJTYPE_PUERTAS Then
                                          Call Bloquear(ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                                          Call Bloquear(ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                        End If
                End If
                
        Next X
Next Y
Exit Sub
fallo:
Call LogError("updateusermap " & Err.number & " D: " & Err.Description)

End Sub

Function DameUserindex(SocketId As Integer) As Integer
On Error GoTo fallo
Dim loopc As Integer
  
loopc = 1
  
Do Until UserList(loopc).ConnID = SocketId

        loopc = loopc + 1
        
        If loopc > MaxUsers Then
                DameUserindex = 0
                Exit Function
        End If
        
Loop
  
DameUserindex = loopc
Exit Function
fallo:
Call LogError("dameuserindex " & Err.number & " D: " & Err.Description)

End Function

Function DameUserIndexConNombre(ByVal Nombre As String) As Integer
On Error GoTo fallo
Dim loopc As Integer
  
loopc = 1
  
Nombre = UCase$(Nombre)

Do Until UCase$(UserList(loopc).Name) = Nombre

        loopc = loopc + 1
        
        If loopc > MaxUsers Then
                DameUserIndexConNombre = 0
                Exit Function
        End If
        
Loop
  
DameUserIndexConNombre = loopc
Exit Function
fallo:
Call LogError("dameuserindexconnombre " & Err.number & " D: " & Err.Description)

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
If Npclist(NpcIndex).MaestroUser > 0 Then
'pluto:2.18
If UserList(UserIndex).Faccion.ArmadaReal > 0 And Not Criminal(Npclist(NpcIndex).MaestroUser) Then
EsMascotaCiudadano = True
Exit Function
End If
'---------------
EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
If EsMascotaCiudadano Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
End If
Exit Function
fallo:
Call LogError("esmascotaciudadano " & Err.number & " D: " & Err.Description)

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
Dim MinPc As npc

MinPc = Npclist(NpcIndex)



On Error GoTo fallo
'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name

  'respown esbirros del caballero de la muerte
   
   
If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 500000 And Npclist(NpcIndex).Stats.MinHP > 499000 Then Call SpawnNpc(727, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 500000 And Npclist(NpcIndex).Stats.MinHP > 499000 Then Call SpawnNpc(728, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 500000 And Npclist(NpcIndex).Stats.MinHP > 499000 Then Call SpawnNpc(729, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 500000 And Npclist(NpcIndex).Stats.MinHP > 499000 Then Call SpawnNpc(730, MinPc.Pos, True, False)


If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 400000 And Npclist(NpcIndex).Stats.MinHP > 399000 Then Call SpawnNpc(727, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 400000 And Npclist(NpcIndex).Stats.MinHP > 399000 Then Call SpawnNpc(728, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 400000 And Npclist(NpcIndex).Stats.MinHP > 399000 Then Call SpawnNpc(729, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 400000 And Npclist(NpcIndex).Stats.MinHP > 399000 Then Call SpawnNpc(730, MinPc.Pos, True, False)


If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 200000 And Npclist(NpcIndex).Stats.MinHP > 199000 Then Call SpawnNpc(727, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 200000 And Npclist(NpcIndex).Stats.MinHP > 199000 Then Call SpawnNpc(728, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 200000 And Npclist(NpcIndex).Stats.MinHP > 199000 Then Call SpawnNpc(729, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 200000 And Npclist(NpcIndex).Stats.MinHP > 199000 Then Call SpawnNpc(730, MinPc.Pos, True, False)


If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 20000 And Npclist(NpcIndex).Stats.MinHP > 18000 Then Call SpawnNpc(727, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 20000 And Npclist(NpcIndex).Stats.MinHP > 18000 Then Call SpawnNpc(728, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 20000 And Npclist(NpcIndex).Stats.MinHP > 18000 Then Call SpawnNpc(729, MinPc.Pos, True, False)

If MinPc.numero = 726 And Npclist(NpcIndex).Pos.Map = 279 And Npclist(NpcIndex).Stats.MinHP < 20000 And Npclist(NpcIndex).Stats.MinHP > 18000 Then Call SpawnNpc(730, MinPc.Pos, True, False)
   
   
   'respown esbirros del caballero de la muerte




'COMPROBAMOS ATAQUE A CASTILLOS
'rey herido
If Npclist(NpcIndex).Pos.Map = 185 And Npclist(NpcIndex).Name = "Defensor Fortaleza" Then
Call SendData(ToAll, 0, 0, "V8")
AtaForta = 1
End If
'If Npclist(NpcIndex).Pos.Map = 185 And Npclist(NpcIndex).Name = "Defensor Fortaleza" And Npclist(NpcIndex).Stats.MinHP > 5000 And Npclist(NpcIndex).Stats.MinHP < 6000 Then Call SendData(ToAll, 0, 0, "V9")

'If Npclist(NpcIndex).Pos.Map = mapa_castillo1 And Npclist(NpcIndex).NPCtype = 33 Or (Npclist(NpcIndex).Pos.Map = mapa_castillo1 + 102 And (Npclist(NpcIndex).NPCtype = 77 Or Npclist(NpcIndex).NPCtype = 78)) Then
'pluto:6.0A cambio la linea de arriba por la de abajo
If Npclist(NpcIndex).Pos.Map = mapa_castillo1 And (Npclist(NpcIndex).NPCtype = 33 Or Npclist(NpcIndex).NPCtype = 78) Then
Call SendData(ToAll, 0, 0, "C1")
AtaNorte = 1
End If

If Npclist(NpcIndex).Pos.Map = mapa_castillo2 And (Npclist(NpcIndex).NPCtype = 33 Or Npclist(NpcIndex).NPCtype = 78) Then
Call SendData(ToAll, 0, 0, "C2")
AtaSur = 1
End If
If Npclist(NpcIndex).Pos.Map = mapa_castillo3 And (Npclist(NpcIndex).NPCtype = 33 Or Npclist(NpcIndex).NPCtype = 78) Then
Call SendData(ToAll, 0, 0, "C3")
AtaEste = 1
End If
If Npclist(NpcIndex).Pos.Map = mapa_castillo4 And (Npclist(NpcIndex).NPCtype = 33 Or Npclist(NpcIndex).NPCtype = 78) Then
Call SendData(ToAll, 0, 0, "C4")
AtaOeste = 1
End If
If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
If EsMascotaCiudadano(NpcIndex, UserIndex) Then
                        Call VolverCriminal(UserIndex)
                        Npclist(NpcIndex).Movement = NPCDEFENSA
                        Npclist(NpcIndex).Hostile = 1
Else
        'Reputacion
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
           
           'pluto:2.11
           If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                                Call VolverCriminal(UserIndex)
           Else
                                Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
           End If
        ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
           Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR / 2, MAXREP)
        End If
        
        'hacemos que el npc se defienda
                   Npclist(NpcIndex).Movement = NPCDEFENSA
                   Npclist(NpcIndex).Hostile = 1
        
End If
   'pluto:2.14
If Npclist(NpcIndex).flags.PoderEspecial2 > 0 Then

If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 1 Or (MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y - 1).UserIndex > 0 Or MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y + 1).UserIndex > 0 Or MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.X - 1).UserIndex > 0 Or MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.X + 1).UserIndex > 0) Then
Dim pvalida As Boolean
Dim Newpos As WorldPos
Dim it As Byte
Do While Not pvalida
Call ClosestLegalPos(UserList(UserIndex).Pos, Newpos, Npclist(NpcIndex).flags.AguaValida) 'Nos devuelve la posicion valida mas cercana
If LegalPosNPC(Newpos.Map, Newpos.X, Newpos.Y, Npclist(NpcIndex).flags.AguaValida) Then
                        'Asignamos las nuevas coordenas solo si son validas
MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

                        Npclist(NpcIndex).Pos.Map = Newpos.Map
                        Npclist(NpcIndex).Pos.X = Newpos.X
                        Npclist(NpcIndex).Pos.Y = Newpos.Y
                        pvalida = True
End If
it = it + 1
If it > 20 Then Exit Sub
Loop
Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Pos.X & "," & Npclist(NpcIndex).Pos.Y & ",0")
End If
End If

'pluto:2.20 añado >0
If Npclist(NpcIndex).flags.PoderEspecial5 > 0 And Npclist(NpcIndex).Stats.MinHP > 0 Then
Dim n2 As Byte
n2 = RandomNumber(1, 100)

If n2 > 70 Then
Call SendData2(ToMap, 0, Npclist(NpcIndex).Pos.Map, 22, Npclist(NpcIndex).Char.CharIndex & "," & 31 & "," & 1)
Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + 300
'Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "TW" & 18)
If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
End If

End If



Exit Sub
fallo:
Call LogError("npcatacado " & Err.number & " D: " & Err.Description)

End Sub
Function PuedeDobleArma(ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
'pluto:2.15
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).SubTipo = 6 Then
 PuedeDobleArma = True
Else
 PuedeDobleArma = False
End If
End If
Exit Function
fallo:
Call LogError("puededoblearma " & Err.number & " D: " & Err.Description)

End Function
Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
'pluto:2.15
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 And UserList(UserIndex).clase <> "Druida" Then
 PuedeApuñalar = _
 ((UserList(UserIndex).Stats.UserSkills(Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UserList(UserIndex).clase = "Asesino") And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
Exit Function
fallo:
Call LogError("puedeapuñalar " & Err.number & " D: " & Err.Description)

End Function
Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
On Error GoTo fallo
'pluto:2.17
If UserList(UserIndex).Bebe > 0 Then Exit Sub

If UserList(UserIndex).flags.Hambre = 0 And _
   UserList(UserIndex).flags.Sed = 0 Then
        Dim Aumenta As Integer
        Dim PROB As Integer
'pluto:6.3--------------
If ServerPrimario = 1 Then
PROB = 15
Else
PROB = 6
End If
'--------------------------
Aumenta = Int(RandomNumber(1, PROB))
        
        Dim lvl As Integer
        lvl = UserList(UserIndex).Stats.ELV
        
        If lvl >= UBound(LevelSkill) Then Exit Sub
        If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
        
        'nati: aumento los skillpoint a 5
        If Aumenta < 3 And UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(lvl).LevelValue Then
                        Call AddtoVar(UserList(UserIndex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
                        Call SendData(ToIndex, UserIndex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts." & "´" & FontTypeNames.FONTTYPE_info)
                        'pluto:2.19
                        Dim Sk As Long
                        Sk = UserList(UserIndex).Stats.UserSkills(Skill) * 20
                        Call AddtoVar(UserList(UserIndex).Stats.exp, Sk, MAXEXP)
                        Call SendData(ToIndex, UserIndex, 0, "||¡Has ganado " & Sk & " puntos de experiencia!" & "´" & FontTypeNames.FONTTYPE_FIGHT)
                        'pluto:2.4.5
                        Call CheckUserLevel(UserIndex)
                        'pluto:2.17
                        Call EnviaUnSkills(UserIndex, Skill)
        End If

End If
Exit Sub
fallo:
Call LogError("subirskill Nom: " & UserList(UserIndex).Name & " Sk: " & Skill & Err.number & " D: " & Err.Description)

End Sub

Sub UserDie(ByVal UserIndex As Integer)
'Call LogTarea("Sub UserDie")
On Error GoTo ErrorHandler


'Sonido
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_USERMUERTE)

'PLUTO:6.8---------------
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'--------------------------

'Quitar el dialogo del user muerto
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 21, UserList(UserIndex).Char.CharIndex)

'pluto:2.11.0
If UserList(UserIndex).GranPoder > 0 Then
UserList(UserIndex).GranPoder = 0
UserGranPoder = ""
UserList(UserIndex).Char.FX = 0
Call SendData2(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 68 & "," & 0)
End If

UserList(UserIndex).ObjetosTirados = 0
UserList(UserIndex).Stats.MinHP = 0
UserList(UserIndex).flags.AtacadoPorNpc = 0
UserList(UserIndex).flags.AtacadoPorUser = 0
UserList(UserIndex).flags.Envenenado = 0
UserList(UserIndex).flags.Muerto = 1
UserList(UserIndex).flags.Morph = 0
UserList(UserIndex).flags.Angel = 0
UserList(UserIndex).flags.Demonio = 0
'pluto:6.2
'UserList(UserIndex).flags.ParejaTorneo = 0
'pluto:2.9.0
If UserList(UserIndex).flags.Comerciando = True Then
Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
Call FinComerciarUsu(UserIndex)
End If

Dim aN As Integer

aN = UserList(UserIndex).flags.AtacadoPorNpc

If aN > 0 Then
          Npclist(aN).Movement = Npclist(aN).flags.OldMovement
          Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
          Npclist(aN).flags.AttackedBy = ""
End If

'<<<< Paralisis >>>>
If UserList(UserIndex).flags.Paralizado = 1 Then
        UserList(UserIndex).flags.Paralizado = 0
  Call SendData2(ToIndex, UserIndex, 0, 68)
End If
' invisibilidad
 If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
        UserList(UserIndex).flags.Invisible = 0
        UserList(UserIndex).Counters.Invisibilidad = 0
        UserList(UserIndex).flags.Oculto = 0
        Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",0")
End If
' estupidez
' If UserList(userindex).Flags.Estupidez = 1 Then
'UserList(userindex).Flags.Estupidez = 0
         'Call SendData(ToIndex, userindex, 0, "NESTUP")
                'End If

' ceguera
If UserList(UserIndex).flags.Ceguera = 1 Then
                UserList(UserIndex).flags.Ceguera = 0
                Call SendData2(ToIndex, UserIndex, 0, 55)
End If
'<<<< Descansando >>>>
If UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call SendData2(ToIndex, UserIndex, 0, 41)
End If

'<<<< Meditando >>>>
If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        Call SendData2(ToIndex, UserIndex, 0, 54)
End If

        'desequipar armadura
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
        End If
        'desequipar arma
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
        End If
        'desequipar casco
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
        End If
          'desequipar casco
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
        End If
        '[GAU]
'desequipar botas
If UserList(UserIndex).Invent.BotaEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.BotaEqpSlot)
End If
'[GAU]
'Pluto:2.4
 If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
End If
'----fin Pluto:2.4---------
        'desequipar herramienta
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
        End If
        'desequipar municiones
        If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
        End If

' << Si es newbie no pierde el inventario >>
        
'pluto:7.0
If UserList(UserIndex).raza = "Goblin" And RandomNumber(1, 100) > 75 Then GoTo notirarnada

If Not EsNewbie(UserIndex) Or Criminal(UserIndex) Then
        Call TirarTodo(UserIndex)
Else
        If EsNewbie(UserIndex) Then Call TirarTodosLosItemsNoNewbies(UserIndex)
End If

notirarnada:

If UserList(UserIndex).Remort = 0 Then
        If UserList(UserIndex).Stats.MaxMAN > STAT_MAXMAN Then UserList(UserIndex).Stats.MaxMAN = STAT_MAXMAN
Else
        Select Case UCase$(UserList(UserIndex).clase)
                Case "MAGO"
                        If UserList(UserIndex).Stats.MaxMAN > 5000 Then UserList(UserIndex).Stats.MaxMAN = 5000
                Case "CLERIGO"
                        If UserList(UserIndex).Stats.MaxMAN > 4000 Then UserList(UserIndex).Stats.MaxMAN = 4000
                Case "DRUIDA"
                        If UserList(UserIndex).Stats.MaxMAN > 3500 Then UserList(UserIndex).Stats.MaxMAN = 3500
                Case "BARDO"
                        If UserList(UserIndex).Stats.MaxMAN > 3500 Then UserList(UserIndex).Stats.MaxMAN = 3500
                Case "PALADIN"
                        If UserList(UserIndex).Stats.MaxMAN > 3000 Then UserList(UserIndex).Stats.MaxMAN = 3000
        End Select
End If
' << Reseteamos los posibles FX sobre el personaje >>
If UserList(UserIndex).Char.loops = LoopAdEternum Then
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
End If
'<< Cambiamos la apariencia del char >>
If UserList(UserIndex).flags.Navegando = 0 Then
        
        If Not Criminal(UserIndex) Then
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        Else
        UserList(UserIndex).Char.Body = iCuerpoMuerto2
        UserList(UserIndex).Char.Head = iCabezaMuerto2
        End If
        
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
'[GAU]
        UserList(UserIndex).Char.Botas = NingunBota
        '[GAU]
Else
        UserList(UserIndex).Char.Body = iFragataFantasmal ';)
End If

Dim i As Integer
For i = 1 To MAXMASCOTAS
If UserList(UserIndex).MascotasIndex(i) > 0 Then
If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
Else
Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
'pluto:2.4
Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))

UserList(UserIndex).MascotasIndex(i) = 0
UserList(UserIndex).MascotasType(i) = 0
End If
End If
Next i
UserList(UserIndex).NroMacotas = 0
'pluto:2.3
Call SendData2(ToIndex, UserIndex, 0, 56)
If UserList(UserIndex).flags.Montura = 1 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax - (UserList(UserIndex).flags.ClaseMontura * 100)
Call SendUserStatsPeso(UserIndex)
End If
UserList(UserIndex).flags.ClaseMontura = 0
UserList(UserIndex).flags.Montura = 0
'pluto:6.3
UserList(UserIndex).flags.Estupidez = 0
Call SendData2(ToIndex, UserIndex, 0, 56)

'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
'        Dim MiObj As Obj
'        Dim nPos As WorldPos
'        MiObj.ObjIndex = RandomNumber(554, 555)
'        MiObj.Amount = 1
'        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
'        Dim ManchaSangre As New cGarbage
'        ManchaSangre.Map = nPos.Map
'        ManchaSangre.X = nPos.X
'        ManchaSangre.Y = nPos.Y
'        Call TrashCollector.Add(ManchaSangre)
'End If

'<< Actualizamos clientes >>
'[GAU] Agregamo NingunBota




Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, val(UserIndex), UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunBota)
'[GAU]
Call senduserstatsbox(UserIndex)




Exit Sub

ErrorHandler:
        Call LogError("Error en SUB USERDIE Nom:" & UserList(UserIndex).Name)

End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal atacante As Integer)
On Error GoTo ErrorHandler

 'pluto:2.11
 If UserList(Muerto).GranPoder > 0 Then
 UserList(Muerto).GranPoder = 0
 UserList(Muerto).Char.FX = 0
 UserList(atacante).GranPoder = 1
 UserGranPoder = UserList(atacante).Name
 End If

   
   
'pluto:hoy
If UserList(atacante).Mision.estado = 1 And UserList(atacante).Mision.clase = UCase$(UserList(Muerto).clase) And ((UserList(Muerto).Stats.ELV) >= UserList(atacante).Mision.Level) Then
        UserList(atacante).Mision.estado = 0
        Call SendData(ToIndex, atacante, 0, "!!Quest Número " & UserList(atacante).Mision.numero & " : " & " Muy bién, has cumplido una misión!!")
'pluto:2-3-04
Call SendData(ToIndex, atacante, 0, "|| Has ganado " & val(Int(UserList(atacante).Mision.numero / 10) + 1) & " DragPuntos." & "´" & FontTypeNames.FONTTYPE_info)
 'pluto:6.0A
                UserList(atacante).Stats.Fama = UserList(atacante).Stats.Fama + 5
'pluto:2.19----------------------------------------------
Call SendData(ToIndex, atacante, 0, "|| Has ganado " & val(Int(UserList(atacante).Mision.numero * 500)) & " Puntos de Experiencia." & "´" & FontTypeNames.FONTTYPE_info)
UserList(atacante).Stats.exp = UserList(atacante).Stats.exp + Int(UserList(atacante).Mision.numero * 500)
SendUserStatsEXP (atacante)
CheckUserLevel (atacante)
'----------------------------------------------

UserList(atacante).Stats.Puntos = UserList(atacante).Stats.Puntos + Int(UserList(atacante).Mision.numero / 10) + 1
End If




'--------fin pluto:2.4----------------

If UserList(atacante).Pos.Map = MAPATORNEO Then
'REVISAR
'pluto:muere en torneo
Call SendData(ToMap, 0, 296, "||Torneo: " & UserList(atacante).Name & " derrota a " & UserList(Muerto).Name & "´" & FontTypeNames.FONTTYPE_TALK)
'Delzak) aviso nix y caos
Call SendData(ToMap, 0, 34, "||Torneo: " & UserList(atacante).Name & " derrota a " & UserList(Muerto).Name & "´" & FontTypeNames.FONTTYPE_TALK)
Call SendData(ToMap, 0, 170, "||Torneo: " & UserList(atacante).Name & " derrota a " & UserList(Muerto).Name & "´" & FontTypeNames.FONTTYPE_TALK)
        'Tite añade aviso Bander
'Call SendData(ToMap, 0, 59, "||Torneo: " & UserList(atacante).Name & " derrota a " & UserList(Muerto).Name & "´" & FontTypeNames.FONTTYPE_talk)
        '\Tite
'gana torneo
        If UserList(atacante).flags.LastCiudMatado <> UserList(Muerto).Name And UserList(atacante).flags.LastCrimMatado <> UserList(Muerto).Name Then
If Criminal(Muerto) Then UserList(atacante).flags.LastCrimMatado = UserList(Muerto).Name Else UserList(atacante).flags.LastCiudMatado = UserList(Muerto).Name
'UserList(atacante).Stats.GLD = UserList(atacante).Stats.GLD + (25 * UserList(Muerto).Stats.ELV)
Call AddtoVar(UserList(atacante).Stats.GLD, (25 * UserList(Muerto).Stats.ELV), MAXORO)

Call SendData(ToIndex, atacante, 0, "||Has ganado " & 25 * UserList(Muerto).Stats.ELV & " monedas." & "´" & FontTypeNames.FONTTYPE_FIGHT)

UserList(atacante).flags.Torneo = UserList(atacante).flags.Torneo + 1
Call SendData(ToPCArea, atacante, UserList(atacante).Pos.Map, "TW" & SND_TORNEO)
'pluto:2.4
UserList(atacante).Stats.GTorneo = UserList(atacante).Stats.GTorneo + 1
UserList(Muerto).Stats.GTorneo = UserList(Muerto).Stats.GTorneo - 1

        End If

If UserList(atacante).flags.Torneo = 5 Then

Call SendData(ToAll, 0, 0, "|| Ganador del torneo 5 veces consecutivas, " & UserList(atacante).Name & " obtiene premio de 2000 oros extras." & "´" & FontTypeNames.FONTTYPE_TALK)
'UserList(atacante).Stats.GLD = UserList(atacante).Stats.GLD + 2000
Call AddtoVar(UserList(atacante).Stats.GLD, 2000, MAXORO)

Call SendData(ToIndex, atacante, 0, "TW" & 180)
UserList(atacante).flags.Torneo = UserList(atacante).flags.Torneo + 1
 'pluto:6.0A
                UserList(atacante).Stats.Fama = UserList(atacante).Stats.Fama + 5
End If
If UserList(atacante).flags.Torneo = 11 Then
Call SendData(ToAll, 0, 0, "|| Ganador del torneo 10 veces consecutivas, " & UserList(atacante).Name & " obtiene premio de 5000 oros extras." & "´" & FontTypeNames.FONTTYPE_TALK)
'UserList(atacante).Stats.GLD = UserList(atacante).Stats.GLD + 5000
Call AddtoVar(UserList(atacante).Stats.GLD, 5000, MAXORO)
Call SendData(ToIndex, atacante, 0, "TW" & 180)
UserList(atacante).flags.Torneo = UserList(atacante).flags.Torneo + 1
 'pluto:6.0A
                UserList(atacante).Stats.Fama = UserList(atacante).Stats.Fama + 15
End If
If UserList(atacante).flags.Torneo = 22 Then
Call SendData(ToAll, 0, 0, "|| Ganador del torneo 20 veces consecutivas, " & UserList(atacante).Name & " obtiene premio de 15000 oros extras." & "´" & FontTypeNames.FONTTYPE_TALK)
'UserList(atacante).Stats.GLD = UserList(atacante).Stats.GLD + 15000
Call AddtoVar(UserList(atacante).Stats.GLD, 15000, MAXORO)
Call SendData(ToIndex, atacante, 0, "TW" & 180)
UserList(atacante).flags.Torneo = UserList(atacante).flags.Torneo + 1
 'pluto:6.0A
                UserList(atacante).Stats.Fama = UserList(atacante).Stats.Fama + 30
End If
'Exit Sub
End If
'pluto:2.12
If UserList(atacante).Pos.Map = MapaTorneo2 Then
UserList(atacante).Torneo2 = UserList(atacante).Torneo2 + 1
If UserList(atacante).Torneo2 > 10 Then UserList(atacante).Torneo2 = 10
UserList(Muerto).Torneo2 = 0
MinutoSinMorir = 0
If UserList(atacante).Torneo2 = 10 Then
'UserList(atacante).Stats.GLD = UserList(atacante).Stats.GLD + TorneoBote
Call AddtoVar(UserList(atacante).Stats.GLD, TorneoBote, MAXORO)
 Call SendData(ToIndex, atacante, 0, "TW" & 180)
 'pluto:6.0A
                UserList(atacante).Stats.Fama = UserList(atacante).Stats.Fama + 10
Call SendUserStatsOro(atacante)
TorneoBote = 0
Torneo2Record = 0

End If

If UserList(atacante).Torneo2 > Torneo2Record Then
Torneo2Record = UserList(atacante).Torneo2
Torneo2Name = UserList(atacante).Name
Call SendData2(ToMap, 0, MapaTorneo2, 96, Torneo2Name & "," & Torneo2Record & "," & TorneoBote)
End If
If UCase$(UserList(Muerto).Name) = UCase$(Torneo2Name) Then
Torneo2Record = 0
End If
'Exit Sub
End If


 '--------------------drags puntos----------------
 'pluto:2-3-04
If UserList(atacante).Stats.Puntos > 0 And UserList(Muerto).Stats.Puntos > 0 And MapInfo(UserList(Muerto).Pos.Map).Pk = True And UserList(Muerto).Stats.ELV > 15 Then
Dim pun As Integer
pun = Int(UserList(Muerto).Stats.ELV / 10) + 1
'pluto.2.5.0
If pun > UserList(Muerto).Stats.Puntos Then pun = UserList(Muerto).Stats.Puntos

UserList(Muerto).Stats.Puntos = UserList(Muerto).Stats.Puntos - pun
'PLUTO:2.4
UserList(atacante).Stats.Puntos = UserList(atacante).Stats.Puntos + pun

If UserList(Muerto).Stats.Puntos < 0 Then UserList(Muerto).Stats.Puntos = 0

Call SendData(ToIndex, atacante, 0, "|| Has ganado " & pun & " DragPuntos." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, Muerto, 0, "|| Has pérdido " & pun & " DragPuntos." & "´" & FontTypeNames.FONTTYPE_info)
UserList(Muerto).flags.CMuerte = 1
End If
   '--------------------fin drags puntos----------------
   
        'pluto:5.2--añado cmuerte (1 minuto)
'If EsNewbie(Muerto) Or UserList(Muerto).flags.CMuerte > 0 Then Exit Sub

 '-----------------------------armadas------------------


'nIcO.M nuevo codigo puntos
'------------------ puntos clan---------------------'nicoo,M
'If UserList(Muerto).GuildInfo.GuildName = "" Or UserList(atacante).GuildInfo.GuildName = "" Or UserList(Muerto).GuildInfo.GuildName = UserList(atacante).GuildInfo.GuildName Or UserList(Muerto).GuildRef.Reputation > 0 Then
If UserList(Muerto).GuildInfo.GuildName = "" Or UserList(atacante).GuildInfo.GuildName = "" Or UserList(atacante).GuildInfo.GuildName = UserList(Muerto).GuildInfo.GuildName Then GoTo gc
UserList(Muerto).GuildRef.Reputation = UserList(Muerto).GuildRef.Reputation - 1
UserList(atacante).GuildRef.Reputation = UserList(atacante).GuildRef.Reputation + 1
UserList(atacante).Stats.PClan = UserList(atacante).Stats.PClan + 1
UserList(Muerto).Stats.PClan = UserList(Muerto).Stats.PClan - 1
Call SendData(ToIndex, atacante, 0, "||Has sumado 1 Punto de membresia!!" & "´" & FontTypeNames.FONTTYPE_pluto)
Call SendData(ToIndex, Muerto, 0, "||Has Restado 1 Punto al Clan!!" & "´" & FontTypeNames.FONTTYPE_pluto)
UserList(Muerto).flags.CMuerte = 1

gc:


'pluto:2.5.0
If Criminal(Muerto) Then
                If UserList(atacante).flags.LastCrimMatado <> UserList(Muerto).Name Then
                                UserList(atacante).flags.LastCrimMatado = UserList(Muerto).Name
'pluto:5.2
UserList(atacante).MuertesTime = UserList(atacante).MuertesTime + 1
UserList(Muerto).flags.CMuerte = 1

                                UserList(atacante).Faccion.CriminalesMatados = UserList(atacante).Faccion.CriminalesMatados + 1
                End If
                
                If UserList(atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
                        UserList(atacante).Faccion.CriminalesMatados = 0
                        UserList(atacante).Faccion.RecompensasReal = 0
                End If
Else
                If UserList(atacante).flags.LastCiudMatado <> UserList(Muerto).Name Then
                        UserList(atacante).flags.LastCiudMatado = UserList(Muerto).Name
'pluto:5.2
UserList(atacante).MuertesTime = UserList(atacante).MuertesTime + 1
UserList(Muerto).flags.CMuerte = 1

                        UserList(atacante).Faccion.CiudadanosMatados = UserList(atacante).Faccion.CiudadanosMatados + 1
                End If
                
                If UserList(atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
                        UserList(atacante).Faccion.CiudadanosMatados = 0
                        UserList(atacante).Faccion.RecompensasCaos = 0
                End If
End If
'pluto:2.15
Call SendUserMuertos(atacante)
Exit Sub
'pluto:2.6.0
ErrorHandler:
Call LogError("Error CONTARMUERTE --> " & Err.number & " D: " & Err.Description & "--> " & UserList(atacante).Name & " -- " & UserList(Muerto).Name & " Puntosclan: " & UserList(atacante).Stats.PClan & "/" & UserList(Muerto).Stats.PClan & " DraGPUntos: " & UserList(atacante).Stats.Puntos & "/" & UserList(Muerto).Stats.Puntos)

End Sub

Sub Tilelibre(Pos As WorldPos, nPos As WorldPos)
'Call LogTarea("Sub Tilelibre")
On Error GoTo fallo
Dim Notfound As Boolean
Dim loopc As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
hayobj = False
nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y) Or hayobj
        
        If loopc > 15 Then
                Notfound = True
                Exit Do
        End If
        
        For tY = Pos.Y - loopc To Pos.Y + loopc
                For tX = Pos.X - loopc To Pos.X + loopc
                
                        If LegalPos(nPos.Map, tX, tY) = True Then
                           hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0)
                           If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                                         nPos.X = tX
                                         nPos.Y = tY
                                         tX = Pos.X + loopc
                                         tY = Pos.Y + loopc
                                End If
                        End If
                
                Next tX
        Next tY
        
        loopc = loopc + 1
        
Loop

If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0
End If
Exit Sub
fallo:
Call LogError("tilelibre " & Err.number & " D: " & Err.Description)

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
On Error GoTo fallo

'pluto:6.5
'DoEvents

'Quitar el dialogo
Dim nPos As WorldPos
Dim xpos As WorldPos
Dim x1 As Byte
Dim Y1 As Byte
If UserIndex = 0 Or Map = 0 Then Exit Sub
'PLUTO:6.6
If UserList(UserIndex).Char.CharIndex = 0 Then
Call LogError("Charindex CERO: index: " & UserIndex & " Name: " & UserList(UserIndex).Name & " Map: " & UserList(UserIndex).Pos.Map & " xy: " & UserList(UserIndex).Pos.X & " " & UserList(UserIndex).Pos.Y)
Exit Sub
End If


'pluto:6.3 aleatorios zonas conflictivas
'If UserList(UserIndex).flags.Privilegios > 0 Then GoTo Noalea

'Select Case Map
'Case 268 To 271
   ' If (X = 40 And Y = 53) Or (X = 43 And Y = 26) Then
        'x1 = RandomNumber(1, 10)
        'Y1 = RandomNumber(1, 10)
   ' X = X + x1
  '  Y = Y + Y1
 '   End If
'Case 185
        'If Y = 82 Then
   ' X = RandomNumber(37, 64)
  '  Y = RandomNumber(83, 91)
 '   End If
'Case 140  'veril
        'If Y = 90 Then
   ' X = RandomNumber(48, 57)
  '  Y = RandomNumber(84, 89)
 '   End If
'Case 141
'   If Y > 90 Then
'    X = RandomNumber(40, 49)
'    Y = RandomNumber(50, 57)
'    End If
'Case 48
'    X = RandomNumber(44, 48)
'    Y = RandomNumber(53, 57)
'Case 156
'   X = RandomNumber(75, 79)
'    Y = RandomNumber(72, 75)
'Case 162
'X = RandomNumber(76, 80)
'Y = RandomNumber(81, 85)
'Case 165
'X = RandomNumber(70, 75)
'Y = RandomNumber(19, 23)
'Case 59
'If Y = 50 Then
'X = RandomNumber(43, 52)
'Y = RandomNumber(47, 51)
'End If
'Case 166 To 169
'If Y = 81 Then
'X = RandomNumber(40, 57)
'Y = RandomNumber(77, 84)
'End If
'Case MAPATORNEO
'X = RandomNumber(52, 71)
'Y = RandomNumber(44, 59)
'Case 291 To 295
'X = RandomNumber(52, 71)
'Y = RandomNumber(44, 59)
'Case MapaTorneo2
'X = RandomNumber(52, 71)
'Y = RandomNumber(44, 59)
'Case 296
'X = RandomNumber(70, 76)
'Y = RandomNumber(60, 66)
'End Select
'----------------------------------

'Noalea:
'pluto:2.18
'If Map > 267 And Map < 272 And ((x = 40 And Y = 53) Or (x = 43 And Y = 26)) Then

'x1 = RandomNumber(1, 10)
'Y1 = RandomNumber(1, 10)
'x = x + x1
'Y = Y + Y1
'End If

'If Map = 185 And Y = 82 Then
'x = RandomNumber(37, 64)
'Y = RandomNumber(83, 91)
'End If
'-----------------------
'pluto:6.4 demonios/angeles en castillos
If Map = MapaAngel Or (Map > 165 And Map < 170) Or Map = 185 Then
        If UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Then
        UserList(UserIndex).Stats.MinSta = 0
        End If
End If

'pluto:2.18----------------------------
xpos.Map = Map
xpos.Y = Y
xpos.X = X
Dim aguita As Byte
If UserList(UserIndex).flags.Navegando = 1 Then aguita = 1 Else aguita = 0
'pluto:6.0A-----------------------------------
If UserList(UserIndex).flags.Privilegios = 0 Then
Call ClosestLegalPos(xpos, nPos, aguita)
Else
nPos.X = X
nPos.Y = Y
End If
'---------------------------------------------
If nPos.X <> 0 And nPos.Y <> 0 Then 'end if al final
X = nPos.X
Y = nPos.Y

Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 21, UserList(UserIndex).Char.CharIndex)

Call SendData2(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, 5)

'pluto:2.7.1
'If Y > 90 Then Y = Y - 2


Dim Oldmap As Integer
Dim OldX As Integer
Dim OldY As Integer

Oldmap = UserList(UserIndex).Pos.Map
'UserList(UserIndex).flags.MapaIncor = Oldmap
OldX = UserList(UserIndex).Pos.X
OldY = UserList(UserIndex).Pos.Y
'pluto:2.9.0 ropa futbol
'If OldMap = 192 And Map <> 192 Then
'If TieneObjetos(1005, 1, UserIndex) Then
'Call QuitarObjetos(1005, 10000, UserIndex)
'End If
'If TieneObjetos(1006, 1, UserIndex) Then
'Call QuitarObjetos(1006, 10000, UserIndex)
'End If
'If TieneObjetos(1007, 1, UserIndex) Then
'Call QuitarObjetos(1007, 10000, UserIndex)
'End If
'If TieneObjetos(1008, 1, UserIndex) Then
'Call QuitarObjetos(1008, 10000, UserIndex)
'End If

'End If '192

Call EraseUserCharMismoIndex(UserIndex)

'pluto:6.2-----------------------
If Oldmap = 291 And UserList(UserIndex).flags.ParejaTorneo > 0 Then
UserList(UserList(UserIndex).flags.ParejaTorneo).flags.ParejaTorneo = 0
UserList(UserIndex).flags.ParejaTorneo = 0
End If
'pluto:6.8---
If Oldmap = 292 Then
        If UserList(UserIndex).GuildInfo.GuildName = TorneoClan(1).Nombre Then TorneoClan(1).numero = TorneoClan(1).numero - 1
        If UserList(UserIndex).GuildInfo.GuildName = TorneoClan(2).Nombre Then TorneoClan(2).numero = TorneoClan(2).numero - 1
End If
'---------------
'If Oldmap = 292 And UserList(UserIndex).flags.Privilegios = 0 Then
 '   If UserList(UserIndex).GuildInfo.GuildName = TorneoClan(1).Nombre Then
  '      TorneoClan(1).Numero = TorneoClan(1).Numero - 1
   '         If TorneoClan(1).Numero = 0 Then
        '        TClanOcupado = TClanOcupado - 1
         '       TorneoClan(1).Nombre = ""
          '      End If
 '   ElseIf UserList(UserIndex).GuildInfo.GuildName = TorneoClan(2).Nombre Then
  '      TorneoClan(2).Numero = TorneoClan(2).Numero - 1
   '         If TorneoClan(2).Numero = 0 Then
        '        TClanOcupado = TClanOcupado - 1
         '       TorneoClan(2).Nombre = ""
          '      End If
   ' End If
'End If
'--------------------------
'pluto:2.19
'mapa de conquista
'If MapInfo(Map).Terreno = "CONQUISTA" Then
'Dim r12
'Dim y12
   '    r12 = RandomNumber(33, 52)
          ' y12 = RandomNumber(25, 32)
'UserList(UserIndex).Pos.X = r12
'UserList(UserIndex).Pos.Y = y12
 'GoTo tt
'End If
'---------------------------------------

UserList(UserIndex).Pos.X = X
UserList(UserIndex).Pos.Y = Y
tt:
UserList(UserIndex).Pos.Map = Map




If Oldmap <> Map Then
        Call SendData2(ToIndex, UserIndex, 0, 14, Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion)
        If MapInfo(Map).Terreno = "BOSQUE" Then
        MapInfo(Map).Music = "58-1"
        ElseIf MapInfo(Map).Terreno = "MAR" Then
        MapInfo(Map).Music = "74-1"
        End If
        'pluto:6.0a
        If MapInfo(Map).Music <> MapInfo(Oldmap).Music Then
        Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)
        End If
'Call SendData(ToIndex, UserIndex, 0, "TM" & 25)
        Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        
        Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

        'Update new Map Users
   If UserList(UserIndex).flags.Privilegios = 0 Then MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        
 'Update old Map Users
        If UserList(UserIndex).flags.Privilegios = 0 Then MapInfo(Oldmap).NumUsers = MapInfo(Oldmap).NumUsers - 1
          
        If MapInfo(Oldmap).NumUsers < 0 Then
                MapInfo(Oldmap).NumUsers = 0
        End If


        
'pluto:6.0A-------------solidos mapa 274--------------------------
If Map = 274 Then
Dim a As Byte
'Dim x As Byte
Dim b As Byte
Dim Salida As Byte
Dim obj As obj
If SolidoGirando = 0 Then GoTo nogiraba
'SolidoGirando = 5
b = 45 + (SolidoGirando * 2)
'quitamos al que gira-----------------
If MapData(Map, b, 26).OBJInfo.ObjIndex = 1170 + SolidoGirando Then
Call EraseObj(ToMap, UserIndex, Map, 10000, Map, b, 26)
 obj.Amount = 1
 obj.ObjIndex = 1175 + SolidoGirando
 Call MakeObj(ToMap, 0, Map, obj, Map, b, 26)
 Salida = 6 + (SolidoGirando * 10)
 MapData(Map, Salida, 11).TileExit.Map = 28
 MapData(Map, Salida, 11).TileExit.X = 46
 MapData(Map, Salida, 11).TileExit.Y = 86
 End If
'fin quitamos girar---------

'ponemos nuevo solido a girar------------
nogiraba:
a = RandomNumber(1, 5)
b = 45 + (a * 2)

If MapData(Map, b, 26).OBJInfo.ObjIndex = 1175 + a Then
Call EraseObj(ToMap, UserIndex, Map, 10000, Map, b, 26)
 obj.Amount = 1
 obj.ObjIndex = 1170 + a
 Call MakeObj(ToMap, 0, Map, obj, Map, b, 26)
 SolidoGirando = a
  Salida = 6 + (SolidoGirando * 10)
 MapData(Map, Salida, 11).TileExit.Map = 276
 MapData(Map, Salida, 11).TileExit.X = 43
 MapData(Map, Salida, 11).TileExit.Y = 83
 End If

End If 'mapa 274
'---------fin solidos------------------------------------------









'pluto:2.12
If MapInfo(Oldmap).NumUsers = 0 And Oldmap = MapaTorneo2 Then MinutoSinMorir = 0
If Oldmap = MapaTorneo2 Then
UserList(UserIndex).Torneo2 = 0
Torneo2Record = 0
Call SendData2(ToIndex, UserIndex, 0, 96, Torneo2Name & "," & Torneo2Record & "," & TorneoBote)
End If
'pluto:6.8 lo coloco aquí sólo cuando es distinto mapa
Call UpdateUserMap(UserIndex)

Else 'mismo mapa
        
        
        Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

End If

'pluto:6.8 cambio a arriba en distintos mapas
'Call UpdateUserMap(UserIndex)

'pluto:2-3-04
If FX And UserList(UserIndex).flags.Privilegios = 0 Then 'FX
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_WARP)
        Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If
'[MerLiNz:X]
If (UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
        Call SendData2(ToMap, 0, Map, 17, UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
        Call SendData2(ToIndex, UserIndex, 0, 17, UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
End If
'[\END]
'pluto:6.9------------
'Call EfectoIncor(UserIndex)
If UserList(UserIndex).flags.MapaIncor <> Map Then
UserList(UserIndex).flags.Incor = True
UserList(UserIndex).Counters.Incor = 0
'UserList(UserIndex).flags.MapaIncor = Oldmap
End If
UserList(UserIndex).flags.MapaIncor = Oldmap
'PLUTO:6.3---------------
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'--------------------------

'UserList(UserIndex).flags.Macreanda = 0
'Call SendData2(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 61 & "," & 1)
'UserList(UserIndex).Char.FX = 61
'-----------------------

Call WarpMascotas(UserIndex)

'pluto:2.12
If Map = MapaTorneo2 And UserList(UserIndex).flags.Privilegios = 0 And Oldmap <> MapaTorneo2 Then
If Torneo2Name = "" Then Torneo2Name = UserList(UserIndex).Name: Torneo2Record = 0
TorneoBote = TorneoBote + 100
Call SendData2(ToMap, 0, MapaTorneo2, 96, Torneo2Name & "," & Torneo2Record & "," & TorneoBote)

'Call SendData2(ToIndex, UserIndex, 0, 96, Torneo2Name & "," & Torneo2Record & "," & TorneoBote)
End If


End If 'npos<>0

Exit Sub
fallo:
Call LogError("WarpUserChar " & Err.number & " D: " & Err.Description)

End Sub

Sub WarpUserChar2(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
On Error GoTo fallo
'Quitar el dialogo
Dim nPos As WorldPos
Dim xpos As WorldPos
Dim x1 As Byte
Dim Y1 As Byte
If UserIndex = 0 Or Map = 0 Then Exit Sub

'pluto:6.3 aleatorios zonas conflictivas

'pluto:6.4
If Map = MapaAngel Or (Map > 165 And Map < 170) Or Map = 185 Then
        If UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Then
        UserList(UserIndex).Stats.MinSta = 0
        End If
End If

'pluto:2.18----------------------------
xpos.Map = Map
xpos.Y = Y
xpos.X = X
Dim aguita As Byte
If UserList(UserIndex).flags.Navegando = 1 Then aguita = 1 Else aguita = 0
'pluto:6.0A-----------------------------------
'If UserList(UserIndex).flags.Privilegios = 0 Then
Call ClosestLegalPos(xpos, nPos, aguita)
'Else
'nPos.X = X
'nPos.Y = Y
'End If
'---------------------------------------------
If nPos.X <> 0 And nPos.Y <> 0 Then 'end if al final
X = nPos.X
Y = nPos.Y
'---------------------------------------
Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 21, UserList(UserIndex).Char.CharIndex)

Call SendData2(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, 5)

'pluto:2.7.1
'If Y > 90 Then Y = Y - 2


Dim Oldmap As Integer
Dim OldX As Integer
Dim OldY As Integer

Oldmap = UserList(UserIndex).Pos.Map
OldX = UserList(UserIndex).Pos.X
OldY = UserList(UserIndex).Pos.Y


Call EraseUserChar(ToMap, 0, Oldmap, UserIndex)

'pluto:6.2-----------------------
If Oldmap = 291 And UserList(UserIndex).flags.ParejaTorneo > 0 Then
UserList(UserList(UserIndex).flags.ParejaTorneo).flags.ParejaTorneo = 0
UserList(UserIndex).flags.ParejaTorneo = 0
End If

If Oldmap = 292 And UserList(UserIndex).flags.Privilegios = 0 Then
        If UserList(UserIndex).GuildInfo.GuildName = TorneoClan(1).Nombre Then
                TorneoClan(1).numero = TorneoClan(1).numero - 1
                        If TorneoClan(1).numero = 0 Then
                        TClanOcupado = TClanOcupado - 1
                        TorneoClan(1).Nombre = ""
                        End If
        ElseIf UserList(UserIndex).GuildInfo.GuildName = TorneoClan(2).Nombre Then
                TorneoClan(2).numero = TorneoClan(2).numero - 1
                        If TorneoClan(2).numero = 0 Then
                        TClanOcupado = TClanOcupado - 1
                        TorneoClan(2).Nombre = ""
                        End If
        End If
End If


UserList(UserIndex).Pos.X = X
UserList(UserIndex).Pos.Y = Y
tt:
UserList(UserIndex).Pos.Map = Map




If Oldmap <> Map Then
        Call SendData2(ToIndex, UserIndex, 0, 14, Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion)
        If MapInfo(Map).Terreno = "BOSQUE" Then
        MapInfo(Map).Music = "58-1"
        ElseIf MapInfo(Map).Terreno = "MAR" Then
        MapInfo(Map).Music = "74-1"
        End If
        'pluto:6.0a
        If MapInfo(Map).Music <> MapInfo(Oldmap).Music Then
        Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)
        End If
'Call SendData(ToIndex, UserIndex, 0, "TM" & 25)
        Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        
        Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

        'Update new Map Users
        'pluto:6.8
        If UserList(UserIndex).flags.Privilegios = 0 Then MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        
 'Update old Map Users
        If UserList(UserIndex).flags.Privilegios = 0 Then MapInfo(Oldmap).NumUsers = MapInfo(Oldmap).NumUsers - 1
          
        If MapInfo(Oldmap).NumUsers < 0 Then
                MapInfo(Oldmap).NumUsers = 0
        End If


        
'pluto:6.0A-------------solidos mapa 274--------------------------
If Map = 274 Then
Dim a As Byte
'Dim x As Byte
Dim b As Byte
Dim Salida As Byte
Dim obj As obj
If SolidoGirando = 0 Then GoTo nogiraba
'SolidoGirando = 5
b = 45 + (SolidoGirando * 2)
'quitamos al que gira-----------------
If MapData(Map, b, 26).OBJInfo.ObjIndex = 1170 + SolidoGirando Then
Call EraseObj(ToMap, UserIndex, Map, 10000, Map, b, 26)
 obj.Amount = 1
 obj.ObjIndex = 1175 + SolidoGirando
 Call MakeObj(ToMap, 0, Map, obj, Map, b, 26)
 Salida = 6 + (SolidoGirando * 10)
 MapData(Map, Salida, 11).TileExit.Map = 28
 MapData(Map, Salida, 11).TileExit.X = 46
 MapData(Map, Salida, 11).TileExit.Y = 86
 End If
'fin quitamos girar---------

'ponemos nuevo solido a girar------------
nogiraba:
a = RandomNumber(1, 5)
b = 45 + (a * 2)

If MapData(Map, b, 26).OBJInfo.ObjIndex = 1175 + a Then
Call EraseObj(ToMap, UserIndex, Map, 10000, Map, b, 26)
 obj.Amount = 1
 obj.ObjIndex = 1170 + a
 Call MakeObj(ToMap, 0, Map, obj, Map, b, 26)
 SolidoGirando = a
  Salida = 6 + (SolidoGirando * 10)
 MapData(Map, Salida, 11).TileExit.Map = 276
 MapData(Map, Salida, 11).TileExit.X = 43
 MapData(Map, Salida, 11).TileExit.Y = 83
 End If

End If 'mapa 274
'---------fin solidos------------------------------------------









'pluto:2.12
If MapInfo(Oldmap).NumUsers = 0 And Oldmap = MapaTorneo2 Then MinutoSinMorir = 0
If Oldmap = MapaTorneo2 Then
UserList(UserIndex).Torneo2 = 0
Torneo2Record = 0
Call SendData2(ToIndex, UserIndex, 0, 96, Torneo2Name & "," & Torneo2Record & "," & TorneoBote)
End If

Else 'mismo mapa
        
        Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

End If


Call UpdateUserMap(UserIndex)
'pluto:2-3-04
If FX And UserList(UserIndex).flags.Privilegios = 0 Then 'FX
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_WARP)
        Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If
'[MerLiNz:X]
If (UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
        Call SendData2(ToMap, 0, Map, 17, UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
        Call SendData2(ToIndex, UserIndex, 0, 17, UserList(MapData(Map, X, Y).UserIndex).Char.CharIndex & ",1")
End If
'[\END]
'pluto:6.2------------
'Call EfectoIncor(UserIndex)
UserList(UserIndex).flags.Incor = True
UserList(UserIndex).Counters.Incor = 0
'PLUTO:6.3---------------
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'--------------------------

'UserList(UserIndex).flags.Macreanda = 0
'Call SendData2(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 61 & "," & 1)
'UserList(UserIndex).Char.FX = 61
'-----------------------

Call WarpMascotas(UserIndex)

'pluto:2.12
If Map = MapaTorneo2 And UserList(UserIndex).flags.Privilegios = 0 And Oldmap <> MapaTorneo2 Then
If Torneo2Name = "" Then Torneo2Name = UserList(UserIndex).Name: Torneo2Record = 0
TorneoBote = TorneoBote + 100
Call SendData2(ToMap, 0, MapaTorneo2, 96, Torneo2Name & "," & Torneo2Record & "," & TorneoBote)

'Call SendData2(ToIndex, UserIndex, 0, 96, Torneo2Name & "," & Torneo2Record & "," & TorneoBote)
End If


End If 'npos<>0

Exit Sub
fallo:
Call LogError("WarpUserChar2 " & Err.number & " D: " & Err.Description)

End Sub

Sub WarpMascotas(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer

NroPets = UserList(UserIndex).NroMacotas

For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
                PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
                PetTypes(i) = UserList(UserIndex).MascotasType(i)
                PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
        End If
Next i

For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
                UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
                UserList(UserIndex).MascotasType(i) = PetTypes(i)
                'Controlamos que se sumoneo OK
                If UserList(UserIndex).MascotasIndex(i) = MAXNPCS Then
                                UserList(UserIndex).MascotasIndex(i) = 0
                                UserList(UserIndex).MascotasType(i) = 0
                                If UserList(UserIndex).NroMacotas > 0 Then UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
                                Exit Sub
                End If
                Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
                Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = SIGUE_AMO
                Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
                Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNpc = 0
                Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
                'pluto:6.0A
                If MapInfo(UserList(UserIndex).Pos.Map).Mascotas = 1 Then
                If Npclist(UserList(UserIndex).MascotasIndex(i)).NPCtype <> 60 Then Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 1
                End If
                
                Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
        End If
Next i

UserList(UserIndex).NroMacotas = NroPets


Exit Sub
fallo:
Call LogError("warpmascotas " & Err.number & " D: " & Err.Description)

End Sub


Sub RepararMascotas(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim i As Integer
Dim MascotasReales As Integer

For i = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
Next i

If MascotasReales <> UserList(UserIndex).NroMacotas Then UserList(UserIndex).NroMacotas = 0
Exit Sub
fallo:
Call LogError("repararmascotas " & Err.number & " D: " & Err.Description)

End Sub
Sub Cerrar_Usuario(ByVal UserIndex As Integer, Optional ByVal Tiempo As Integer = -1)
        CloseSocket (UserIndex)
        Exit Sub
  If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
        
   If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
                UserList(UserIndex).Counters.Saliendo = True
                UserList(UserIndex).Counters.Salir = IIf(UserList(UserIndex).flags.Privilegios > 0 Or Not MapInfo(UserList(UserIndex).Pos.Map).Pk, 0, Tiempo)
                
                
                Call SendData(ToIndex, UserIndex, 0, "||Cerrando...Se cerrará el juego en " & UserList(UserIndex).Counters.Salir & " segundos..." & "´" & FontTypeNames.FONTTYPE_info)
        
        'ElseIf Not UserList(UserIndex).Counters.Saliendo Then
        '    If NumUsers <> 0 Then NumUsers = NumUsers - 1
        '    Call SendData(ToIndex, UserIndex, 0, "||Gracias por jugar Argentum Online" & FONTTYPENAMES.FONTTYPE_INFO)
        '    Call SendData(ToIndex, UserIndex, 0, "FINOK")
        '
        '    Call CloseUser(UserIndex)
        '    UserList(UserIndex).ConnID = -1: UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
        '    frmMain.Socket2(UserIndex).Cleanup
        '    Unload frmMain.Socket2(UserIndex)
        '    Call ResetUserSlot(UserIndex)
        
        
        
        End If
End Sub
