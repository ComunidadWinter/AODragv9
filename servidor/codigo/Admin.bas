Attribute VB_Name = "ModAdmin"
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

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC


'TIMERS
Public NPC_AI As Integer

'INTERVALO
Public User_AtacarMelee As Integer
Public User_LanzarMagia As Integer


'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloOculto As Integer '[Nacho]
Public IntervaloGolpeUsar As Long
Public IntervaloMagiaGolpe As Long
Public IntervaloGolpeMagia As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long
Public IntervaloPuedeSerAtacado As Long
Public IntervaloPuedeCambiardeMapa As Long
Public IntervaloPuedeMakrear As Integer 'MaxTus
Public IntervaloMorphPJ As Integer
'BALANCE

Public PorcentajeRecuperoMana As Integer

Public MinutosWs As Long
Public Puerto As Integer

Public MultiplicadorGP As Double
Public MultiplicadorGPN As Double

Public BootDelBackUp As Byte
Public DeNoche As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
VersionOK = (Ver = ULTIMAVERSION)
End Function

Sub ReSpawnOrigPosNpcs()
On Error Resume Next

Dim i As Integer
Dim MiNPC As npc
   
For i = 1 To LastNPC
   'OJO
   If NPCList(i).flags.Active Then
        
        If InMapBounds(NPCList(i).Orig.Map, NPCList(i).Orig.X, NPCList(i).Orig.Y) And NPCList(i).Numero = Guardias Then
                MiNPC = NPCList(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        'tildada por sugerencia de yind
        'If NPCList(i).Contadores.TiempoExistencia > 0 Then
        '        Call MuereNpc(i, 0)
        'End If
   End If
   
Next i

End Sub

Sub WorldSave()
On Error Resume Next
'Call LogTarea("Sub WorldSave")

Dim LoopX As Integer
Dim Porc As Long

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))
    
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    
    Dim j As Integer, k As Integer
    
    For j = 1 To NumMaps
        If MapInfo(j).BackUp = 1 Then k = k + 1
    Next j

   ' frmCargando.cargar.min = 0
 '   frmCargando.cargar.max = k
  '  frmCargando.cargar.Value = 0

    For LoopX = 1 To NumMaps
        'DoEvents
        
        If MapInfo(LoopX).BackUp = 1 Then
        
                Call GrabarMapa(LoopX, App.Path & "\WorldBackUp\Mapa" & LoopX)
                FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
        End If
    
    Next LoopX
    
    FrmStat.Visible = False
    
    If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
    'If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")
    
    For LoopX = 1 To LastNPC
        If NPCList(LoopX).flags.BackUp = 1 Then
                Call BackUPnPc(LoopX)
        End If
    Next
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub PurgarPenas()
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call WriteConsoleMsg(i, "Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                    Call FlushBuffer(i)
                End If
            End If
        End If
    Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)
        
        UserList(UserIndex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
        If LenB(GmName) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
        End If
        
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean

BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean

PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'Unban the character
Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")

'Remove it from the banned people database
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function

Public Sub BanIpAgrega(ByVal ip As String)
    BanIPs.Add ip
    
    Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
Dim Dale As Boolean
Dim LoopC As Long

Dale = True
LoopC = 1
Do While LoopC <= BanIPs.Count And Dale
    Dale = (BanIPs.Item(LoopC) <> ip)
    LoopC = LoopC + 1
Loop

If Dale Then
    BanIpBuscar = 0
Else
    BanIpBuscar = LoopC - 1
End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

On Error Resume Next

Dim N As Long

N = BanIpBuscar(ip)
If N > 0 Then
    BanIPs.Remove N
    BanIpGuardar
    BanIpQuita = True
Else
    BanIpQuita = False
End If

End Function

Public Sub BanIpGuardar()
Dim ArchivoBanIp As String
Dim ArchN As Long
Dim LoopC As Long

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

ArchN = FreeFile()
Open ArchivoBanIp For Output As #ArchN

For LoopC = 1 To BanIPs.Count
    Print #ArchN, BanIPs.Item(LoopC)
Next LoopC

Close #ArchN

End Sub

Public Sub BanIpCargar()
Dim ArchN As Long
Dim Tmp As String
Dim ArchivoBanIp As String

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

Do While BanIPs.Count > 0
    BanIPs.Remove 1
Loop

ArchN = FreeFile()
Open ArchivoBanIp For Input As #ArchN

Do While Not EOF(ArchN)
    Line Input #ArchN, Tmp
    BanIPs.Add Tmp
Loop

Close #ArchN

End Sub

Public Sub ActualizaEstadisticasWeb()

Static Andando As Boolean
Static Contador As Long
Dim Tmp As Boolean

Contador = Contador + 1

If Contador >= 10 Then
    Contador = 0
    Tmp = EstadisticasWeb.EstadisticasAndando()
    
    If Andando = False And Tmp = True Then
        Call InicializaEstadisticas
    End If
    
    Andando = Tmp
End If

End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
'***************************************************
'Author: Unknown
'Last Modification: 03/02/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************
    If EsAdmin(Name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User
    End If
End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/02/07
'
'***************************************************
    Dim tUser As Integer
    Dim userPriv As Byte
    Dim cantPenas As Byte
    Dim Rank As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)
    
    Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_TALK)
            
            If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                userPriv = UserDarPrivilegioLevel(UserName)
                
                If (userPriv And Rank) > (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No podes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                        Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex, reason)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                        'ponemos el flag de ban a 1
                        Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                        'ponemos la pena
                        cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(reason) & " " & Date & " " & time)
                        
                        If (userPriv And Rank) = (.flags.Privilegios And Rank) Then
                            .flags.Ban = 1
                            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                            Call CloseSocket(bannerUserIndex)
                        End If
                        
                        Call LogGM(.Name, "BAN a " & UserName)
                    End If
                End If
            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No podes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call LogBan(tUser, bannerUserIndex, reason)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))
            
            'Ponemos el flag de ban a 1
            UserList(tUser).flags.Ban = 1
            
            If (UserList(tUser).flags.Privilegios And Rank) = (.flags.Privilegios And Rank) Then
                .flags.Ban = 1
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                Call CloseSocket(bannerUserIndex)
            End If
            
            Call LogGM(.Name, "BAN a " & UserName)
            
            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(reason) & " " & Date & " " & time)
            
            Call CloseSocket(tUser)
        End If
    End With
End Sub

Sub NpcAutoSacerdote(ByVal NPCIndex As Integer) ' GSZAO
'***************************************************
'Author: ^[GS]^
'Last Modification: 29/07/2012 - ^[GS]^
'***************************************************
    Dim X As Integer
    Dim Y As Integer
    Dim UserIndex As Integer
    
    With NPCList(NPCIndex)
        For Y = .Pos.Y - 5 To .Pos.Y + 5
             For X = .Pos.X - 5 To .Pos.X + 5
                 If MapData(.Pos.Map, X, Y).UserIndex > 0 Then
                    UserIndex = MapData(.Pos.Map, X, Y).UserIndex
                    If .NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex) Then
                        If iniSacerdoteCuraVeneno Then
                            If UserList(UserIndex).flags.Envenenado <> 0 Then
                                UserList(UserIndex).flags.Envenenado = 0
                            End If
                        End If
                        
                        If UserList(UserIndex).flags.Muerto = 1 Then
                            Call RevivirUsuario(UserIndex)
                            UserList(UserIndex).Stats.MinHP = VidaMaxima(UserIndex)
                            Call WriteUpdateHP(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "El Sacerdote levanta las manos, pronuncia unas palabras sagradas y tu cuerpo comienza a tomar forma. ¡Has resucitado!", FontTypeNames.FONTTYPE_INFOBOLD)
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(99, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 5, 0))
                            Exit Sub
                        End If
                        
                        If UserList(UserIndex).Stats.MinHP <> VidaMaxima(UserIndex) Then
                            UserList(UserIndex).Stats.MinHP = VidaMaxima(UserIndex)
                            Call WriteUpdateHP(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "El Sacerdote levanta las manos, pronuncia unas palabras sagradas y tus heridas comienzan a sanar rapidamente. ¡Has sido curado! ", FontTypeNames.FONTTYPE_INFOBOLD)
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(104, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 5, 0))
                            Exit Sub
                        End If
                    End If
                 End If
            Next X
        Next Y
    End With

End Sub

Public Function BanHD_rem(ByVal HD As String) As Boolean
' GSZ-AO - Remueve un SerialHD como baneado

   On Error Resume Next
 
    Dim N As Long
   
    N = BanHD_find(HD) ' buscar
    If N > 0 Then
        BanHDs.Remove N ' quitar
        BanHD_save ' guardar los cambios
        BanHD_rem = True
    Else
        BanHD_rem = False
    End If
   
End Function
Public Sub BanHD_add(ByVal HD As String)
' GSZ-AO - Agrega un nuevo SerialHD como baneado

    Dim N As Long
   
    N = BanHD_find(HD) ' buscar
    If N > 0 Then
        ' ya estaba
    Else
        BanHDs.Add HD ' agregar
        Call BanHD_save ' guardar los cambios
    End If
    
End Sub
Public Function BanHD_find(ByVal HD As String) As Long
' GSZ-AO - Busca si un SerialHD está baneado

    Dim Dale As Boolean
    Dim LoopC As Long
   
    Dale = True
    LoopC = 1
    Do While LoopC <= BanHDs.Count And Dale
        Dale = (BanHDs.Item(LoopC) <> HD)
        LoopC = LoopC + 1
    Loop
   
    If Dale Then
        BanHD_find = 0
    Else
        BanHD_find = LoopC - 1
    End If
    
End Function
Public Sub BanHD_save()
' GSZ-AO - Guarda el listado de SerialHD's baneados
On Error Resume Next
    Dim ArchivoBanHD As String
    Dim ArchN As Long
    Dim LoopC As Long
   
    ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"
       
    ArchN = FreeFile()
    Open ArchivoBanHD For Output As #ArchN
   
    For LoopC = 1 To BanHDs.Count
        Print #ArchN, BanHDs.Item(LoopC)
    Next LoopC
   
    Close #ArchN
   
End Sub
Public Sub BanHD_load()
' GSZ-AO - Carga el listado de SerialHD's baneados
On Error Resume Next

    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanHD As String
   
    ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"
   
    Set BanHDs = New Collection
   
    ArchN = FreeFile()
    Open ArchivoBanHD For Input As #ArchN
   
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanHDs.Add Tmp
    Loop
   
    Close #ArchN
End Sub

