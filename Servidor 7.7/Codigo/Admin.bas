Attribute VB_Name = "Admin"
Option Explicit

Public MaxLines As Integer
Public MOTD() As String

Public NPCs As Long
Public DebugSocket As Boolean

Public Horas As Long
Public Dias As Long
Public MinsRunning As Long


Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloParalisisPJ As Integer
Public IntervaloMorphPJ As Integer
Public Intervaloceguera As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloMover As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloUserPuedeAtacar As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
'pluto:2.17
Public TimeEmbarazo As Integer
Public TimeAborto As Integer
Public ProbEmbarazo As Integer
'pluto:2.8.0
Public IntervaloUserPuedeFlechas As Long
Public IntervaloRegeneraVampiro As Integer
'pluto:2.10
Public IntervaloUserPuedeTomar As Long

Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long
Public MinutosWs As Long
Public MinutosGp As Long


Public MAXPASOS As Long

Public BootDelBackUp As Byte
Public Lloviendo As Boolean

Public IpList As New Collection
Public ClientsCommandsQueue As Byte

Function VersionOK(ByVal Ver As String) As Boolean
  On Error GoTo fallo
    VersionOK = (Ver = ULTIMAVERSION)
'quitar esto
'VersionOK = True
Exit Function
fallo:
Call LogError("VERIONOK " & Err.number & " D: " & Err.Description)

End Function

Sub ReSpawnOrigPosNpcs()
On Error GoTo fallo

Dim i As Integer
Dim MinPc As npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
      'pluto:6.0----------------------
     ' If Npclist(i).Raid > 0 Then
      'Npclist(i).Stats.MinHP = Npclist(i).Stats.MaxHP
     ' End If
       '-----------------------------
        
        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).numero = Guardias Then
                MinPc = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MinPc)
        End If
        
        If Npclist(i).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(i, 0)
        End If
   End If
   
Next i

Exit Sub
fallo:
Call LogError("RESPAWNORIGPOSNPCS " & Err.number & " D: " & Err.Description)


End Sub
Sub ReSpawnCambioGuardias()
On Error GoTo fallo

Dim i As Integer
Dim MinPc As npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And (Npclist(i).numero = Guardias Or Npclist(i).numero = 115) Then
                MinPc = Npclist(i)
                Call QuitarNPC(i)
                'guau
                If MapInfo(MinPc.Pos.Map).Dueño = 2 Then MinPc.numero = 115 Else MinPc.numero = 6
                
                Call ReSpawnNpc(MinPc)
        End If
        
        If Npclist(i).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(i, 0)
        End If
   End If
   
Next i

Exit Sub
fallo:
Call LogError("RESPAWNORIGPOSNPCS " & Err.number & " D: " & Err.Description)


End Sub
Sub WorldSave()
On Error GoTo fallo

Dim loopX As Integer
Dim Porc As Long

'Call SendData(ToAll, 0, 0, "||%%%%POR FAVOR ESPERE, INICIANDO WORLDSAVE%%%%" & FONTTYPENAMES.FONTTYPE_INFO)
 Call SendData(ToAll, 0, 0, "L8")
Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

Dim j As Integer, k As Integer

For j = 1 To NumMaps
    If MapInfo(j).BackUp = 1 Then k = k + 1
Next j

FrmStat.ProgressBar1.Min = 0
FrmStat.ProgressBar1.max = k
FrmStat.ProgressBar1.value = 0

For loopX = 1 To NumMaps
    DoEvents
    'quitar esto
   ' MapInfo(215).BackUp = 1
    'MapInfo(234).BackUp = 1
   ' MapInfo(97).BackUp = 1
   ' MapInfo(266).BackUp = 1
    If MapInfo(loopX).BackUp = 1 Then
    
            Call SaveMapData(loopX)
            'FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
    End If

Next loopX

FrmStat.Visible = False

If FileExist(DatPath & "\bkNPCs.dat", vbNormal) Then Kill (DatPath & "bkNPCs.dat")
If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

For loopX = 1 To LastNPC
    If Npclist(loopX).flags.BackUp = 1 Then
            Call BackUPnPc(loopX)
    End If
Next
 Call SendData(ToAll, 0, 0, "L9")
'Call SendData(ToAll, 0, 0, "||%%%%WORLDSAVE DONE%%%%" & FONTTYPENAMES.FONTTYPE_INFO)

Exit Sub
fallo:
Call LogError("WORLDSAVE " & Err.number & " D: " & Err.Description)


End Sub

Public Sub PurgarPenas()
 On Error GoTo fallo
Dim i As Integer
For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
        If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
   'pluto:2.18
    If EsNewbie(i) Then

Select Case UCase$(UserList(i).raza)
Case "ORCO"
Call WarpUserChar(i, Pobladoorco.Map, Pobladoorco.X, Pobladoorco.Y, True)
Case "HUMANO"
Call WarpUserChar(i, Pobladohumano.Map, Pobladohumano.X, Pobladohumano.Y, True)
Case "CICLOPE"
Call WarpUserChar(i, Pobladohumano.Map, Pobladohumano.X, Pobladohumano.Y, True)

Case "ELFO"
Call WarpUserChar(i, Pobladoelfo.Map, Pobladoelfo.X, Pobladoelfo.Y, True)
Case "ELFO OSCURO"
Call WarpUserChar(i, Pobladoelfo.Map, Pobladoelfo.X, Pobladoelfo.Y, True)
Case "VAMPIRO"
Call WarpUserChar(i, Pobladovampiro.Map, Pobladovampiro.X, Pobladovampiro.Y, True)
Case "ENANO"
Call WarpUserChar(i, Pobladoenano.Map, Pobladoenano.X, Pobladoenano.Y, True)
Case "GNOMO"
Call WarpUserChar(i, Pobladoenano.Map, Pobladoenano.X, Pobladoenano.Y, True)
Case "GOBLIN"
Call WarpUserChar(i, Pobladoenano.Map, Pobladoenano.X, Pobladoenano.Y, True)

End Select

    Else
Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
Call SendData(ToIndex, i, 0, "||Has sido liberado!" & "´" & FontTypeNames.FONTTYPE_info)
    End If

End If
End If
End If

Next i
Exit Sub
fallo:
Call LogError("PULGAR PENAS " & Err.number & " D: " & Err.Description)


End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = "")
 On Error GoTo fallo
      
      'PLUTO:2.18-------------------------
      'If EsNewbie(UserIndex) Then Exit Sub
      '-----------------------------------
        UserList(UserIndex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
        If GmName = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & "´" & FontTypeNames.FONTTYPE_info)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & "´" & FontTypeNames.FONTTYPE_info)
        End If
        Call SendData(ToIndex, UserIndex, 0, "TW" & 179)
Exit Sub
fallo:
Call LogError("ENCARCELAR " & Err.number & " D: " & Err.Description)
    
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
 On Error GoTo fallo
    BANCheck = (val(GetVar(CharPath & Left$(Name, 1) & "\" & Name & ".chr", "FLAGS", "Ban")) = 1)

Exit Function
fallo:
Call LogError("BANCHECK " & Err.number & " D: " & Err.Description)


End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean
    On Error GoTo fallo
    If Not FileExist(CharPath & Left$(Name, 1), vbDirectory) Then MkDir (CharPath & Left$(Name, 1))
    PersonajeExiste = FileExist(CharPath & Left$(Name, 1) & "\" & UCase$(Name) & ".chr", vbArchive)
Exit Function
fallo:
Call LogError("PERSONAJEEXISTE " & Err.number & " D: " & Err.Description)


End Function


Public Function UnBan(ByVal Name As String) As Boolean
 On Error GoTo fallo
'Unban the character
Call WriteVar(CharPath & Left$(Name, 1) & "\" & Name & ".chr", "FLAGS", "Ban", "0")
'Remove it from the banned people database
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NOONE")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Fecha", "NOONE")

Exit Function
fallo:
Call LogError("UNBAN " & Err.number & " D: " & Err.Description)

End Function
