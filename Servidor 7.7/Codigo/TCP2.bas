Attribute VB_Name = "tcp2"
Private Sub IniDeleteSection(ByVal sIniFile As String, ByVal sSection As String)
    Call writeprivateprofilestring(sSection, 0&, 0&, sIniFile)
End Sub
Sub TCP2(ByVal UserIndex As Integer, ByVal rdata As String)

On Error GoTo ErrorComandoPj:
Dim LC As Byte
Dim tot As Integer
Dim sndData As String
Dim CadenaOriginal As String
Dim Moverse As Byte
Dim loopc As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim Tindex As Integer
Dim tName As String
Dim tNome As String
Dim tpru As String
Dim tMessage As String
Dim auxind As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim pass As String
Dim Mapa As Integer
Dim Name As String
Dim ind
Dim n As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim HayGM As Boolean
Dim GM1 As String
'pluto:6.0A
CadenaOriginal = rdata
If rdata = "" Then Exit Sub
'pluto:2.10
'¿Tiene un indece valido?
If UserIndex <= 0 Then
    Call CloseSocket(UserIndex)
    Call LogError(Date & " Userindex no válido")
    Exit Sub
End If
'¿Está logeado?
If UserList(UserIndex).flags.UserLogged = False Then
    'Call LogError(Date & " We: " & UserList(UserIndex).ip & " / " & Cuentas(UserIndex).mail)
'pluto:2.19 añade true
    Call CloseSocket(UserIndex, True)
    Exit Sub
End If

'Delzak sos offline
'If (Left$(rdata, 5)) = "/DAME" Then
 '   rdata = Right$(rdata, Len(rdata) - 5)
  '  Dim M As String
   ' M = Ayuda.Busca(val(rdata), UserIndex) & Ayuda.CuantasVecesAparece(UserIndex)
    'Call SendData2(ToIndex, UserIndex, 0, 111, M)
    'Exit Sub
'End If

If UCase$(Left$(rdata, 4)) = "/GM " Then

Dim rdata2 As String
rdata = Right$(rdata, Len(rdata) - 4)
        If (Len(rdata) < 1) Then
            Call SendData(ToIndex, UserIndex, 0, "||Utiliza: /GM motivo" & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
        'pluto:2.15
        rdata2 = rdata
        rdata = " (" & UserList(UserIndex).Stats.ELV & "):" & rdata
        
        If Not Ayuda.Existe(UserList(UserIndex).Name) Then
            Call SendData(ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & "´" & FontTypeNames.FONTTYPE_info)
            Call Ayuda.Push(rdata, UserList(UserIndex).Name & ";" & rdata)
           'pluto:6.0A
            Call SendData(ToAdmins, UserIndex, 0, "|| SOS de " & UserList(UserIndex).Name & ": " & rdata2 & "´" & FontTypeNames.FONTTYPE_talk)
Exit Sub
        Else
            Call Ayuda.Quitar(UserList(UserIndex).Name)
            Call Ayuda.Push(rdata, UserList(UserIndex).Name & ";" & rdata)
            Call SendData(ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & "´" & FontTypeNames.FONTTYPE_info)
        End If
        Exit Sub
        End If

'nati: /SMSUSER NICK#ASUNTO#MENSAJE
'@Nati: wwww.juegosdrag.es - 2011
If UCase(Left(rdata, 9)) = "/SMSUSER " Then
Dim smsSuma As String
Dim smsResta As String
Dim asunto As String
Dim mensaje As String
'Call SendData(ToIndex, UserIndex, 0, "|| HOLAHOLA " & rdata & "´" & FontTypeNames.FONTTYPE_info)
rdata = Right$(rdata, Len(rdata) - 9)
nick = ReadField(1, rdata, 35)
asunto = ReadField(2, rdata, 35)
mensaje = ReadField(3, rdata, 35)
If Not PersonajeExiste(nick) Then
Call SendData(ToIndex, UserIndex, 0, "||El personaje no existe." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
If Not FileExist(App.Path & "\MAIL\" & Left$(UCase$(nick), 1), vbDirectory) Then
'cambiamos el esto: antes era: Call MkDir(App.Path & "\MAIL\" & Left$(UCase$(nick), 1))
    MkDir (App.Path & "\MAIL\" & Left$(UCase$(nick), 1))
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "SMS", "0")
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "BAN", "0")
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "AVISO", "0")
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "FECHA", "1")
End If
'If Not FileExist("\MAIL\" & nick & Left$(nick, 1) & "\" & ".MAIL", vbArchive) Then
'Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "SMS", 0)
'End If
smsResta = GetVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "SMS")
smsSuma = val(smsResta) + 1
If smsResta = 25 Then
Call SendData(ToIndex, UserIndex, 0, "||El personaje tiene la bandeja llena, no puedes enviarle mensajes." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", este, "Reason", Name)
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "SMS", smsSuma)
'Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", este, "Reason", Name)
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "MENSAJE" & smsSuma, "DE", UserList(UserIndex).Name)
'Call WriteVar(App.Path & "\Ubicación en la carpeta\" & "Nombre de archivo" & ".tipo de archivo", "Contenido", "Contenido1", Text1.Text)
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "MENSAJE" & smsSuma, "ASUNTO", asunto)
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "MENSAJE" & smsSuma, "FECHA", Format(Now, "dd/mm/yy"))
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "MENSAJE" & smsSuma, "MENSAJE", mensaje)
bansms = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "BAN")
If bansms = 1 Then
Exit Sub
End If
Call SendData(ToIndex, UserIndex, 0, "||Mensaje Enviado" & "´" & FontTypeNames.FONTTYPE_info)
smsmensaje = GetVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "MENSAJE" & smsSuma, "mensaje")
'Call SendData(ToIndex, UserIndex, 0, "||Mensaje: " & smsmensaje & "´" & FontTypeNames.FONTTYPE_info)
Tindex = NameIndex(nick)
If Tindex = 0 Then
Else
Call SendData2(ToIndex, Tindex, 0, 114)
End If
End If

'@Nati: wwww.juegosdrag.es - 2011
If UCase$(rdata) = "/SMSREFRESH" Then
Dim mensajes As String
Dim fecha As String
Dim Nombre As String
Dim asuntosms As String
Dim mensajesx As String
Dim fechax As String
Dim nombrex As String
Dim asuntosmsx As String
Dim smsTOTAL As String
If Not FileExist(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1), vbDirectory) Then
    Call MkDir(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1))
End If
If Not FileExist(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", vbArchive) Then
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "SMS", "0")
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "BAN", "0")
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "AVISO", "0")
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "FECHA", "1")
End If
'If Not FileExist("\MAIL\" & nick & Left$(UserList(UserIndex).Name, 1) & "\" & ".MAIL", vbArchive) Then
'Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "SMS", 0)
'End If
smsTOTAL = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "SMS")
For natillas = 1 To smsTOTAL
Nombre = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & natillas, "DE")
asuntosms = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & natillas, "ASUNTO")
mensajes = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & natillas, "MENSAJE")
fecha = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & natillas, "FECHA")
'seña = "#"
'Call SendData2(ToIndex, UserIndex, 0, "xD")
'Call SendData2(ToIndex, UserIndex, 0, 112 & nombre & seña & asuntosms & seña & mensajes & seña & fecha)
Call SendData2(ToIndex, UserIndex, 0, 112, Nombre & "#" & asuntosms & "#" & mensajes & "#" & fecha & "#" & natillas)
Next
End If

'@Nati: wwww.juegosdrag.es - 2011
If UCase$(Left(rdata, 8)) = "/SMSPAM " Then
'Exit Sub
Dim avisojj As String

rdata = Right$(rdata, Len(rdata) - 8)
nick = ReadField(1, rdata, 35)
asunto = ReadField(2, rdata, 35)
mensaje = ReadField(3, rdata, 35)
avisojj = GetVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "AVISO")
avisojj = avisojj + 1
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "AVISO", avisojj)

Dim SMSPAM As Integer
SMSPAM = FreeFile ' obtenemos un canal
Open App.Path & "\logs\mensajesSPAM.log" For Append As #SMSPAM
Print #SMSPAM, "-----------------------------------"
Print #SMSPAM, "Usuario denunciado: " & nick
Print #SMSPAM, "Asunto: " & asunto
Print #SMSPAM, "Asunto: " & mensaje
Print #SMSPAM, "Por: " & UserList(UserIndex).Name
Print #SMSPAM, "-----------------------------------"
Close #SMSPAM
smsResta = GetVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "SMS")
smsSuma = val(smsResta) + 1
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "SMS", smsSuma)
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "SMS", smsSuma)
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "MENSAJE" & smsSuma, "DE", "AODragbot")
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "MENSAJE" & smsSuma, "ASUNTO", "Denuncia")
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "MENSAJE" & smsSuma, "FECHA", Format(Now, "dd/mm/yy"))
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "MENSAJE" & smsSuma, "MENSAJE", "Has sido denunciado por el usuario: " & UserList(UserIndex).Name & " Tienes: " & avisojj & " de denuncias.")

If avisojj > 15 Then
Dim fechatrucha As String
fechoy = Format(Now, "dd/mm/yy")
fechatrucha = 7 + (Left(fechoy, 2))
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "BAN", "1")
Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "FECHA", fechatrucha)
End If
fechaban = GetVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "FECHA")
If fechaban = 0 Then
Else
'If fechaban - Format(Now, "dd/mm/yy") = 0 Then
'Call WriteVar(App.Path & "\MAIL\" & Left$(nick, 1) & "\" & nick & ".MAIL", "INFO", "FECHA", "0")
'End If
End If
End If

'@Nati: wwww.juegosdrag.es - 2011
If UCase(rdata) = "/DESPAM" Then
fechaban = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "FECHA")
If fechaban = 1 Then
Exit Sub
Else
fechoy = Format(Now, "dd/mm/yy")
fecharesta = (Left(fechaban, 2)) - (Left(fechoy, 2))
If fecharesta = "0" Or fechaban > fechoy Then
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "FECHA", "0")
Call SendData(ToIndex, UserIndex, 0, "||Has sido desbaneado, ya puedes usar el sistema de mensajeria." & "´" & FontTypeNames.FONTTYPE_info)
End If
End If
End If

'@Nati: wwww.juegosdrag.es - 2011
'@Nati: Comando muy costoso :(
If UCase(Left(rdata, 9)) = "/SMSKILL " Then
rdata = Right$(rdata, Len(rdata) - 9)
Dim smsALL As String
Dim smsREM As String
'SMS TOTALES AHORA
smsALL = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "SMS")
'SMS TOTALES DESPUES
smsREM = val(smsALL) - 1
'SMS OK
'smsOK = WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "SMS", smsREM)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "INFO", "SMS", smsREM)
If smsALL = rdata Then
sFicINI = App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL"
file = App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL"
sSeccion = "MENSAJE" & rdata
IniDeleteSection sFicINI, sSeccion
Exit Sub
End If
'ESTRUCTURA DEL MENSAJE FUERA
sFicINI = App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL"
file = App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL"
sSeccion = "MENSAJE" & rdata
IniDeleteSection sFicINI, sSeccion
Call SendData(ToIndex, UserIndex, 0, "||El mensaje ha sido borrado con exito." & "´" & FontTypeNames.FONTTYPE_info)
'AQUI ORGANIZAMOS LOS MENSAJES.
If smsALL < 1 Then Exit Sub
For n = 1 To smsALL
nombrex = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "DE")
asuntosmsx = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "ASUNTO")
mensajesx = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "MENSAJE")
fechax = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "FECHA")
sFicINI = App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL"
file = App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL"
sSeccion = "MENSAJE" & n
IniDeleteSection sFicINI, sSeccion
DoEvents
If n = rdata Then
borranormal = False
borramensajenulo = True
If n = 1 Then
n = n + 1
cambion = True
End If
End If
If n - 1 = 0 Then
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "DE", nombrex)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "ASUNTO", asuntosmsx)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "FECHA", fechax)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "MENSAJE", mensajesx)
borranormal = True
Else
If borranormal = True Then
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "DE", nombrex)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "ASUNTO", asuntosmsx)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "FECHA", fechax)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "MENSAJE", mensajesx)
End If
If borra2 = True Then
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n - 1, "DE", nombrex)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n - 1, "ASUNTO", asuntosmsx)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n - 1, "FECHA", fechax)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n - 1, "MENSAJE", mensajesx)
'borranormal2 = True
End If
If borramensajenulo = True Then
If cambion = True Then
n = n - 1
cambion = False
End If
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "DE", nombrex)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "ASUNTO", asuntosmsx)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "FECHA", fechax)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n, "MENSAJE", mensajesx)
sFicINI = App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL"
file = App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL"
sSeccion = "MENSAJE" & n
IniDeleteSection sFicINI, sSeccion
borra2 = True
borramensajenulo = False
End If
If borranormal2 = True Then
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n - 1, "DE", nombrex)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n - 1, "ASUNTO", asuntosmsx)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n - 1, "FECHA", fechax)
Call WriteVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & n - 1, "MENSAJE", mensajesx)
borra2 = False
End If
End If
Next n
DoEvents
End If
'@Nati: wwww.juegosdrag.es - 2011
If UCase(Left(rdata, 9)) = "/SMSREAD " Then
rdata = Right$(rdata, Len(rdata) - 9)
numero = rdata
Nombre = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & numero, "DE")
asuntosms = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & numero, "ASUNTO")
fecha = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & numero, "FECHA")
mensaje = GetVar(App.Path & "\MAIL\" & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".MAIL", "MENSAJE" & numero, "MENSAJE")
Call SendData2(ToIndex, UserIndex, 0, 113, Nombre & "#" & asuntosms & "#" & mensaje)
End If

'Delzak SOS offline---------------------

    'Compruebo si hay gms online
   ' HayGM = False
   ' For loopc = 1 To LastUser
    '        If (UserList(loopc).Name <> "") And UserList(loopc).flags.Privilegios <> 0 Then
     '           GM1 = UCase(UserList(loopc).Name)
      '          If GM1 <> "AODRAGBOT" Then HayGM = True
       '     End If
    'Next
                    
    'If HayGM = False Then
     '       Call SendData(ToIndex, UserIndex, 0, "||En estos momentos no hay ningún gm online, su duda se ha guardado en el buzón de correo del Staff. La próxima vez que conecte un GM le responderá, para ver si su duda ha sido respondida, pulse F12")
      '      Call SendData(ToIndex, UserIndex, 0, "||Recuerda que si no has explicado bien tu duda, posiblemente el GM no podrá contestarte. Es aconsejable que si no la has explicado bien, vuelvas a mandarla pero esta vez bien explicada")
    'End If
'-----------------
        'If Not Ayuda.ExisteDelzak(UserList(UserIndex).Name) Then
         '   Call SendData(ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & "´" & FontTypeNames.FONTTYPE_info)
          '  Call Ayuda.Escribe(UserList(UserIndex).Stats.ELV, UserList(UserIndex).Name, rdata)
          ' Call Ayuda.Push(rdata, UserList(UserIndex).Name & ";" & rdata)
           'pluto:6.0A
           ' Call SendData(ToGM, UserIndex, 0, "|| SOS de " & UserList(UserIndex).Name & ": " & rdata2 & "´" & FontTypeNames.FONTTYPE_talk)
'Exit Sub
        'Else
           ' Call Ayuda.Borra(UserList(UserIndex).Name)
 '           Call Ayuda.Escribe(UserList(UserIndex).Stats.ELV, UserList(UserIndex).Name, rdata)
  '          Call SendData(ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & "´" & FontTypeNames.FONTTYPE_info)
   '     End If
    '    Exit Sub
'-------------------------------------


If UCase$(Left$(rdata, 5)) = "/BUG " Then
        n = FreeFile
        Open App.Path & "\BUGS\BUGs.log" For Append Shared As n
        Print #n, "--------------------------------------------"
        Print #n, "Usuario:" & UserList(UserIndex).Name & "  Fecha:" & Date & "    Hora:" & Time
        Print #n, "BUG:"
        Print #n, Right$(rdata, Len(rdata) - 5)
        Close #n
Call SendData(ToIndex, UserIndex, 0, "|| Entregado mensaje de BUG: " & Right$(rdata, Len(rdata) - 5) & " .Muchas Gracias por tu Colaboración." & "´" & FontTypeNames.FONTTYPE_info)

'pluto:2.17
Tindex = NameIndex("AoDraGBoT")
If Tindex <= 0 Then Exit Sub
Call SendData(ToIndex, Tindex, 0, "|| BUG: " & UserList(UserIndex).Name & " " & Right$(rdata, Len(rdata) - 5) & "´" & FontTypeNames.FONTTYPE_talk)
Exit Sub
 End If
   'pluto.6.2
If UCase$(Left$(rdata, 7)) = "/MACRO " Then
rdata = Right$(rdata, Len(rdata) - 7)
    If UserList(UserIndex).flags.ComproMacro = 0 Then Exit Sub
    If CodigoMacro = val(rdata) Then
    Call SendData(ToIndex, UserIndex, 0, "||Código Correcto. Muchas Gracias!!" & "´" & FontTypeNames.FONTTYPE_talk)
    UserList(UserIndex).flags.ComproMacro = 0
    'COMPROBANDOMACRO = False
    
    Else
    Call SendData(ToIndex, UserIndex, 0, "||Código Incorrecto !!" & "´" & FontTypeNames.FONTTYPE_talk)
    End If
Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/DESC " Then

        rdata = Right$(rdata, Len(rdata) - 6)
        If Not AsciiDescripcion(rdata) Then
            Call SendData(ToIndex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
        UserList(UserIndex).Desc = Trim$(rdata)
        Call SendData(ToIndex, UserIndex, 0, "||La descripción ha cambiado." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
End If
    
If UCase$(Left$(rdata, 6)) = "/VOTO " Then
 
            rdata = Right$(rdata, Len(rdata) - 6)
            Call ComputeVote(UserIndex, rdata)
            Exit Sub
End If
            
If UCase$(Left$(rdata, 8)) = "/REMORT " Then
'nati: durante la beta no tendremos el remort, en la oficial sacaremos una expansion donde habilitaremos el remort
'pero aun hay que tocarlo, queda pendiente.
'Exit Sub
        'pluto:2-3-04
        If TieneObjetos(882, 1, UserIndex) Then
            Call DoRemort(Right$(rdata, Len(rdata) - 8), UserIndex)
        Else
            Call SendData(ToIndex, UserIndex, 0, "|| No tienes Amuleto Ankh." & "´" & FontTypeNames.FONTTYPE_info)
        End If
      Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/PASSWD " Then

        rdata = Right$(rdata, Len(rdata) - 8)
        If Len(rdata) < 6 Then
             Call SendData(ToIndex, UserIndex, 0, "||El password debe tener al menos 6 caracteres." & "´" & FontTypeNames.FONTTYPE_info)
        Else
             Call SendData(ToIndex, UserIndex, 0, "||El password ha sido cambiado." & "´" & FontTypeNames.FONTTYPE_info)
             Cuentas(UserIndex).passwd = rdata
        End If
        Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/RETIRAR " Then
'RETIRA ORO EN EL BANCO
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
         End If
         'Se asegura que el target es un npc
         If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "L4")
              Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 9)
         If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
         Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
         If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "L2")
              Exit Sub
         End If
         If Not PersonajeExiste(UserList(UserIndex).Name) Then
              Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
              CloseUser (UserIndex)
              Exit Sub
         End If
       'pluto:2.19
       
         If val(rdata) >= 1 And Int(val(rdata)) <= UserList(UserIndex).Stats.Banco Then
              UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - Int(val(rdata))
              'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Int(val(rdata))
              Call AddtoVar(UserList(UserIndex).Stats.GLD, Int(val(rdata)), MAXORO)

              Call SendData(ToIndex, UserIndex, 0, "||6°Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
         Else
              Call SendData(ToIndex, UserIndex, 0, "||6°No tenes esa cantidad.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
         End If
         Call SendUserStatsOro(val(UserIndex))
         Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/DEPOSITAR " Then
'DEPOSITAR ORO EN EL BANCO
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "L4")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 11)
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
        Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "L2")
              Exit Sub
        End If
        If Int(val(rdata)) >= 1 And Int(val(rdata)) <= UserList(UserIndex).Stats.GLD Then
              'UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + Int(val(rdata))
              Call AddtoVar(UserList(UserIndex).Stats.Banco, Int(val(rdata)), MAXORO)

              
              UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Int(val(rdata))
              Call SendData(ToIndex, UserIndex, 0, "||6°Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        Else
              Call SendData(ToIndex, UserIndex, 0, "||6°No tenes esa cantidad.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        End If
        Call SendUserStatsOro(val(UserIndex))
        Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/PAGAR " Then
'cambiar exp por oro
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
        'comprueba level
        If UserList(UserIndex).Stats.ELV < 18 Then
       Call SendData(ToIndex, UserIndex, 0, "||6°Necesitas ser Level 18 o superior para comprender mis enseñanzas.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
                       
        Exit Sub
        End If
        
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "L4")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 7)
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_EXP _
        Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "L2")
              Exit Sub
        End If
        If CLng(val(rdata)) > 0 And CLng(val(rdata)) <= UserList(UserIndex).Stats.GLD Then
             UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + CLng(val(rdata) / 2)
             UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
              Call SendData(ToIndex, UserIndex, 0, "||°6Has subido " & CLng(val(rdata) / 2) & " puntos de experiencia." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        Call CheckUserLevel(UserIndex)
        
        Else
              Call SendData(ToIndex, UserIndex, 0, "||6°No tenes esa cantidad.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        End If
        Call SendUserStatsOro(val(UserIndex))
        Call SendUserStatsEXP(val(UserIndex))
        Exit Sub
End If
'pluto:7.0
    'Case "/BOVEDA"
    'pluto:7.0 cajas
If UCase$(Left$(rdata, 7)) = "/BOVEDA" Then
rdata = Right$(rdata, Len(rdata) - 7)
    
    
    
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
              If UserList(UserIndex).flags.Navegando = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Deja de Navegar!!" & "´" & FontTypeNames.FONTTYPE_info)
                 Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If UserList(UserIndex).flags.TargetNpc > 0 Then
              If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
              End If
            
              '------------------------
              
              'pluto:7.0
              If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 4 Or Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 25 Then
                'meto en Ncaja el número de la caja
                UserList(UserIndex).flags.NCaja = val(rdata)
                If Cuentas(UserIndex).Cajas > val(rdata) Or Cuentas(UserIndex).Cajas = val(rdata) Then
                Call IniciarDeposito(UserIndex)
              Else
              Call SendData(ToIndex, UserIndex, 0, "||Tienes " & Cuentas(UserIndex).Cajas & " baúles disponibles, para comprar mas dirigete a http://www.juegosdrag.es sección DragCréditos. " & "´" & FontTypeNames.FONTTYPE_info)
                Exit Sub
              End If
              End If
        Else
          Call SendData(ToIndex, UserIndex, 0, "L4")
        End If
        Exit Sub
        End If





If UCase$(Left$(rdata, 9)) = "/APOSTAR " Then
'casino
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "L4")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 9)
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_CASINO _
        Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "L2")
              Exit Sub
        End If
        
        If val(rdata) >= 1 And val(rdata) < 1001 And val(rdata) <= UserList(UserIndex).Stats.GLD Then
        Dim res As Integer
        Dim ros As Integer
        Dim casino As Integer
        
        res = RandomNumber(1, 1000)
        ros = RandomNumber(1, 40)
          If res > 998 Then casino = 100
          If res > 990 And res < 999 Then casino = 10
          If res > 970 And res < 991 Then casino = 5
          If res > 900 And res < 971 Then casino = 2
          If res > 700 And res < 901 Then casino = 1
          If res < 701 Then casino = 0
        If res > 998 And ros = 5 Then casino = 1000
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - CLng(val(rdata))
        If casino > 0 Then
        'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + CLng((val(rdata) * casino))
           Call AddtoVar(UserList(UserIndex).Stats.GLD, CLng((val(rdata) * casino)), MAXORO)
        Call SendData(ToIndex, UserIndex, 0, "||6°Has apostado " & CLng(val(rdata)) & " y Has GANADO " & CLng(val(rdata) * casino) & " Monedas de oro.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW176")

        End If
        If casino = 1000 Then
        Call SendData(ToAll, 0, 0, "||NOTICIA DE AODRAG: " & UserList(UserIndex).Name & " acaba de ganar su apuesta x1000 !!!!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
        Call SendData(ToAll, 0, 0, "TW" & SND_DINERO)
        Call LogCasino("Jugador:" & UserList(UserIndex).Name & "  Premio:x" & casino & "  Apostó:" & CLng(val(rdata)) & "  Ganó:" & CLng(val(rdata) * casino))

        End If
        
        If casino = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||6°Has apostado " & CLng(val(rdata)) & " y Has pérdido " & CLng(val(rdata)) & " Monedas de oro.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        Call SendData(ToIndex, UserIndex, 0, "TW" & SND_DINERO)
        End If
        
        Else
              Call SendData(ToIndex, UserIndex, 0, "||6°No puedes apostar esa cantidad.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        End If
        Call SendUserStatsOro(val(UserIndex))
        Exit Sub
End If


If UCase$(Left$(rdata, 6)) = "/CLAN " Then
    'hablar al clan
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "L3")
        Exit Sub
    End If
    If UserList(UserIndex).GuildInfo.GuildName = "" Then
        Call SendData(ToIndex, UserIndex, 0, "||No perteneces a ningún clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
    If UserList(UserIndex).Stats.GLD > 249 Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 250
        Call SendUserStatsOro(UserIndex)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||No tienes 250 oros para mandar mensaje. " & rdata & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
    rdata = Right$(rdata, Len(rdata) - 6)
    If rdata <> "" Then
        Call SendData(ToGuildMembers, UserIndex, 0, "|,[" & UserList(UserIndex).Name & "]: " & rdata & "´" & FontTypeNames.FONTTYPE_guildmsg)
'pluto:2-3-04
        If UCase$(Cotilla) = UCase$(UserList(UserIndex).GuildInfo.GuildName) Then
            Call SendData(ToGM, UserIndex, 0, "||" & UserList(UserIndex).Name & ": " & rdata & "´" & FontTypeNames.FONTTYPE_GUILD)
        End If
        End If
        
    Exit Sub
End If

If UCase$(Left$(rdata, 3)) = "/P " Then
rdata = Right$(rdata, Len(rdata) - 3)
If rdata = "" Then Exit Sub
    'hablar party
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "L3")
        Exit Sub
    End If
    If UserList(UserIndex).flags.party = False Then
        Call SendData(ToIndex, UserIndex, 0, "||No perteneces a ningúna party." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If rdata <> "" Then
        'Call SendData(toAL, 0, 0, "|*[" & UserList(UserIndex).Name & "]: " & rdata & "´" & FontTypeNames.FONTTYPE_GLOBAL)
        Call SendData(toParty, UserIndex, 0, "º;" & "[" & UserList(UserIndex).Name & "]: " & rdata & "´" & FontTypeNames.FONTTYPE_PARTY)
    End If
    
    Exit Sub
End If


'pluto:7.0
If UCase$(Left$(rdata, 4)) = "/C* " Then
Exit Sub
    'hablar al clan
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "L3")
        Exit Sub
    End If
    
    
    rdata = Right$(rdata, Len(rdata) - 4)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "|*[" & UserList(UserIndex).Name & "]: " & rdata & "´" & FontTypeNames.FONTTYPE_GLOBAL)

    End If
    Exit Sub
End If

'pluto:2.13
   'If UCase$(Left$(rdata, 7)) = "/TRAER " Then
    'rdata = Right$(rdata, Len(rdata) - 7)
    'pluto:2.5.0
    'If UserList(userindex).GuildInfo.GuildName = "" Then Exit Sub

    'tindex = NameIndex(rdata)
    'If tindex <= 0 Then
        'Call SendData(ToIndex, userindex, 0, "||El jugador no esta online." & FONTTYPENAMES.FONTTYPE_INFO)
        'Exit Sub
    'End If
    'pluto:2.4.5
    'If tindex = userindex Then
        'Call SendData(ToIndex, userindex, 0, "||No puedes hacer eso!!" & FONTTYPENAMES.FONTTYPE_INFO)
        'Exit Sub
    'End If
    
 'If UserList(userindex).GuildInfo.GuildName <> UserList(tindex).GuildInfo.GuildName Then
 'Call SendData(ToIndex, userindex, 0, "||No es de tu mismo Clan." & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
'End If
 
 'If Int(UserList(userindex).GuildInfo.GuildPoints / 1000) < Int(UserList(tindex).GuildInfo.GuildPoints / 1000) + 2 And UserList(userindex).GuildInfo.ClanFundado = "" Then
 'Call SendData(ToIndex, userindex, 0, "||No tienes suficiente rango." & FONTTYPENAMES.FONTTYPE_INFO)
 'Exit Sub
 'End If
'If Not ((UserList(userindex).Pos.Map > 165 And UserList(userindex).Pos.Map < 170) Or UserList(userindex).Pos.Map = 185) Then
'Call SendData(ToIndex, userindex, 0, "||Debes estar en un castillo." & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
'End If
'PLUTO:2.4.2
'If UserList(tindex).Counters.Pena > 0 Or UserList(tindex).Pos.Map = 191 Then
'Call SendData(ToIndex, userindex, 0, "||No puede salir de la cárcel." & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
'End If
 'pluto:2.11
 'If UserList(tindex).flags.Paralizado > 0 Or UserList(tindex).flags.Muerto > 0 Then
 'Call SendData(ToIndex, userindex, 0, "||Está muerto o Paralizado." & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
 'End If
    
    'Call SendData(ToIndex, tindex, 0, "||" & UserList(userindex).name & " te ha transportado." & FONTTYPENAMES.FONTTYPE_INFO)
 'pluto:2.9.0
 'Dim aa As Integer
 'If UserList(userindex).Pos.Y > 90 Then aa = -1 Else aa = 1
 
'Call WarpUserChar(tindex, UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y + aa, True)
    
'Exit Sub
'End If
'---------------------fin del /traer -------------------------
  'pluto:2.3
'[Tite] Comando /critico que activa o descactiva el seguro de golpes criticos
If UCase$(Left$(rdata, 8)) = "/CRITICO" Then
    If UserList(UserIndex).flags.SegCritico = True Then
        UserList(UserIndex).flags.SegCritico = False
        Call SendData(ToIndex, UserIndex, 0, "DD1A")
        'Call SendData(ToIndex, UserIndex, 0, "|| Seguro de golpes críticos desactivado." & FONTTYPENAMES.FONTTYPE_INFO)
    Else
        UserList(UserIndex).flags.SegCritico = True
        Call SendData(ToIndex, UserIndex, 0, "DD2A")
        'Call SendData(ToIndex, UserIndex, 0, "|| Seguro de golpes críticos activado." & FONTTYPENAMES.FONTTYPE_INFO)
    End If
Exit Sub
End If

'[/Tite]
'[Tite]Comando /ciudades para ver los dueños de las ciudades
'If UCase$(Left$(rdata, 9)) = "/CIUDADES" Then
   ' Call sendCiudades(UserIndex)
'Exit Sub
'End If
'[\Tite]
'[Tite] Party

'DESCOMENTAR PA VERSION 5.1
'----------------------------
If UCase$(Left$(rdata, 6)) = "/PARTY" Then
    Dim privada As Byte
    If Len(rdata) < 8 Then Exit Sub
    rdata = Right$(rdata, Len(rdata) - 7)
    privada = val(ReadField(1, rdata, 44))
    rdata = Right$(rdata, Len(rdata) - 2)
    Tindex = NameIndex(rdata & "$")
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    tot = val(UserList(NameIndex(rdata & "$")).Stats.ELV)
    If Abs(UserList(UserIndex).Stats.ELV - tot) > 10 Then
        Call SendData(ToIndex, UserIndex, 0, "DD3A")
        Exit Sub
    End If
    'Modificar aqui la diferencia de lvl
    If UserList(UserIndex).Bebe > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "DD4A")
        Exit Sub
    End If
    If NameIndex(rdata & "$") = 0 Or NameIndex(rdata & "$") = UserIndex Then
        Call SendData(ToIndex, UserIndex, 0, "DD5A")
        Exit Sub
    End If
    If UserList(UserIndex).flags.party = True And esLider(UserIndex) = False Then
        Call SendData(ToIndex, UserIndex, 0, "DD6A")
        Exit Sub
    End If
    UserList(UserIndex).flags.privado = privada
    Call InvitaParty(UserIndex, NameIndex(rdata & "$"))
    Exit Sub
End If
        

If UCase$(Left$(rdata, 9)) = "/FINPARTY" Then
    Call quitParty(UserIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/UNIRME" Then
   If UserList(UserIndex).flags.invitado = "" Then
        Call SendData(ToIndex, UserIndex, 0, "DD25")
        Exit Sub
    Else
        Tindex = NameIndex(UserList(UserIndex).flags.invitado & "$")
        If Tindex <= 0 Then
            Call SendData(ToIndex, UserIndex, 0, "DD24")
            Exit Sub
        End If
    End If
  
   'Modificar aqui la diferencia de lvl
    tot = UserList(Tindex).Stats.ELV
    If Abs(UserList(UserIndex).Stats.ELV - tot) > 10 Then
        Call SendData(ToIndex, UserIndex, 0, "DD3A")
        Exit Sub
    End If
    If UserList(UserIndex).Bebe > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "DD4A")
        Exit Sub
    End If
    If esLider(Tindex) = True Then
        Call addUserParty(UserIndex, UserList(Tindex).flags.partyNum)
    Else
        Call creaParty(Tindex, UserList(Tindex).flags.privado)
        Call addUserParty(UserIndex, UserList(Tindex).flags.partyNum)
    End If
    If UserList(UserIndex).flags.party = True Then
        Call SendData(ToIndex, UserIndex, 0, "DD7A" & UserList(partylist(UserList(UserIndex).flags.partyNum).lider).Name)
'        Call SendData(ToIndex, UserIndex, 0, "||Te has unido a la party de " & UserList(partylist(UserList(UserIndex).flags.partyNum).lider).Name & "." & FONTTYPENAMES.FONTTYPE_INFO)
        UserList(UserIndex).flags.invitado = ""
    End If
    Exit Sub
End If
If UCase$(Left$(rdata, 5)) = "/SOLI" Then
    If Len(rdata) < 7 Then Exit Sub
    rdata = Right$(rdata, Len(rdata) - 6)
    Tindex = NameIndex(rdata & "$")
    If Tindex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If esLider(Tindex) = False Then Exit Sub
    If partylist(UserList(NameIndex(rdata & "$")).flags.partyNum).privada = 1 Then
        Exit Sub
    End If
   'Modificar aqui la diferencia de lvl
    tot = UserList(NameIndex(rdata & "$")).Stats.ELV
    If Abs(UserList(UserIndex).Stats.ELV - tot) > 10 Then
        Call SendData(ToIndex, UserIndex, 0, "DD3A")
        Exit Sub
    End If
    If UserList(UserIndex).Bebe > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "DD4A")
        Exit Sub
    End If
    If NameIndex(rdata & "$") = 0 Or NameIndex(rdata & "$") = UserIndex Then
        Call SendData(ToIndex, UserIndex, 0, "DD5A")
        Exit Sub
    End If
    If UserList(UserIndex).flags.party = True And esLider(UserIndex) = False Then
        Call SendData(ToIndex, UserIndex, 0, "DD6A")
        Exit Sub
    End If
    Call addSoliParty(UserIndex, UserList(NameIndex(rdata & "$")).flags.partyNum)
End If

If UCase$(Left$(rdata, 11)) = "/SALIRPARTY" Then
        If UserList(UserIndex).flags.party = False Then
        Call SendData(ToIndex, UserIndex, 0, "DD8A")
'        Call SendData(ToIndex, UserIndex, 0, "||No estas en ninguna party" & FONTTYPENAMES.FONTTYPE_INFO)
        Exit Sub
    End If
    If UserIndex = partylist(UserList(UserIndex).flags.partyNum).lider Then
        Call quitParty(UserIndex)
    Else
        If partylist(UserList(UserIndex).flags.partyNum).numMiembros <= 2 Then
            Call quitParty(partylist(UserList(UserIndex).flags.partyNum).lider)
        Else
            Call quitUserParty(UserIndex)
        End If
    End If
    Exit Sub
End If

'---------------------------------------
'---------------------------------------




'If UCase$(Left$(rdata, 10)) = "/UNIRPARTY" Then
'    If UserList(userindex).flags.party = False Then
'        rdata = Right$(rdata, Len(rdata) - 11)
'        Dim lpp As Byte
'        Dim flpp As Boolean
'        lpp = 1
'        flpp = False
'        For lpp = 1 To MAXPARTYS
'            If partylist(lpp).lider <> 0 The
'                If UCase(UserList(partylist(lpp).lider).Name) = UCase(rdata) Then
'                    flpp = True
'                End If
'            End If
'        Next
'
'        If flpp = True Then
'            Call addUserParty(userindex, partyid(rdata))
'            If UserList(userindex).flags.party = True Then
'                If UserList(userindex).flags.partyNum = partyid(rdata) Then
'                    Call SendData(ToIndex, userindex, 0, "|| Te has incorporado a la party de " & rdata & "." & FONTTYPENAMES.FONTTYPE_INFO)
'                End If
'            End If
'        Else
'            Call SendData(ToIndex, userindex, 0, "|| No hay ninguna party creada por " & rdata & "." & FONTTYPENAMES.FONTTYPE_INFO)
'        End If
'    Else
'        Call SendData(ToIndex, userindex, 0, "||Ya perteneces a una party." & FONTTYPENAMES.FONTTYPE_INFO)
'    End If
'    Exit Sub
'End If
'If UCase$(Left$(rdata, 5)) = "/SEND" Then
'    Call sendMiembrosParty(userindex)
'End If
'If UCase$(Left$(rdata, 12)) = "/RESETPARTYS" Then
'    Call resetPartys
'End If
'[\Tite]

If UCase$(Left$(rdata, 12)) = "/DARMASCOTA " Then
'Exit Sub
'If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub
    'Se asegura que el target es un npc
    If UserList(UserIndex).flags.TargetNpc = 0 Then
     Call SendData(ToIndex, UserIndex, 0, "|| Antes debes seleccionar el NPC Cuidadora de Mascotas." & "´" & FontTypeNames.FONTTYPE_info)

        Exit Sub
    End If
    If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 19 _
    Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
    If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
        Call SendData(ToIndex, UserIndex, 0, "L2")
        Exit Sub
    End If
    If UserList(UserIndex).flags.Montura <> 2 Then
        Call SendData(ToIndex, UserIndex, 0, "|| Debes tener la mascota a tu lado." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    rdata = Right$(rdata, Len(rdata) - 12)
    Call DarMontura(UserIndex, rdata)
    Exit Sub
End If

    If UCase$(Left$(rdata, 8)) = "/VIAJAR " Then
    rdata = Right$(rdata, Len(rdata) - 8)
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        If UserList(UserIndex).flags.TargetNpc = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "L4")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
            End If
                Call SistemaViajes(UserIndex, rdata)
        
                Call SendUserStatsOro(UserIndex)
        
        End If

'Teleportar castillo
If UCase$(Left$(rdata, 10)) = "/CASTILLO " Then
If UserList(UserIndex).Pos.Map = 303 Or UserList(UserIndex).Pos.Map = 66 Then Exit Sub
    If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
        Call SendData(ToIndex, UserIndex, 0, "||Tú salud debe estar completa." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If UserList(UserIndex).Counters.Pena > 0 Or UserList(UserIndex).Pos.Map = 191 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes salir de la cárcel." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
'pluto:6.8 añado mapa dueloclanes
    If UserList(UserIndex).Pos.Map = MapaTorneo2 Or UserList(UserIndex).Pos.Map = 192 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes salir de esta sala." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    If UserList(UserIndex).flags.Paralizado = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes paralizado." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
        'pluto:6.8
        'If UserList(UserIndex).Stats.PClan < 0 Then
        ' Call SendData(ToIndex, UserIndex, 0, "||Puntos Clan en negativo!!" & "´" & FontTypeNames.FONTTYPE_info)
        'Exit Sub
        'End If
        
    rdata = Right$(rdata, Len(rdata) - 10)
    If rdata = "" Then Exit Sub
    If UCase$(rdata) <> "NORTE" And UCase$(rdata) <> "SUR" And UCase$(rdata) <> "ESTE" And UCase$(rdata) <> "OESTE" Then Exit Sub
    X = RandomNumber(48, 55)
    Y = RandomNumber(50, 60)
    Mapa = 0
    If UCase$(rdata) = "NORTE" Then
        If UserList(UserIndex).GuildInfo.GuildName <> castillo1 Then Exit Sub
        Mapa = mapa_castillo1
    End If
    If UCase$(rdata) = "SUR" Then
        If UserList(UserIndex).GuildInfo.GuildName <> castillo2 Then Exit Sub
        Mapa = mapa_castillo2
    End If
    If UCase$(rdata) = "ESTE" Then
        If UserList(UserIndex).GuildInfo.GuildName <> castillo3 Then Exit Sub
        Mapa = mapa_castillo3
    End If
    If UCase$(rdata) = "OESTE" Then
        If UserList(UserIndex).GuildInfo.GuildName <> castillo4 Then Exit Sub
        Mapa = mapa_castillo4
    End If
        'If UCase$(rdata) = "FORTALEZA" Then
           ' If UserList(UserIndex).GuildInfo.GuildName <> fortaleza Then Exit Sub
           ' mapa = 185
        'End If
        
    If Mapa = 0 Then Exit Sub
    Call WarpUserChar(UserIndex, Mapa, X, Y, True)
    Call SendData(ToIndex, UserIndex, 0, "||" & UserList(UserIndex).Name & " transportado." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
End If
    
   'pluto:6.8------------
If UCase$(Left$(rdata, 11)) = "/DUELOCLAN " Then
    'pluto:6.9
    If UserList(UserIndex).Pos.Map = 303 Or UserList(UserIndex).Pos.Map = 66 Then Exit Sub
    If UserList(UserIndex).Counters.Pena > 0 Then Exit Sub
    
    rdata = Right$(rdata, Len(rdata) - 10)
    If rdata = "" Or val(rdata) < 2 Or val(rdata) > 6 Then
    Call SendData(ToIndex, UserIndex, 0, "||Debes indicar el número de participantes (entre 2 y 6) con /DUELOCLAN (espacio) Número." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    
'Case "/DUELOCLAN"
'TClanOcupado = 0
        If UserList(UserIndex).GuildInfo.GuildName = "" Then
        Call SendData(ToIndex, UserIndex, 0, "||No perteneces a ningún clan." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        If UserList(UserIndex).GuildInfo.GuildPoints < 4000 Then
        Call SendData(ToIndex, UserIndex, 0, "||No tienes suficiente rango de clan." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If

        
        
        'TClanOcupado = 0
If TClanOcupado = 0 Then
TClanOcupado = 0
TorneoClan(1).Nombre = ""
TorneoClan(1).numero = 0
TorneoClan(2).Nombre = ""
TorneoClan(2).numero = 0
TClanNumero = val(rdata)
 MsgTorneo = "El Clan " & UserList(UserIndex).GuildInfo.GuildName & " busca rival. Duelo de " & TClanNumero & " Participantes. Si tu clan quiere aceptar el desafío escribe /DUELOCLAN " & TClanNumero
 Call SendData(ToAll, 0, 0, "||" & MsgTorneo & "´" & FontTypeNames.FONTTYPE_pluto)
TClanOcupado = 1
frmMain.Torneo.Enabled = True
frmMain.Torneo.Interval = 20000
TorneoClan(1).Nombre = UserList(UserIndex).GuildInfo.GuildName
Exit Sub
ElseIf TClanOcupado = 1 Then
    If TorneoClan(1).Nombre = UserList(UserIndex).GuildInfo.GuildName Then
    Call SendData(ToIndex, UserIndex, 0, "||Ya estás apuntado." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    If val(rdata) <> TClanNumero Then
    Call SendData(ToIndex, UserIndex, 0, "||El Duelo es de " & TClanNumero & " Participantes. Debes escribir /DUELOCLAN " & TClanNumero & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    
TorneoClan(2).Nombre = UserList(UserIndex).GuildInfo.GuildName
Call SendData(ToAll, 0, 0, "||El Clan " & UserList(UserIndex).GuildInfo.GuildName & " ha aceptado el Desafío." & "´" & FontTypeNames.FONTTYPE_pluto)
MsgTorneo = "Duelo de Clanes: " & TorneoClan(1).Nombre & " vs " & TorneoClan(2).Nombre & " en unos instantes se comunicará el nombre de los participantes."
Call SendData(ToClan, 0, 0, "Duelo de Clanes: " & TorneoClan(1).Nombre & " vs " & TorneoClan(2).Nombre & "´" & FontTypeNames.FONTTYPE_pluto)
TClanOcupado = 2
frmMain.Torneo.Interval = 10000
Else
Call SendData(ToIndex, UserIndex, 0, "||Ya hay un duelo disputandose: " & TorneoClan(1).Nombre & " vs " & TorneoClan(2).Nombre & "´" & FontTypeNames.FONTTYPE_pluto)
End If
 End If
   '------------------------------
   



Select Case UCase$(rdata)
  ' Case "/ADONLINE"
   ' For loopc = 1 To MaxUsers
   ' If (UserList(loopc).Name <> "") Then
   ' tStr = tStr & UserList(loopc).Name & ", "
   ' End If
   ' Next loopc
   ' Call SendData(ToIndex, UserIndex, 0, "H9" & NumUsers & "," & tStr)
   ' Exit Sub
Case "/NODUELOCLAN"
UserList(UserIndex).flags.NoTorneos = True
Call SendData(ToIndex, UserIndex, 0, "||NO estás Disponible para Duelos de Clanes." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub

Case "/SIDUELOCLAN"
If UserList(UserIndex).Pos.Map = 303 Or UserList(UserIndex).Pos.Map = 66 Then Exit Sub
UserList(UserIndex).flags.NoTorneos = False
Call SendData(ToIndex, UserIndex, 0, "||SI estás Disponible para Duelos de Clanes." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub

Case "/DUELOCLAN"
Call SendData(ToIndex, UserIndex, 0, "||Debes indicar el número de participantes (entre 2 y 6) con /DUELOCLAN (espacio) Número." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
    Case "/PING"
If UserList(UserIndex).flags.Privilegios < 2 Then Exit Sub
    Call SendData(ToIndex, UserIndex, 0, "PONG")
    Exit Sub
    


   Case "/ONLINE"
        For loopc = 1 To LastUser
        '    'pluto:2.4.7 --> Quita gms de /online
            If (UserList(loopc).Name <> "") And UserList(loopc).flags.Privilegios < 2 Then
                tStr = tStr & UserList(loopc).Name & ", "
            End If
        Next loopc
        'pluto:2.4.7 --> Quita gms de /online
        If tStr = "" Then Exit Sub
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, UserIndex, 0, "||Número de usuarios: " & Round(NumUsers) & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
            
   'pluto:clan online
    Case "/ONLINECLAN"
 'pluto:2.6.0
        If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub
 'pluto:2.8.0
        If UserList(UserIndex).GuildInfo.GuildName = "" Then Exit Sub
        For loopc = 1 To LastUser
  'pluto:2.4.1 añado rango clan
  'nati: Nuevo metodo para obtener un titulo en el clan, segun puntos.
            Dim a As String
            a = " (Soldado)"
            If UserList(loopc).Stats.PClan >= 100 Then a = " (Teniente)"
            If UserList(loopc).Stats.PClan >= 250 Then a = " (Capitán)"
            If UserList(loopc).Stats.PClan >= 500 Then a = " (General)"
            If UserList(loopc).Stats.PClan >= 1000 Then a = " (Comandante)"
            If UserList(loopc).Stats.PClan >= 1500 Then a = " (SubLider)"
            If UserList(loopc).GuildInfo.GuildPoints >= 5000 Then a = " (Lider)"
            'If UserList(loopc).GuildInfo.GuildPoints >= 1000 Then a = " (Teniente)"
            'If UserList(loopc).GuildInfo.GuildPoints >= 2000 Then a = " (Capitán)"
            'If UserList(loopc).GuildInfo.GuildPoints >= 3000 Then a = " (General)"
            'If UserList(loopc).GuildInfo.GuildPoints >= 4000 Then a = " (SubLider)"
            'If UserList(loopc).GuildInfo.GuildPoints >= 5000 Then a = " (Lider)"
            If UserList(loopc).Name <> "" And UserList(loopc).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName Then
                tStr = tStr & UserList(loopc).Name & " <" & a & ">" & ", "
            End If
        Next loopc
        If tStr = "" Then Exit Sub
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub

    Case "/SALIR"
    'nati: añado que si está transformado no puede salir.
        If UserList(UserIndex).flags.Paralizado > 0 Or UserList(UserIndex).flags.Ceguera > 0 Or UserList(UserIndex).flags.Estupidez > 0 Or UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Or UserList(UserIndex).flags.Morph > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||Este comando esta prohibido en tu estado actual." & "´" & FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        'pluto:6.2
        If MapInfo(UserList(UserIndex).Pos.Map).Terreno = "TORNEO" Then
            Call SendData(ToIndex, UserIndex, 0, "||Este comando esta prohibido en este Mapa." & "´" & FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        Call SendData2(ToIndex, UserIndex, 0, 7)
      
        Call CloseUser(UserIndex)
        Exit Sub
        
    Case "/FUNDARCLAN"
        If UserList(UserIndex).GuildInfo.FundoClan = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya has fundado un clan, solo se puede fundar uno por personaje." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
        If CanCreateGuild(UserIndex) Then
            Call SendData2(ToIndex, UserIndex, 0, 67)
        End If
        Exit Sub
        'pluto:6.0A
        Case "/NIVELCLAN"
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No eres el Lider del Clan!!." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 31 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
        Call SubirLevelClan(UserIndex)
        Exit Sub
        '---------------
'pluto:2.4
    Case "/RECORD"
        Call SendData2(ToIndex, UserIndex, 0, 81, UserCiu & "," & UserCrimi & "," & NNivCiuON & "," & NNivCrimiON & "," & NNivCiu & "," & NNivCrimi & "," & NMoroOn & "," & NMoro & "," & NMaxTorneo & "," & NomClan(1) & "," & NomClan(2)) ' & "," & PuntClan(1) & "," & PuntClan(2))
        Exit Sub
    Case "/TORNEOCLANES"
         For n = 1 To 8
             Call SendData(ToIndex, UserIndex, 0, "||" & n & " - " & NomClan(n) & " ---> " & PuntClan(n) & "´" & FontTypeNames.FONTTYPE_info)
          Next
                Exit Sub

 Case "/PROTECTORES"
   If UserList(UserIndex).flags.Privilegios = 0 Then
          UserList(UserIndex).flags.Privilegios = 3
          'pluto:7.0
Call SendData(ToIndex, UserIndex, 0, "||No puedes convocar más protectores." & FontTypeNames.FONTTYPE_GUILD)
UserList(UserIndex).Stats.PesoMax = 10000
     Else
       UserList(UserIndex).flags.Privilegios = 0
         Exit Sub
    End If
        
'convocar critaturas clan
   ' Case "/NPC1"
    'If UserList(userindex).Stats.GLD < 50000 Then
    'Call SendData(ToIndex, userindex, 0, "||No tienes oro suficiente." & FONTTYPENAMES.FONTTYPE_GUILD)
 'Exit Sub
 'End If
  '  If UserList(userindex).GuildInfo.EsGuildLeader = 0 Then
'Call SendData(ToIndex, userindex, 0, "||Sólo el lider de un clan puede convocar Npcs para defender el Castillo." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
 '      If UserList(userindex).Pos.Map < mapa_castillo1 Or UserList(userindex).Pos.Map > mapa_castillo4 Then
'Call SendData(ToIndex, userindex, 0, "||Debes ir al castillo para convocar Npcs." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
    
   'pluto:2.4
'If NPCHostiles(UserList(userindex).Pos.Map) > 6 Then
'Call SendData(ToIndex, userindex, 0, "||No puedes convocar más protectores." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
    
 '   UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 50000
  '  Call SendUserStatsOro(userindex)
  '  Call SpawnNpc(Criatura_1, UserList(userindex).Pos, True, False)
'Exit Sub
 '     Case "/NPC3"
  '  If UserList(userindex).Stats.GLD < 100000 Then
  '  Call SendData(ToIndex, userindex, 0, "||No tienes oro suficiente." & FONTTYPENAMES.FONTTYPE_GUILD)
 'Exit Sub
 'End If
'    If UserList(userindex).GuildInfo.EsGuildLeader = 0 Then
'Call SendData(ToIndex, userindex, 0, "||Sólo el lider de un clan puede convocar Npcs para defender el Castillo." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
'       If UserList(userindex).Pos.Map < mapa_castillo1 Or UserList(userindex).Pos.Map > mapa_castillo4 Then
'Call SendData(ToIndex, userindex, 0, "||Debes ir al castillo para convocar Npcs." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
 'pluto:2.4
'If NPCHostiles(UserList(userindex).Pos.Map) > 6 Then
'Call SendData(ToIndex, userindex, 0, "||No puedes convocar más protectores." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
    
'    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 100000
 '   Call SendUserStatsOro(userindex)
  '  Call SpawnNpc(Criatura_2, UserList(userindex).Pos, True, False)
'Exit Sub
'   Case "/NPC2"
'    If UserList(userindex).Stats.GLD < 75000 Then
'    Call SendData(ToIndex, userindex, 0, "||No tienes oro suficiente." & FONTTYPENAMES.FONTTYPE_GUILD)
' Exit Sub
' End If
'    If UserList(userindex).GuildInfo.EsGuildLeader = 0 Then
'Call SendData(ToIndex, userindex, 0, "||Sólo el lider de un clan puede convocar Npcs para defender el Castillo." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
'       If UserList(userindex).Pos.Map < mapa_castillo1 Or UserList(userindex).Pos.Map > mapa_castillo4 Then
'Call SendData(ToIndex, userindex, 0, "||Debes ir al castillo para convocar Npcs." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
''pluto:2.4
'If NPCHostiles(UserList(userindex).Pos.Map) > 6 Then
'Call SendData(ToIndex, userindex, 0, "||No puedes convocar más protectores." & FONTTYPENAMES.FONTTYPE_GUILD)
'Exit Sub
'End If
    
'    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 75000
'    Call SendUserStatsOro(userindex)
'    Call SpawnNpc(Criatura_3, UserList(userindex).Pos, True, False)
'Exit Sub
    

Case "/SALIRCLAN"
    If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||Un lider no puede abandonar su clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If
    Dim oGuild As cGuild
    Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
    If oGuild Is Nothing Then Exit Sub
    oGuild.RemoveMember (UserList(UserIndex).Name)
    Set oGuild = Nothing
    UserList(UserIndex).GuildInfo.GuildPoints = 0
    UserList(UserIndex).GuildInfo.GuildName = ""
    'pluto:2.9.0
    UserList(UserIndex).Stats.PClan = 0
    Call SendData(ToIndex, UserIndex, 0, "||Has dejado de pertenecer al clan." & "´" & FontTypeNames.FONTTYPE_GUILD)
    Exit Sub

'Delzak sos offline

'Case "/LEERSOS"

'Dim M As String ---(Duplicada)
'M = Ayuda.Respuesta(UserIndex) & Ayuda.CuantasVecesAparece(UserIndex)
'Call SendData2(ToIndex, UserIndex, 0, 111, M)

'pluto:6.0A liberamascota
Case "/LIBERARMASCOTA"


If UserList(UserIndex).flags.Montura <> 2 Then
Call SendData(ToIndex, UserIndex, 0, "||Debes tener la mascota a tu lado." & "´" & FontTypeNames.FONTTYPE_VENENO)
Exit Sub
End If
Dim xx As Byte
Dim Tipi As Byte
Dim userfile As String
xx = UserList(UserIndex).flags.ClaseMontura
Tipi = UserList(UserIndex).Montura.index(xx)
Call LogMascotas("Liberar: " & UserList(UserIndex).Name & " mascota tipo " & xx & " del INDEX " & Tipi)

'ponemos todo a cero
Call ResetMontura(UserIndex, xx)
'grabamos ficha todo a cero
userfile = CharPath & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".chr"
Call WriteVar(userfile, "MONTURA" & Tipi, "NIVEL", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "EXP", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "ELU", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "VIDA", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "GOLPE", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "NOMBRE", "")
Call WriteVar(userfile, "MONTURA" & Tipi, "ATCUERPO", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "DEFCUERPO", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "ATFLECHAS", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "DEFFLECHAS", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "ATMAGICO", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "DEFMAGICO", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "EVASION", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "LIBRES", 0)
Call WriteVar(userfile, "MONTURA" & Tipi, "TIPO", 0)

Call QuitarObjetos(UserList(UserIndex).flags.ClaseMontura + 887, 1, UserIndex)
Call LogMascotas("Liberar: " & UserList(UserIndex).Name & " quitamos objeto " & UserList(UserIndex).flags.ClaseMontura + 887)

Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
            Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
            Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
            UserList(UserIndex).MascotasIndex(i) = 0
            UserList(UserIndex).MascotasType(i) = 0
        End If
    End If
Next i
UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
'If UserList(UserIndex).Nmonturas > 0 Then
UserList(UserIndex).Nmonturas = UserList(UserIndex).Nmonturas - 1
Call LogMascotas("Liberar: " & UserList(UserIndex).Name & " ahora tiene " & UserList(UserIndex).Nmonturas)

UserList(UserIndex).flags.Montura = 0
UserList(UserIndex).flags.ClaseMontura = 0
Call WriteVar(userfile, "MONTURAS", "NroMonturas", val(UserList(UserIndex).Nmonturas))
Exit Sub
'---------fin pluto:2.4--------------------
  
    Case "/BALANCE"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "L4")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
        Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If Not PersonajeExiste(UserList(UserIndex).Name) Then
              Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
              CloseUser (UserIndex)
              Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "||6°Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        Exit Sub
    Case "/QUIETO" ' << Comando a mascotas
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "L3")
                      Exit Sub
         End If
         'Se asegura que el target es un npc
         If UserList(UserIndex).flags.TargetNpc = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "L4")
                  Exit Sub
         End If
         If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "L2")
                      Exit Sub
         End If
         If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> _
            UserIndex Then Exit Sub
         Npclist(UserList(UserIndex).flags.TargetNpc).Movement = ESTATICO
         Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
         Exit Sub
    Case "/ACOMPAÑAR"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "L4")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> _
          UserIndex Then Exit Sub
        Call FollowAmo(UserList(UserIndex).flags.TargetNpc)
        Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
        Exit Sub
   ' Case "/ENTRENAR"
        '¿Esta el user muerto? Si es asi no puede comerciar
        'If UserList(UserIndex).flags.Muerto = 1 Then
               '  Call SendData(ToIndex, UserIndex, 0, "L3")
                '  Exit Sub
       ' End If
        'Se asegura que el target es un npc
        'If UserList(UserIndex).flags.TargetNpc = 0 Then
             ' Call SendData(ToIndex, UserIndex, 0, "L4")
             ' Exit Sub
       ' End If
        'If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                '  Call SendData(ToIndex, UserIndex, 0, "L2")
                '  Exit Sub
        'End If
       ' If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
       ' Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNpc)
       ' Exit Sub
    Case "/DESCANSAR"
    
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        'pluto.7.0
            If UserList(UserIndex).flags.Macreanda > 0 Then Exit Sub

    'Delzak (28-8-10)
        If UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes descansar estando transformado." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If

        'If UserList(UserIndex).flags.Paralizado > 0 Then exit sub
        
        
        If HayOBJarea(UserList(UserIndex).Pos, FOGATA) Then
                Call SendData2(ToIndex, UserIndex, 0, 41)
                If Not UserList(UserIndex).flags.Descansar Then
                    Call SendData(ToIndex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & "´" & FontTypeNames.FONTTYPE_info)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Te levantas." & "´" & FontTypeNames.FONTTYPE_info)
                End If
                UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
        Else
                If UserList(UserIndex).flags.Descansar Then
                    Call SendData(ToIndex, UserIndex, 0, "||Te levantas." & "´" & FontTypeNames.FONTTYPE_info)
                    
                    UserList(UserIndex).flags.Descansar = False
                    Call SendData2(ToIndex, UserIndex, 0, 41)
                    Exit Sub
                End If
                Call SendData(ToIndex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & "´" & FontTypeNames.FONTTYPE_info)
        End If
        Exit Sub
    Case "/MEDITAR"
        'pluto:2.15
        If UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then Exit Sub
            'pluto.7.0
            If UserList(UserIndex).flags.Macreanda > 0 Then Exit Sub
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        Call SendData2(ToIndex, UserIndex, 0, 54)
        If Not UserList(UserIndex).flags.Meditando Then
           Call SendData(ToIndex, UserIndex, 0, "||Comenzas a meditar." & "´" & FontTypeNames.FONTTYPE_info)
        Else
           Call SendData(ToIndex, UserIndex, 0, "G7")
        'pluto:2.5.0
        Call SendData2(ToIndex, UserIndex, 0, 15, UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)

        End If
        UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
    If UserList(UserIndex).flags.Meditando Then
            UserList(UserIndex).Char.loops = LoopAdEternum
    
    'Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 131 & "," & LoopAdEternum)
'UserList(UserIndex).Char.FX = 131
 '  Exit Sub
            'pluto:6.5
            If UserList(UserIndex).flags.DragCredito5 = 1 Then
            Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 17 & "," & LoopAdEternum)
            UserList(UserIndex).Char.FX = 17
              Exit Sub
            End If
          
            '----------------------
            
            'pluto:2.14 meditar para remorts
        If UserList(UserIndex).Remort > 0 Then
            
            If Not Criminal(UserIndex) Then
              Select Case UserList(UserIndex).Stats.ELV
              Case Is < 10
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 98 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 98
              Case 10 To 19
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 127 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 127
              Case 20 To 29
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 125 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 125
               Case 30 To 39
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 117 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 132
                  Case 40 To 49
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 97 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 97
                'pluto:6.9
                Case 50 To 59
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 112 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 112
                'pluto:6.9
                Case Is > 59
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 112 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 112 '130
              End Select
            
            
        Else
            
               Select Case UserList(UserIndex).Stats.ELV
              Case Is < 10
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 99 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 99
              Case 10 To 19
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 126 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 126
             Case 20 To 29
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 124 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 124
            Case 30 To 39
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 118 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 118
            Case 40 To 49
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 96 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 96
                'pluto:6.9
                Case 50 To 59
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 111 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 111
                'pluto:6.9
                Case Is > 59
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 111 & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = 111 '131
              End Select
           
              End If
            
            
            Exit Sub
        End If 'REMORT
            '----------------MEDITACION PARA NO REMORTS------
            
            If UserList(UserIndex).Stats.ELV < 10 Then
                Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARCHICO
            
            ElseIf UserList(UserIndex).Stats.ELV < 20 Then
                Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
            
            ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARGRANDE
            ElseIf UserList(UserIndex).Stats.ELV < 40 Then
                Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & FXMEDITARRAYOS & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARRAYOS
            ElseIf UserList(UserIndex).Stats.ELV < 50 Then
                
                If Not Criminal(UserIndex) Then
                Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 70 & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = 70
                Else
                Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 69 & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = 69
                End If
            ElseIf UserList(UserIndex).Stats.ELV > 60 Then
                If Not Criminal(UserIndex) Then
                Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 131 & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = 131
                Else
                Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 130 & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = 130
                End If
            ElseIf Not Criminal(UserIndex) Then
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & FXMEDITARorbitalazul & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = FXMEDITARorbitalazul
            Else
              Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & FXMEDITARorbitalrojo & "," & LoopAdEternum)
              UserList(UserIndex).Char.FX = FXMEDITARorbitalrojo
            End If
    Else 'DEJAR DE MEDITAR
            UserList(UserIndex).Char.FX = 0
            UserList(UserIndex).Char.loops = 0
            'pluto:2-3-04 bug fx meditar
             Call SendData2(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
    End If
        Exit Sub
  'pluto:2.4.2
Case "/FORTALEZA"
If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
Call SendData(ToIndex, UserIndex, 0, "||Tú salud debe estar completa." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
If UserList(UserIndex).Counters.Pena > 0 Or UserList(UserIndex).Pos.Map = 191 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes salir de la cárcel." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

 'pluto:2.12
        If MapInfo(UserList(UserIndex).Pos.Map).Terreno = "TORNEO" Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes salir de esta sala." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Paralizado = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes paralizado." & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If


'pluto:2.4.5
If UCase$(UserList(UserIndex).GuildInfo.GuildName) <> UCase$(fortaleza) Then Exit Sub
          
X = RandomNumber(60, 70)
Y = RandomNumber(29, 35)
Call WarpUserChar(UserIndex, 185, X, Y, True)
Call SendData(ToIndex, UserIndex, 0, "||" & UserList(UserIndex).Name & " transportado." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub




Case "/RESUCITAR"
       'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 1 _
       Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
       If UserList(UserIndex).flags.Navegando > 0 Then
       Call SendData(ToIndex, UserIndex, 0, "Deja de Navegar!!.")
       Exit Sub
       End If
       'If Not PersonajeExiste(UserList(UserIndex).Name) Then
           'Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
           'CloseUser (UserIndex)
           'Exit Sub
       'End If
       Call RevivirUsuario(UserIndex)
       Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido resucitado!!" & "´" & FontTypeNames.FONTTYPE_info)
       'pluto:2.14
       Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 72 & "," & 1)

       UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
       Call SendUserStatsVida(val(UserIndex))
       Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido curado!!" & "´" & FontTypeNames.FONTTYPE_info)
       Exit Sub



Case "/AYUDA"
       Call SendHelp(UserIndex)
           Exit Sub

Case "/ANGEL"
'pluto:6.4
If UserList(UserIndex).Pos.Map = MapaAngel Or (UserList(UserIndex).Pos.Map > 165 And UserList(UserIndex).Pos.Map < 170) Or UserList(UserIndex).Pos.Map = 185 Then
Call SendData(ToIndex, UserIndex, 0, "||¡¡No te puedes transformar en este Mapa!!" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'pluto:2.12
If MapInfo(UserList(UserIndex).Pos.Map).Terreno = "TORNEO" Then Exit Sub
'pluto:2.4
If Criminal(UserIndex) Or UserList(UserIndex).Stats.ELV < 50 Or UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Invisible > 0 Or UserList(UserIndex).flags.Muerto > 0 Or UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Oculto > 0 Then Exit Sub
If UserList(UserIndex).flags.Montura > 0 Then Exit Sub
If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub

'pluto:6.9
If UserList(UserIndex).flags.Invisible > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes estando invisible!!" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

'pluto:
If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
Call SendData(ToIndex, UserIndex, 0, "||¡¡No tienes suficiente energía!!" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'UserList(UserIndex).Counters.Morph = IntervaloMorphPJ
UserList(UserIndex).flags.Angel = UserList(UserIndex).Char.Body
'[gau]
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, val(234), val(0), UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 1 & "," & 0)
Exit Sub
         
Case "/DEMONIO"
'pluto:2.15
If UserList(UserIndex).Pos.Map = MapaAngel Or (UserList(UserIndex).Pos.Map > 165 And UserList(UserIndex).Pos.Map < 170) Or UserList(UserIndex).Pos.Map = 185 Then
Call SendData(ToIndex, UserIndex, 0, "||¡¡No te puedes transformar en este Mapa!!" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'pluto:6.2
If MapInfo(UserList(UserIndex).Pos.Map).Terreno = "TORNEO" Then Exit Sub

If Not Criminal(UserIndex) Or UserList(UserIndex).Stats.ELV < 50 Or UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Demonio > 0 Or UserList(UserIndex).flags.Invisible > 0 Or UserList(UserIndex).flags.Muerto > 0 Or UserList(UserIndex).flags.Oculto > 0 Then Exit Sub
If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
'pluto:2.4
If UserList(UserIndex).flags.Montura > 0 Then Exit Sub
'pluto:6.9
If UserList(UserIndex).flags.Invisible > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes estando invisible!!" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'pluto:
If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
Call SendData(ToIndex, UserIndex, 0, "||¡¡No tienes suficiente energía!!" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'UserList(UserIndex).Counters.Morph = IntervaloMorphPJ
UserList(UserIndex).flags.Demonio = UserList(UserIndex).Char.Body
'[gau]
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, val(239), val(0), UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 1 & "," & 0)
Exit Sub
     Case "/EST"
        Call SendUserStatstxt(UserIndex, UserIndex)
        Exit Sub
'pluto:2-3-04
'pluto:2.4
Case "/DRAGPUNTOS"
Call SendData(ToIndex, UserIndex, 0, "||Dragpuntos: " & UserList(UserIndex).Stats.Puntos & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Puntos Torneos: " & UserList(UserIndex).Stats.GTorneo & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Puntos Aportados al Clan: " & UserList(UserIndex).Stats.PClan & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Guildpoints: " & UserList(UserIndex).GuildInfo.GuildPoints & "´" & FontTypeNames.FONTTYPE_info)
'pluto:2.20
Call SendData(ToIndex, UserIndex, 0, "||Quest Completadas: " & UserList(UserIndex).Mision.numero & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Clanes Participado: " & UserList(UserIndex).GuildInfo.ClanesParticipo & "´" & FontTypeNames.FONTTYPE_info)
Call SendData(ToIndex, UserIndex, 0, "||Solicitudes Restantes: " & (10 + Int(UserList(UserIndex).Mision.numero / 20) - UserList(UserIndex).GuildInfo.ClanesParticipo) & "´" & FontTypeNames.FONTTYPE_info)
'------------
Exit Sub
 
'pluto:2.14
Case "/BODA"

    If MapData(188, 49, 47).UserIndex > 0 And MapData(188, 50, 47).UserIndex > 0 Then
Dim boda1 As Integer
Dim boda2 As Integer
boda1 = MapData(188, 49, 47).UserIndex
boda2 = MapData(188, 50, 47).UserIndex

If ((UserList(boda1).Madre = UserList(boda2).Madre) And UserList(boda1).Madre <> "") Or (UserList(boda1).Genero = UserList(boda2).Genero) Or UserList(boda1).Esposa > "" Or UserList(boda2).Esposa > "" Or UserList(boda1).Bebe > 0 Or UserList(boda2).Bebe > 0 Then Exit Sub
'pluto:6.0A comprueba anillos y los quita
If Not TieneObjetos(990, 1, boda1) Or Not TieneObjetos(990, 1, boda2) Then
Call SendData(ToIndex, UserIndex, 0, "||Os faltan los Anillos de Boda." & "´" & FontTypeNames.FONTTYPE_talk)
Exit Sub
End If
'pluto:6.2---------------
If UserList(boda1).Invent.AnilloEqpObjIndex > 0 Or UserList(boda2).Invent.AnilloEqpObjIndex > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||Los Anillos deben estar desequipados." & "´" & FontTypeNames.FONTTYPE_talk)
Exit Sub
End If
'-------------------------
Call QuitarObjetos(990, 1, boda1)
Call QuitarObjetos(990, 1, boda2)

'---------------
UserList(boda1).Esposa = UserList(boda2).Name
UserList(boda2).Esposa = UserList(boda1).Name
Call SendData(ToAll, 0, 0, "||Felicidades a " & UserList(boda1).Name & " y " & UserList(boda2).Name & " que acaban de celebrar su Boda." & "´" & FontTypeNames.FONTTYPE_talk)
Call SendData2(ToPCArea, boda1, UserList(boda1).Pos.Map, 22, UserList(boda1).Char.CharIndex & "," & 88 & "," & 35)
Call SendData2(ToPCArea, boda2, UserList(boda2).Pos.Map, 22, UserList(boda2).Char.CharIndex & "," & 88 & "," & 35)
Call SendData(ToMap, boda1, UserList(boda1).Pos.Map, "TM" & 25)

    'pluto:6.0A
    Else
  Call SendData(ToIndex, UserIndex, 0, "||Situaros los dos justo delante del Altar." & "´" & FontTypeNames.FONTTYPE_talk)
  End If
Exit Sub
 'pluto:2.17
 Case "/DIVORCIO"

If UserList(UserIndex).Esposa = "" Then Exit Sub
'Dim Tindex As Integer
Tindex = NameIndex(UserList(UserIndex).Esposa & "$")
    'esta online
    If Tindex > 0 Then
UserList(Tindex).Esposa = ""
UserList(Tindex).Amor = 0
Call SendData(ToIndex, Tindex, 0, "||Tu Pareja se ha divorciado." & "´" & FontTypeNames.FONTTYPE_talk)
   Else ' no esta online
'Dim userfile As String
userfile = CharPath & Left$(UserList(UserIndex).Esposa, 1) & "\" & UCase$(UserList(UserIndex).Esposa) & ".chr"
Call WriteVar(userfile, "INIT", "Esposa", "")
Call WriteVar(userfile, "INIT", "Amor", 0)
    End If

UserList(UserIndex).Esposa = ""
UserList(UserIndex).Amor = 0
Call SendData(ToIndex, UserIndex, 0, "||Te has Divorciado de tu Pareja." & "´" & FontTypeNames.FONTTYPE_talk)
Exit Sub

'NICO
 Case "/REGRESAR"
If UserList(UserIndex).Pos.Map = 66 Then 'REMPLAZEN EL 66 POR EL NUM DE MAP DE LA CARCEL DE SU SV
            Call SendData(ToIndex, UserIndex, 0, "||No escaparás de la carcel." & FONTTYPE_talk) ' Juance!
            Exit Sub
        End If
 
        If UserList(UserIndex).Pos.Map = 303 Then 'REMPLAZEN EL MAPA POR EL QUE QUIERAN
            Call SendData(ToIndex, UserIndex, 0, "||No podés abandonar este mapa, si deseas regresar, pidele a un GM via /GM" & "´" & FontTypeNames.FONTTYPE_talk)
            Exit Sub
        End If
        If UserList(UserIndex).Pos.Map = 49 Then 'REMPLAZEN EL MAPA POR EL QUE QUIERAN
            Call SendData(ToIndex, UserIndex, 0, "||No podés abandonar este mapa, si deseas regresar, pidele a un GM via /GM " & "´" & FontTypeNames.FONTTYPE_talk)
            Exit Sub
        End If
       If UserList(UserIndex).flags.Muerto = 0 Then 'SI EL USER ESTA VIVO, NO PUEDE REGRESAR.
             Call SendData(ToIndex, UserIndex, 0, "||No podés volver a la ciudad si estas vivo!" & "´" & FontTypeNames.FONTTYPE_talk)
            Exit Sub
        End If
       If UserList(UserIndex).Pos.Map = 34 Then 'Si el user esta en ulla, no lo deja regresar, podria laguear el sv con eso
            Call SendData(ToIndex, UserIndex, 0, "||Ya estas en Nix!" & "´" & FontTypeNames.FONTTYPE_talk)
           Exit Sub
        End If
        
        If UserList(UserIndex).Stats.GLD < 20000 Then
                Call SendData(ToIndex, UserIndex, 0, "||Necesitas 20000 de oro." & "´" & FontTypeNames.FONTTYPE_talk)
                Exit Sub
            End If
                   
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 20000
            Call SendUserStatsOro(UserIndex)
                   
        Call WarpUserChar(UserIndex, 34, 50, 50, True)
       
    Exit Sub
 
 
 'pluto:7.0
 Case "/CIUDAD"
 If UserList(UserIndex).raza <> "Vampiro" Then Exit Sub

        If UserList(UserIndex).Counters.Pena > 0 Or UserList(UserIndex).Pos.Map = 191 Then Exit Sub
 
        If UserList(UserIndex).flags.Paralizado > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No puedes estando paralizado!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        
        If UserList(UserIndex).Char.Body <> 9 And UserList(UserIndex).Char.Body <> 260 Then
        Call SendData(ToIndex, UserIndex, 0, "||Debes estar Transformado para la Teleportación!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
        
        Dim C As Byte
        C = RandomNumber(1, 5)
        
        If C = 1 Then
        va1 = Nix.Map
        va2 = Nix.X + C
        va3 = Nix.Y
        End If
        
        If C = 2 Then
        va1 = Banderbill.Map
        va2 = Banderbill.X
        va3 = Banderbill.Y - C
        End If
        
        If C = 3 Then
        va1 = Ullathorpe.Map
        va2 = Ullathorpe.X + C
        va3 = Ullathorpe.Y
        End If
        'If C = 4 Then
        'va1 = Lindos.Map
        'va2 = Lindos.X
        'va3 = Lindos.Y
        'End If
        
        If C = 4 Then
        va1 = 170
        va2 = 34
        va3 = 34 + C
        End If
        
Call WarpUserChar(UserIndex, va1, va2, va3, True)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 100 & "," & 1)
'Sonido
SendData ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_tele
'solo una vez por transformación.
UserList(UserIndex).Counters.Morph = 0
UserList(UserIndex).Stats.MinSta = 0
Exit Sub
 
 'pluto:2.8.0
 Case "/VAMPIRO"
'pluto:2.11
   Dim abody As Integer
If UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Muerto > 0 Or UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Then Exit Sub

    If UCase$(UserList(UserIndex).raza) = "VAMPIRO" Then
    UserList(UserIndex).Counters.Morph = IntervaloMorphPJ
    UserList(UserIndex).flags.Morph = UserList(UserIndex).Char.Body
 
    If UserList(UserIndex).Stats.ELV < 40 Then abody = 9 Else abody = 260
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, val(abody), val(0), UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
    Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(42).FXgrh & "," & Hechizos(25).loops)
    Exit Sub
    End If
    
    'pluto:7.0 berserker
    If UCase$(UserList(UserIndex).raza) = "ORCO" Then
    If UserList(UserIndex).flags.Montura > 0 Then Exit Sub
    If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
    If UserList(UserIndex).flags.Invisible > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes estando invisible!!" & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
    Call SendData(ToIndex, UserIndex, 0, "||¡¡No tienes suficiente energía!!" & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    UserList(UserIndex).Counters.Morph = IntervaloMorphPJ
    UserList(UserIndex).flags.Morph = UserList(UserIndex).Char.Body
    'Dim abody As Integer
    If UserList(UserIndex).Genero = "Hombre" Then abody = 214 Else abody = 214
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, val(abody), val(0), UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
    Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(42).FXgrh & "," & Hechizos(25).loops)
    Call SendData(ToIndex, UserIndex, 0, "||Te has transformado en Berserker !!" & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If
    
Exit Sub

'pluto:6.0A
 Case "/MINOTAURO"
If UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Muerto > 0 Or UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Then Exit Sub
If UserList(UserIndex).flags.Minotauro = 0 Then Exit Sub

UserList(UserIndex).Counters.Morph = IntervaloMorphPJ
UserList(UserIndex).flags.Morph = UserList(UserIndex).Char.Body
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, 380, val(0), UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(43).FXgrh & "," & Hechizos(25).loops)
Exit Sub
    
    'pluto:6.9
 Case "/HIPOPOTAMO"
 If UserList(UserIndex).flags.DragCredito6 <> 1 And UserList(UserIndex).flags.DragCredito6 <> 4 Then Exit Sub
If UserList(UserIndex).flags.Montura <> 1 Then Exit Sub
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, 365, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(43).FXgrh & "," & Hechizos(25).loops)
Exit Sub
        'pluto:6.9
 Case "/PANTERA"
 If UserList(UserIndex).flags.DragCredito6 <> 2 And UserList(UserIndex).flags.DragCredito6 <> 4 Then Exit Sub
If UserList(UserIndex).flags.Montura <> 1 Then Exit Sub
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, 350, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(43).FXgrh & "," & Hechizos(25).loops)
Exit Sub

        'pluto:6.9
 Case "/CIERVO"
 If UserList(UserIndex).flags.DragCredito6 <> 3 And UserList(UserIndex).flags.DragCredito6 <> 4 Then Exit Sub
If UserList(UserIndex).flags.Montura <> 1 Then Exit Sub
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, 344, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & Hechizos(43).FXgrh & "," & Hechizos(25).loops)
Exit Sub
    
    Case "/MUERTES"
        Call SendUserMuertes(UserIndex, UserIndex)
        Exit Sub
    'pluto:2.3
    Case "/MONTURA"
    'Call EnviarMontura(UserIndex)
    Exit Sub
 

    Case "/COMERCIAR"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
        '¿El target es un NPC valido?
        'If UserList(userindex).flags.TargetNpc > 0 Then
              '¿El NPC puede comerciar?
              'If Npclist(UserList(userindex).flags.TargetNpc).Comercia = 0 Then
                 'If Len(Npclist(UserList(userindex).flags.TargetNpc).Desc) > 0 Then Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "||6°No tengo ningun interes en comerciar.°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
                 'Exit Sub
              'End If
              'If Distancia(Npclist(UserList(userindex).flags.TargetNpc).Pos, UserList(userindex).Pos) > 3 Then
                  'Call SendData(ToIndex, userindex, 0, "L2")
                 ' Exit Sub
             ' End If
              'Iniciamos la rutina pa' comerciar.
             'Call IniciarCOmercioNPC(userindex)
         'pluto:2.6.0
         '[Alejo]
        'Else
        If UserList(UserIndex).flags.TargetUser > 0 Then
       'pluto:6.9
       If UserList(UserIndex).Pos.Map = 171 Or UserList(UserIndex).Pos.Map = 177 Or MapInfo(UserList(UserIndex).Pos.Map).Terreno = "TORNEO" Then Exit Sub
       
   
            'Comercio con otro usuario
            'Puede comerciar ?
            If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes comerciar con los muertos!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
                Exit Sub
            End If

'soy yo ?
            If UserList(UserIndex).flags.TargetUser = UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes comerciar contigo mismo..." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
                Exit Sub
            End If
            'pluto:2.9.0
            If UserList(UserIndex).flags.Privilegios > 0 Or UserList(UserList(UserIndex).flags.TargetUser).flags.Privilegios > 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes comerciar con el GM" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
                Exit Sub
            End If
            
            'ta muy lejos ?
            If Distancia(UserList(UserList(UserIndex).flags.TargetUser).Pos, UserList(UserIndex).Pos) > 3 Then
                Call SendData(ToIndex, UserIndex, 0, "G9")
                Exit Sub
            End If
            'Ya ta comerciando ? es con migo o con otro ?
            If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
                Exit Sub
            End If

'pluto:2.7.0
'maximo inventario
Dim ii As Byte
'pluto:2.9.0
Dim i1 As Byte
Dim i2 As Byte
i1 = 0
i2 = 0
For ii = 1 To MAX_INVENTORY_SLOTS
If UserList(UserIndex).Invent.Object(ii).ObjIndex = 0 Then i1 = i1 + 1
If i1 > 3 Then GoTo u1
Next ii
Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes comerciar tienes el inventario muy lleno!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Exit Sub
u1:
For ii = 1 To MAX_INVENTORY_SLOTS
If UserList(UserList(UserIndex).flags.TargetUser).Invent.Object(ii).ObjIndex = 0 Then i2 = i2 + 1
Next ii
If i2 > 3 Then GoTo u2
Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes comerciar porque el tiene su inventario muy lleno!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Exit Sub
u2:
 
If UserList(UserIndex).flags.Montura > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||¡¡No uses la mascota mientras comercias!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Exit Sub
End If
 If UserList(UserIndex).flags.Navegando > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||¡¡No comercies mientras navegas!!" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Exit Sub
End If
 '---------------------------------------
            'inicializa unas variables...
            UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
            UserList(UserIndex).ComUsu.Cant = 0
            UserList(UserIndex).ComUsu.Objeto = 0
            UserList(UserIndex).ComUsu.Acepto = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
        Else
            Call SendData(ToIndex, UserIndex, 0, "L4")
        End If
        Exit Sub
    '[/Alejo]
    '[KEVIN]------------------------------------------

   
    
    'pluto:hoy
        Case "/QUEST"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
        '¿El target es un NPC valido?
        If UserList(UserIndex).flags.TargetNpc > 0 Then
              If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
              End If
             'pluto:2-3-04
        If UserList(UserIndex).Mision.numero = 203 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Todas las misiones Completadas!!" & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
        End If
             
              
              If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 14 Then
                'pluto:7.0
                'UserList(UserIndex).Mision.estado = 0
                If UserList(UserIndex).Mision.estado > 0 Then
                Call ContinuarQuest(UserIndex)
                Else
                Call iniciarquest(UserIndex)
                End If
    
              Else
                Exit Sub
              End If
        Else
          Call SendData(ToIndex, UserIndex, 0, "L4")
        End If
        Exit Sub
     
     '[/KEVIN]------------------------------------
 

     Case "/HOGAR"
        If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Or UserList(UserIndex).Pos.Map = 66 Then Exit Sub
            Call SendData(ToIndex, UserIndex, 0, "||No puedes regresar a tu ciudad desde esta ubicación." & FONTTYPE_info)
            Exit Sub
        If UserList(UserIndex).flags.Muerto = 0 Then Call UserDie(UserIndex)
        WarpUserChar UserIndex, 34, 60, 67, True
        Call SendData(ToIndex, UserIndex, 0, "||Has regresado a Nix." & FONTTYPE_info)
        Exit Sub
 
   


    Case "/ENLISTAR"
        'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then
              Call EnlistarArmadaReal(UserIndex)
       End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 1 Then
       Call EnlistarCaos(UserIndex)
       End If
       'enlistar legion
       If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 2 Then
       'pluto:2.15 Fuera legión
       'Call Enlistarlegion(UserIndex)
       End If
       
       Exit Sub
    Case "/INFORMACION"
       'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then
            If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||6°No perteneces a las tropas reales!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "||6°Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa.°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
       Else
            If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||6°No perteneces a las fuerzas del caos!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "||6°Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa.°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
       End If
       Exit Sub
    
       'pluto:2.24
 Case "/GRIAL"
      'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 28 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
   
    If Not TieneObjetos(157, 3, UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6°No tienes las 3 Copas Griales!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
     Exit Sub
    End If
    
    Call QuitarObjetos(157, 3, UserIndex)
    Call CambiarGriaL(UserIndex)
    Exit Sub
    
 Case "/CABALLERO"
      'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 120 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
   
    If Not TieneObjetos(1241, 5, UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6°No tienes las 5 Bolas de Cristal!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
     Exit Sub
    End If
    
    Call QuitarObjetos(1241, 5, UserIndex)
    Call CambiarBola(UserIndex)
    Exit Sub
    
    
    Case "/TROFEO"
      'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 130 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
   
    If Not TieneObjetos(1245, 3, UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6°No tienes las 3 Trofeos de Primer Puesto!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
     Exit Sub
    End If
    
    Call QuitarObjetos(1245, 3, UserIndex)
    Call CambiarTrofeo(UserIndex)
    Exit Sub
    
    Case "/TROFEO2"
      'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 140 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
   
    If Not TieneObjetos(1246, 3, UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6°No tienes las 3 Trofeos de Segundo Puesto!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
     Exit Sub
    End If
    
    Call QuitarObjetos(1246, 3, UserIndex)
    Call CambiarTrofeo(UserIndex)
    Exit Sub
    'pluto:2.3
    
     Case "/DRAGON"
      'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 18 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
     Dim ge As Integer
     For ge = 406 To 413
    If Not TieneObjetos(ge, 1, UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6°No tienes todas las Gemas!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
    End If
    Next ge
    If Not TieneObjetos(598, 1, UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||6°No tienes todas las Gemas!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
     Exit Sub
    End If
    For ge = 406 To 413
    Call QuitarObjetos(ge, 1, UserIndex)
    Next ge
    Call QuitarObjetos(598, 1, UserIndex)
    Call CambiarGemas(UserIndex)
    Exit Sub
    
    Case "/RECOMPENSA"
       'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "L4")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "L2")
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then
            If UserList(UserIndex).Faccion.ArmadaReal <> 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||6°No perteneces a las tropas reales!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call RecompensaArmadaReal(UserIndex)
        End If
        
        If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 1 Then
            If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||6°No perteneces a las fuerzas del caos!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call RecompensaCaos(UserIndex)
       End If
       'recompensa legion
       If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 2 Then
            If UserList(UserIndex).Faccion.ArmadaReal <> 2 Then
                Call SendData(ToIndex, UserIndex, 0, "||6°No perteneces a las tropas de la Legión!!!°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call Recompensalegion(UserIndex)
       End If
       
       Exit Sub


        
         Case "/ROSTRO"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
       
        
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "L4")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
        End If
        
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_CIRUJANO _
        Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "L2")
              Exit Sub
        End If
        
       Dim u As Integer
       If UserList(UserIndex).Genero = "Hombre" Then
        Select Case (UserList(UserIndex).raza)
        
                Case "Humano"
                    u = CInt(RandomNumber(1, 49))
                    If u = 27 Then u = 28
                Case "Ciclope"
                u = CInt(RandomNumber(1, 2)) + 801
                'If u = 27 Then u = 28
                
                Case "Elfo"
                    u = CInt(RandomNumber(1, 19)) + 100
                    If u > 119 Then u = 119
                    
                Case "Elfo Oscuro"
                    u = CInt(RandomNumber(1, 16)) + 200
                    If u > 216 Then u = 216
                    
                Case "Enano"
                    u = RandomNumber(1, 11) + 300
                    If u > 311 Then u = 311
                'pluto:7.0
                 Case "Goblin"
                    u = RandomNumber(1, 8) + 704
                    If u > 712 Then u = 712
                
                Case "Gnomo"
                    u = RandomNumber(1, 10) + 400
                    If u > 410 Then u = 410
                  Case "Orco"
                  u = CInt(RandomNumber(1, 6)) + 600
                    If u > 606 Then u = 606
                Case "Vampiro"
                u = CInt(RandomNumber(1, 8)) + 504
                    If u > 512 Then u = 512
                Case Else
                    u = 1
        End Select
        End If
        'mujer
        If UserList(UserIndex).Genero = "Mujer" Then
        Select Case (UserList(UserIndex).raza)
                Case "Humano"
                    u = CInt(RandomNumber(1, 13)) + 69
                    If u > 82 Then u = 82
                Case "Ciclope"
                u = 801
                    'If u > 82 Then u = 82
                Case "Elfo"
                    u = CInt(RandomNumber(1, 11)) + 169
                    If u > 180 Then u = 180
                    
                Case "Elfo Oscuro"
                    u = CInt(RandomNumber(1, 8)) + 269
                    If u > 277 Then u = 277
                    
                Case "Goblin"
                    u = RandomNumber(1, 4) + 700
                    If u > 704 Then u = 704
                Case "Gnomo"
                    u = RandomNumber(1, 6) + 469
                    If u > 475 Then u = 475
                Case "Enano"
                    u = RandomNumber(1, 3) + 369
                    If u > 472 Then u = 472
                Case "Orco"
                u = RandomNumber(1, 3) + 606
                    If u > 609 Then u = 609
                Case "Vampiro"
                u = RandomNumber(1, 3) + 500
                    If u > 503 Then u = 503
                Case Else
                    u = 70
                    
        End Select
        End If
        
        If UserList(UserIndex).Char.Head = u Then
        Call SendData(ToIndex, UserIndex, 0, "||6°No puedo operar ahora, vuelva más tarde.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        Exit Sub
        End If
        
        If UserList(UserIndex).Stats.GLD > 9999 Then
             UserList(UserIndex).Char.Head = u
             UserList(UserIndex).OrigChar.Head = u
             UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 10000
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu rostro ha sido operado por 10000 oros." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & "´" & FontTypeNames.FONTTYPE_info)
            '[gau]
            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, val(u), UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
      
        Else
              Call SendData(ToIndex, UserIndex, 0, "||6°No tenes esa cantidad.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
        End If
        Call SendUserStatsOro(val(UserIndex))
        
        Exit Sub
        

Case "/TORNEO"
            Dim r10
            Dim y10
            r10 = RandomNumber(52, 71)
            y10 = RandomNumber(44, 59)
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "L3")
            Exit Sub
        End If
        'pluto:6.0A
        If UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Demonio > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes entrar transformado a Torneo.!!" & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
        
        'If EsNewbie(UserIndex) Then
            'Call SendData(ToIndex, UserIndex, 0, "||¡¡Los Newbies no pueden acceder a los Torneos.!!" & "´" & FontTypeNames.FONTTYPE_info)
            'Exit Sub
        'End If
        If UserList(UserIndex).flags.Montura > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||No se permiten Mascotas" & "´" & FontTypeNames.FONTTYPE_info)
            Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "L4")
            Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
            Call SendData(ToIndex, UserIndex, 0, "L2")
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 7)
        'pluto:6.2
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_TORNEO And Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 22 And Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 41 Then Exit Sub
        
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
            Call SendData(ToIndex, UserIndex, 0, "L2")
            Exit Sub
        End If
      
       'controla la entrada al torneo
        If UserList(UserIndex).NroMacotas > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||6°No puedes llevar mascotas al torneo.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Invisible > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||6°No puedes ir invisible al torneo.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
            Exit Sub
        End If
        
'pluto:6.2 torneo 1v1
If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = NPCTYPE_TORNEO Then
            If MapInfo(MAPATORNEO).NumUsers > 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||6°El mapa de torneo está ocupado ahora mismo.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
                Exit Sub
            Else
                If MapInfo(MAPATORNEO).NumUsers = 0 Then
                    Call SendData(ToMap, 0, 296, "||Torneo 1vs1: " & UserList(UserIndex).Name & " espera rival en la Sala De Torneos." & "´" & FontTypeNames.FONTTYPE_talk)
                    'Call SendData(ToMap, 0, 170, "||Torneo: " & UserList(UserIndex).Name & " espera rival en la Sala De Torneos." & "´" & FontTypeNames.FONTTYPE_talk)
                    '[Tite añade aviso a Bander]
                    'Call SendData(ToMap, 0, 59, "||Torneo: " & UserList(UserIndex).Name & " espera rival en la Sala De Torneos." & "´" & FontTypeNames.FONTTYPE_talk)
                End If
                If MapInfo(MAPATORNEO).NumUsers > 0 Then
                    Call SendData(ToMap, 0, 296, "||Torneo 1vs1: " & UserList(UserIndex).Name & " acepto el desafio!!!" & "´" & FontTypeNames.FONTTYPE_talk)
                    'Call SendData(ToMap, 0, 170, "||Torneo: " & UserList(UserIndex).Name & " acepto el desafio!!!" & "´" & FontTypeNames.FONTTYPE_talk)
                    Call SendData(ToMap, 0, 164, "||Torneo 1vs1: " & UserList(UserIndex).Name & " acepto el desafio!!!" & "´" & FontTypeNames.FONTTYPE_talk)
                    'Call SendData(ToMap, 0, 59, "||Torneo: " & UserList(UserIndex).Name & " acepto el desafio!!!" & "´" & FontTypeNames.FONTTYPE_talk)
                    '[/Tite]
                End If
            End If
       'manda al mapa de torneo
           ' Dim r10
           ' Dim y10
           ' r10 = RandomNumber(52, 71)
           ' y10 = RandomNumber(44, 59)
          
            Call WarpUserChar(UserIndex, MAPATORNEO, r10, y10, True)
'torneo bote
ElseIf Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 22 Then 'npctorneo bote
            If MapInfo(MapaTorneo2).NumUsers > 3 Then
                Call SendData(ToIndex, UserIndex, 0, "||6°El mapa de torneo está a tope ahora mismo.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.ELV > 30 Then
                Call SendData(ToIndex, UserIndex, 0, "||6°Tienes demasiado nivel.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.GLD < 100 Then
                Call SendData(ToIndex, UserIndex, 0, "||6°No tienes suficiente Oro.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
                Exit Sub
            End If
            'manda al mapa de torneo
              Call WarpUserChar(UserIndex, MapaTorneo2, r10, y10, True)

'pluto:2.14
            'UserList(UserIndex).flags.Morph = UserList(UserIndex).Char.Body
            'Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, val(BodyTorneo), val(0), UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)

'Call ChangeUserChar(ToMap, 0, UserList(userindex).Pos.Map, userindex, val(25), val(0), UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim, UserList(userindex).Char.Botas)

            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 100
            Call SendUserStatsOro(UserIndex)
'torneo todosvstodos
ElseIf Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 41 Then
            Call WarpUserChar(UserIndex, 293, r10, y10, True)
'torneo clanes
ElseIf Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 42 Then
  'pluto.6.8
  Exit Sub 'desactivado
  'si hay dos clanes dentro comprobamos que el user es de uno de ellos
    'TClanOcupado = 0
    If UserList(UserIndex).GuildInfo.GuildName = "" Then Exit Sub
    'pluto:6.3
    If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub
    If TClanOcupado = 2 Then
        If UserList(UserIndex).GuildInfo.GuildName <> TorneoClan(1).Nombre And UserList(UserIndex).GuildInfo.GuildName <> TorneoClan(2).Nombre Then
        Call SendData(ToIndex, UserIndex, 0, "||5°" & "Mapa ocupado: " & TorneoClan(1).Nombre & " vs " & TorneoClan(2).Nombre & "°" & Npclist(TempCharIndex).Char.CharIndex)
        Exit Sub
        Else 'si es uno de los clanes que estan dentor sumamos
            If UserList(UserIndex).GuildInfo.GuildName = TorneoClan(1).Nombre Then
            TorneoClan(1).numero = TorneoClan(1).numero + 1
            Call WarpUserChar(UserIndex, 292, r10, y10, True)
            ElseIf UserList(UserIndex).GuildInfo.GuildName = TorneoClan(2).Nombre Then
             TorneoClan(2).numero = TorneoClan(2).numero + 1
            Call WarpUserChar(UserIndex, 292, r10, y10, True)
            End If
        End If
    
    
    Else ' si hay hueco para clan nuevo
        TClanOcupado = TClanOcupado + 1
        'si el clan 1 es el nuevo..
        If TorneoClan(1).numero = 0 Then
        TorneoClan(1).Nombre = UserList(UserIndex).GuildInfo.GuildName
        TorneoClan(1).numero = TorneoClan(1).numero + 1
          Call WarpUserChar(UserIndex, 292, r10, y10, True)
        Else ' si lo es el clan 2..
        TorneoClan(2).Nombre = UserList(UserIndex).GuildInfo.GuildName
        TorneoClan(2).numero = TorneoClan(2).numero + 1
          Call WarpUserChar(UserIndex, 292, r10, y10, True)
        End If
    End If


End If 'npctype torneo
        Exit Sub
        
Case "/DDD"
'TorneoPluto.FaseTorneo = 0
If UserList(UserIndex).flags.TorneoPluto = 1 Then UserList(UserIndex).flags.TorneoPluto = 0: Exit Sub
UserList(UserIndex).flags.TorneoPluto = 1
If TorneoPluto.FaseTorneo = 0 Then Call SendData2(ToIndex, UserIndex, 0, 90)
If TorneoPluto.FaseTorneo = 1 Then Call EnviarTorneo(UserIndex)
Exit Sub
   
   
   Case "/CHISME" 'chisme


        '¿Esta el user muerto? Si es asi no puede pedir un chisme
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "L3")
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "L4")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "L2")
                  Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 7)
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_CHISMOSO _
        Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "L2")
              Exit Sub
        End If
        
        ' tiene mil oros para pagar por el chisme?
        
        If UserList(UserIndex).Stats.GLD > 999 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 1000
        'pluto:2.14
        SendUserStatsOro (UserIndex)
    
        Else
            Call SendData(ToIndex, UserIndex, 0, "||6°Por menos de 1000 oros no abro la boca...°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
            Exit Sub
        End If
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"
        ' aqui se supone que elige usuario con una etiqueta para hacer alguna llamada

eligepjnogm:
        Dim eligepj As Integer
        eligepj = RandomNumber(1, LastUser)
        ' que no sea un GM... contar chismes de gm no tiene sentido
        'pluto:6.0A
        If UserList(eligepj).flags.UserLogged = False Then GoTo eligepjnogm
        
        If UserList(eligepj).flags.Privilegios <> 0 Then GoTo eligepjnogm
        ' si es newbie tampoco... pagar para tener chismes de newbies, mejor no
        'If UserList(eligepj).Stats.ELV <= LimiteNewbie Then GoTo eligepjnogm

        ' aqui elige 2 skills aleatorios para su posible uso (es trabajo extra a la cpu si luego no se usa ese chisme...podría ponerse justo en el case...)
eligeskill:
        Dim eligeskill1 As Integer
        Dim eligeskill2 As Integer
        eligeskill1 = RandomNumber(1, NUMSKILLS)
        ' si es wrestiling o supervivencia ponemos el siguiente :P (que chapuza, navegacion y talar, saldrán mas... :PPP)
        'If eligeskill1 = 9 Or eligeskill1 = 20 Then eligeskill1 = eligeskill1 + 1
eligeskilldistinto:
        eligeskill2 = RandomNumber(1, NUMSKILLS)
        If eligeskill2 = 9 Or eligeskill2 = 20 Then eligeskill2 = eligeskill2 + 1
        ' si son iguales los dos skills elegimos otro segundo skill
        If eligeskill1 = eligeskill2 Then GoTo eligeskilldistinto
        
        ' aquí elige 2 atributos aleatorios... igual ke los skill, puede ser trabajo extra :PP
eligeatrib:
        Dim eligeatrib1 As Integer
        Dim eligeatrib2 As Integer
        eligeatrib1 = RandomNumber(1, 5)
        ' si es carisma elige mmm, constitucion que es interesante para todos...(no kiero poner un goto hacia atras)
        If eligeatrib1 = 4 Then eligeatrib1 = 5
eligeatribdistinto:
        eligeatrib2 = RandomNumber(1, 5)
        If eligeatrib2 = 4 Then GoTo eligeatribdistinto
        ' si son iguales los dos atrib elegimos otro segundo (pluto se moskeará cuando vea dos gotos para atrás casi juntos... :PP)
        If eligeatrib1 = eligeatrib2 Then GoTo eligeatribdistinto
        
        
        res = RandomNumber(1, 1000)
        ' aqui selecciona el tipo de mensaje en función del resultado aleatorio
        Select Case res
          Case Is > 950
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°Las malas lenguas dicen que " & UserList(eligepj).Name & " tiene " & UserList(eligepj).Stats.UserAtributos(1) & " de fuerza, " & UserList(eligepj).Stats.UserAtributos(2) & " de agilidad, " & UserList(eligepj).Stats.UserAtributos(3) & " de inteligencia y " & UserList(eligepj).Stats.UserAtributos(5) & " de constitución...vaya birria, no? :PP" & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Exit Sub
          Case 861 To 950
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°Me han contado que " & UserList(eligepj).Name & " sólo ha matado " & UserList(eligepj).Stats.NPCsMuertos & " monstruos, porque se lo comen vivo al tener la poquita vida de " & UserList(eligepj).Stats.MaxHP & " no me extraña...pobrecito..." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Exit Sub
          Case 781 To 860
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°¿Pero tu no sabías que " & UserList(eligepj).Name & " es " & UserList(eligepj).clase & "?..., pero si lo sabe hasta el mas new de AODrag..." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Exit Sub
          Case 691 To 780
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°...como te iba diciendo, han visto a " & UserList(eligepj).Name & " por el mapa " & UserList(eligepj).Pos.Map & "... y digo yo que qué hará por ahí... seguro que nada bueno°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Exit Sub
          Case 601 To 690
          If UserList(eligepj).Stats.GLD < 100000 Then
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°Pobre " & UserList(eligepj).Name & ", como le asalten le robarán las " & UserList(eligepj).Stats.GLD & " monedas que con tanto sudor ganó..." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Else
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°" & UserList(eligepj).Name & " se que lleva " & UserList(eligepj).Stats.GLD & " monedas encima... esa cantidad sólo se consigue haciendo maldades...¡si lo sabré yo que le conozco bien!" & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          End If
          Exit Sub
          Case 511 To 600
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°Sé de buena tinta que " & UserList(eligepj).Name & " con su level " & UserList(eligepj).Stats.ELV & " solo tiene " & UserList(eligepj).Stats.UserSkills(2) & " de magia y " & UserList(eligepj).Stats.MaxMAN & " de maná... con eso tardará dias en matar un lobo" & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Exit Sub
          Case 371 To 510
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°Me he enterado de que " & UserList(eligepj).Name & " tiene " & UserList(eligepj).Stats.UserSkills(eligeskill1) & " de " & SkillsNames(eligeskill1) & ", " & UserList(eligepj).Stats.UserSkills(eligeskill2) & " de " & SkillsNames(eligeskill2) & " y pega por " & UserList(eligepj).Stats.MaxHIT & " de cuando en cuando..." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Exit Sub
          Case 231 To 370
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°Sí... el " & UserList(eligepj).raza & " al que llaman " & UserList(eligepj).Name & ", dicen que su madre es una araña y su padre un zombie, y por eso tiene " & UserList(eligepj).Stats.UserAtributos(eligeatrib1) & " de " & AtributosNames(eligeatrib1) & ", y " & UserList(eligepj).Stats.UserAtributos(eligeatrib2) & " de " & AtributosNames(eligeatrib2) & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Exit Sub
          Case 141 To 230
          'Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°Mi vecina me ha dicho que " & UserList(eligepj).Name & " tiene " & UserList(eligepj).BancoInvent(1).NroItems & " cosas en el banco...es bastante new, porque con su LVL " & UserList(eligepj).Stats.ELV & " yo tenía muchas mas cosas...seguro que son todas pieles de lobo..." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°Sí... el " & UserList(eligepj).raza & " al que llaman " & UserList(eligepj).Name & ", dicen que su madre es una araña y su padre un zombie, y por eso tiene " & UserList(eligepj).Stats.UserAtributos(eligeatrib1) & " de " & AtributosNames(eligeatrib1) & ", y " & UserList(eligepj).Stats.UserAtributos(eligeatrib2) & " de " & AtributosNames(eligeatrib2) & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)

          Exit Sub
          Case 51 To 140
          'pluto:2.14 bug ciudas matados
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°" & UserList(eligepj).Name & " ha matado " & UserList(eligepj).Faccion.CriminalesMatados & " criminales y " & UserList(eligepj).Faccion.CiudadanosMatados & " ciudadanos... habrá que ponerle una estatua por eso?" & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Exit Sub
          Case Is < 51
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||6°A mi me llaman chismosa, pero que sepan todos que tú eres cien veces más cotilla que yo..." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex)
          Call SendData(ToAll, 0, 0, "||NOTICIA DE AODRAG: a " & UserList(UserIndex).Name & " le encantan los chismes y es cotilla de nacimiento!!!!!" & "´" & FontTypeNames.FONTTYPE_GUILD)
          Exit Sub
         End Select
'Exit Sub
' chismes

' 5%---1- el ke dice TODOS los atrib
' 9%---2- el de los npc muertos y vida total
' 8%---3- el ke dice la clase
' 9%---4- el ke dice el mapa
' 9%---5- el ke dice el oro ke lleva encima
' 9%---6- el ke dice el LVL, magia y mana
' 14%--7- el ke dice 2 skills aleatorios y el golpe maximo
' 14%--8- el ke dice la raza y dos atrib aleatorios
' 9%---9- el ke dice el numero de cosas en el banco y el LVL
' 9%---10- el ke dice los ciudad y crimis matados
' 5%---11- el ke le dice al user ke es mas cotilla ke el npc_cotilla
        
        
End Select


Exit Sub
ErrorComandoPj:
 Call LogError("TCP2. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).Name & "UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)
 Call CloseSocket(UserIndex)

End Sub
