Attribute VB_Name = "modMensajesQuest"
Sub MensajesQuest(ByVal UserIndex As Integer)
Dim Nivel As Integer
Dim raza As String
Dim NombrePJ As String
Dim asunto As String
Dim mensaje As String
Dim cuentasuma As Integer
Nivel = UserList(UserIndex).Stats.ELV
raza = UserList(UserIndex).raza
NombrePJ = UserList(UserIndex).Name


If Not FileExist(App.Path & "\MAIL\" & Left$(UCase$(NombrePJ), 1), vbDirectory) Then
    MkDir (App.Path & "\MAIL\" & Left$(UCase$(NombrePJ), 1))
End If

If Nivel = 9 Then

'¿Es Humano?
If raza = "Humano" Then
asunto = "Leihoff, Terrateniente de Montaraz."
mensaje = "Cualquier hombre o mujer de Montaraz capaz de empuñar un arma deberá presentarse al instructor encargado para iniciar la carrera militar con el objetivo de reforzar los distintos frentes de combate. Leihoff, Terrateniente de Montaraz."
cuentasuma = ""
Call MandarMensaje(NombrePJ, asunto, mensaje)
Call SendData(ToIndex, UserIndex, 0, "||¡Tienes 1 Mensaje Nuevo!." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData2(ToIndex, UserIndex, 0, 114)
End If
'Fin ¿Es Humano?

'¿Es Vampiro?
If raza = "Vampiro" Then
asunto = "Maldred, Conde de Transilvanya."
mensaje = "Compañeros en vida, compañeros en muerte, se avecinan tiempos de guerra y todo miembro no flagelado por la desdicha puede resultar un potencial apoyo en combate. Reuniros con el instructor encargado para poner en marcha la enseñanza del arte. Maldred, Conde de Transilvanya."
cuentasuma = ""
Call MandarMensaje(NombrePJ, asunto, mensaje)
Call SendData(ToIndex, UserIndex, 0, "||¡Tienes 1 Mensaje Nuevo!." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData2(ToIndex, UserIndex, 0, 114)
End If
'Fin ¿Es Vampiro?

'¿Es Elfo?
If raza = "Elfo" Then
asunto = "Archidruida Aethas de Rivendel."
mensaje = "¡Hermanos, nuestra tierra nos necesita una vez más! ¡Acudid a la llamada de la naturaleza, por Rivendel! El maestro Kir'al Cantosombrío nos guiará por el camino que el alba recorrió en su día através de la oscuridad."
cuentasuma = ""
Call MandarMensaje(NombrePJ, asunto, mensaje)
Call SendData(ToIndex, UserIndex, 0, "||¡Tienes 1 Mensaje Nuevo!." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData2(ToIndex, UserIndex, 0, 114)
End If
'Fin ¿Es Elfo?

'¿Es Enano o Gnomo?
If raza = "Enano" Or raza = "Gnomo" Then
asunto = "Sobrestante Alarik Forjatiniebla de Tínker."
mensaje = "Amigos míos El Gran Yunque lleva lustros oxidándose, reclamando el sonido de nuestros martillos, reclamando poseer nuestras armas y armaduras con una noble muerte en combate... ¡Otorguemos a nuestra tierra el honor de engullirnos una vez más, por Reox! Thorgan Fraguacero pondrá al rojo vivo vuestras habilidades. "
Call MandarMensaje(NombrePJ, asunto, mensaje)
Call SendData(ToIndex, UserIndex, 0, "||¡Tienes 1 Mensaje Nuevo!." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData2(ToIndex, UserIndex, 0, 114)
End If
'Fin ¿Es Enano o Gnomo?

'¿Es Orco, Goblin o Ciclope?
If raza = "Elfo" Or raza = "Goblin" Or raza = "Ciclope" Then
asunto = "Caudillo Borgut Rajapieles."
mensaje = "Exiliados, foragidos, desterrados y olvidados... nos hallamos entre la espada y la pared. Un nuevo mal se alza y apoyado por muchos de nuestros viejos amigos, avanza inquebrantable. ¡Demostrémosles a nuestros hermanos que están equivocados! Presentaros ante el viejo Gárgaras para recibir más instrucciones... esta vez, somos uno."
cuentasuma = ""
Call MandarMensaje(NombrePJ, asunto, mensaje)
Call SendData(ToIndex, UserIndex, 0, "||¡Tienes 1 Mensaje Nuevo!." & "´" & FontTypeNames.FONTTYPE_info)
Call SendData2(ToIndex, UserIndex, 0, 114)
End If
'Fin ¿Orco, Goblin o Ciclope?

End If


End Sub

Sub MandarMensaje(NombrePJ As String, asunto As String, mensaje As String)
Dim Cuenta As Byte
Cuenta = GetVar(App.Path & "\MAIL\" & Left$(NombrePJ, 1) & "\" & NombrePJ & ".MAIL", "INFO", "SMS")
Call WriteVar(App.Path & "\MAIL\" & Left$(NombrePJ, 1) & "\" & NombrePJ & ".MAIL", "INFO", "SMS", Cuenta + 1)
Call WriteVar(App.Path & "\MAIL\" & Left$(NombrePJ, 1) & "\" & NombrePJ & ".MAIL", "MENSAJE" & Cuenta + 1, "DE", "AOdragbot")
Call WriteVar(App.Path & "\MAIL\" & Left$(NombrePJ, 1) & "\" & NombrePJ & ".MAIL", "MENSAJE" & Cuenta + 1, "ASUNTO", asunto)
Call WriteVar(App.Path & "\MAIL\" & Left$(NombrePJ, 1) & "\" & NombrePJ & ".MAIL", "MENSAJE" & Cuenta + 1, "FECHA", Format(Now, "dd/mm/yy"))
Call WriteVar(App.Path & "\MAIL\" & Left$(NombrePJ, 1) & "\" & NombrePJ & ".MAIL", "MENSAJE" & Cuenta + 1, "MENSAJE", mensaje)
End Sub
