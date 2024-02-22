Attribute VB_Name = "TCP"
Option Explicit

'Buffer en bytes de cada socket
Public Const SOCKET_BUFFER_SIZE = 2048

'Cuantos comandos de cada cliente guarda el server
Public Const COMMAND_BUFFER_SIZE = 1000

Public Const NingunArma = 2

'RUTAS DE ENVIO DE DATOS

'PLUTO:2.15---------------
Public BytesRecibidos As Long
Public BytesEnviados As Long
Public TotalBytesRecibidos As Long
Public TotalBytesEnviados As Long
'Public BytesRecibidos As Long
'Public BytesEnviados As Long
'-----------------------------------
Public Const ToIndex = 0 'Envia a un solo User
Public Const ToAll = 1 'A todos los Users
Public Const ToMap = 2 'Todos los Usuarios en el mapa
Public Const ToPCArea = 3 'Todos los Users en el area de un user determinado
Public Const ToNone = 4 'Ninguno
Public Const ToAllButIndex = 5 'Todos menos el index
Public Const ToMapButIndex = 6 'Todos en el mapa menos el indice
Public Const ToGM = 7
Public Const ToNPCArea = 8 'Todos los Users en el area de un user determinado
Public Const ToGuildMembers = 9
Public Const ToAdmins = 10
'pluto:2.9.0
Public Const ToTorneo = 11
'pluto:2.14
Public Const ToClan = 12
'[Tite]
Public Const toParty = 13 'Miembros de la party
'[\Tite]
Public Const ToPUserAreaCercana = 14 'dos casillas
#If UsarQueSocket = 0 Then
' General constants used with most of the controls
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8

' SocketWrench Control States
Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7

' Societ Address Families
Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2

' Societ Types
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

' Protocol Types
Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256


' Network Addpesses
Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1

' SocketWrench Error Aodes
Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 2500
#End If


Public Function GenCrC(ByVal Key As Long, ByVal sdData As String) As Long

End Function

Sub DarCuerpoYCabeza(UserBody As Integer, userhead As Integer, raza As String, Gen As String)
On Error GoTo fallo
Select Case Gen
 
Case "Hombre"
        Select Case raza
                
                Case "Humano"
                    userhead = CInt(RandomNumber(1, 53))
                    If userhead = 27 Then userhead = 28
                    UserBody = 1
                
                Case "Ciclope"
                userhead = CInt(RandomNumber(1, 3)) + 800
                If userhead = 801 Then userhead = 801
                UserBody = 351
                
                Case "Elfo"
                    userhead = CInt(RandomNumber(1, 19)) + 100
                    If userhead > 119 Then userhead = 119
                    UserBody = 2
                Case "Elfo Oscuro"
                    userhead = CInt(RandomNumber(1, 16)) + 200
                    If userhead > 216 Then userhead = 216
                    UserBody = 3
                Case "Enano"
                    userhead = RandomNumber(1, 15) + 300
                    If userhead > 315 Then userhead = 315
                    UserBody = 52
                    'pluto:7.0
                    Case "Goblin"
                    userhead = RandomNumber(1, 8) + 704
                    If userhead > 712 Then userhead = 712
                    UserBody = 178
                    
                Case "Gnomo"
                    userhead = RandomNumber(1, 11) + 400
                    If userhead > 411 Then userhead = 411
                    UserBody = 52
                Case "Orco"
                    userhead = RandomNumber(1, 6) + 600
                    If userhead > 606 Then userhead = 606
                    UserBody = 218
                Case "Vampiro"
                    userhead = RandomNumber(1, 8) + 504
                    If userhead > 512 Then userhead = 512
                    UserBody = 2
                Case Else
                    userhead = 1
                    UserBody = 1
        End Select
   Case "Mujer"
        Select Case raza
                Case "Humano"
                    userhead = CInt(RandomNumber(1, 13)) + 69
                    If userhead > 82 Then userhead = 82
                    UserBody = 1
                Case "Ciclope"
                userhead = CInt(RandomNumber(1, 3)) + 800
                If userhead = 801 Then userhead = 801
                UserBody = 351
                    
                Case "Elfo"
                    userhead = CInt(RandomNumber(1, 11)) + 169
                    If userhead > 180 Then userhead = 180
                    UserBody = 2
                Case "Elfo Oscuro"
                    userhead = CInt(RandomNumber(1, 8)) + 269
                    If userhead > 277 Then userhead = 277
                    UserBody = 3
                    'pluto:7.0
                     Case "Goblin"
                    userhead = RandomNumber(1, 4) + 700
                    If userhead > 704 Then userhead = 704
                    UserBody = 212
                    
                Case "Gnomo"
                    userhead = RandomNumber(1, 7) + 469
                    If userhead > 476 Then userhead = 476
                    UserBody = 52
                Case "Enano"
                    userhead = RandomNumber(1, 4) + 369
                    If userhead > 373 Then userhead = 373
                    UserBody = 52
                Case "Orco"
                    userhead = RandomNumber(1, 3) + 606
                    If userhead > 609 Then userhead = 609
                    UserBody = 219
                Case "Vampiro"
                    userhead = RandomNumber(1, 3) + 500
                    If userhead > 503 Then userhead = 503
                    UserBody = 3
                Case Else
                    userhead = 70
                    UserBody = 1
        End Select
End Select

   Exit Sub
fallo:
Call LogError("darcuerpoycabeza " & Err.number & " D: " & Err.Description)

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
On Error GoTo fallo
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) And (car <> 64) And (car <> 46) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True
Exit Function
fallo:
Call LogError("asciivalidos " & Err.number & " D: " & Err.Description)

End Function
Function AsciiDescripcion(ByVal cad As String) As Boolean
On Error GoTo fallo
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    If (car < 97 Or car > 122) And (car <> 241) And (car <> 255) And (car <> 32) And (car <> 64) And (car <> 46) Then
        AsciiDescripcion = False
        Exit Function
    End If
    
Next i

AsciiDescripcion = True
Exit Function
fallo:
Call LogError("asciidescripcion " & Err.number & " D: " & Err.Description)

End Function
Sub SendBot(Desc As String)
On Error GoTo fallo
Dim Tindex As Integer
Tindex = NameIndex("AoDraGBoT")
If Tindex > 0 Then
Call SendData(ToIndex, Tindex, 0, "||" & Desc & "´" & FontTypeNames.FONTTYPE_TALK)
End If

Exit Sub
fallo:
Call LogError("SendBot " & Err.number & " D: " & Err.Description)
End Sub
Function Numeric(ByVal cad As String) As Boolean
On Error GoTo fallo
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True
Exit Function
fallo:
Call LogError("numeric " & Err.number & " D: " & Err.Description)

End Function


Function NombrePermitido(ByVal Nombre As String) As Boolean
On Error GoTo fallo
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True
Exit Function
fallo:
Call LogError("nombrepermitido " & Err.number & " D: " & Err.Description)

End Function

Function ValidateAtrib(ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
Dim loopc As Integer

'For loopc = 1 To NUMATRIBUTOS
 '   If UserList(UserIndex).Stats.UserAtributos(loopc) > 18 Or UserList(UserIndex).Stats.UserAtributos(loopc) < 1 Then Exit Function
'Next loopc

ValidateAtrib = True
Exit Function
fallo:
Call LogError("validateatrib " & Err.number & " D: " & Err.Description)

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
Dim loopc As Integer

For loopc = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(loopc) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(loopc) > 200 Then UserList(UserIndex).Stats.UserSkills(loopc) = 200
    End If
Next loopc

ValidateSkills = True
    
Exit Function
fallo:
Call LogError("validateskills " & Err.number & " D: " & Err.Description)

End Function

Sub ConnectNewUser(UserIndex As Integer, Name As String, Password As String, Body As Integer, Head As Integer, UserRaza As String, UserSexo As String, UserClase As String, _
UA1 As String, UA2 As String, UA3 As String, UA4 As String, UA5 As String, _
US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
US21 As String, US22 As String, US23 As String, US24 As String, US25 As String, _
US26 As String, US27 As String, US28 As String, US29 As String, US30 As String, _
US31 As String, UserEmail As String, Hogar As String, Totalda As Integer, P1 As Byte, P2 As Byte, P3 As Byte, P4 As Byte, P5 As Byte, P6 As Byte)
On Error GoTo fallo


If Not NombrePermitido(Name) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Los nombres de los personajes deben pertencer a la fantasia, el nombre indicado es invalido.")
    Exit Sub
End If
'pluto:6.7
If Left$(Name, 1) = " " Or Right$(Name, 1) = " " Then
Call SendData2(ToIndex, UserIndex, 0, 79, UserIndex)
Call LogError("Intento Nombre con Espacio: " & Name & " Ip:" & UserList(UserIndex).ip)
Exit Sub
End If


If Len(Name) > 15 Or Len(Name) < 4 Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Nombre demasiado largo o demasiado corto.")
    Exit Sub
End If

If Not AsciiValidos(Name) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Nombre invalido.")
    Exit Sub
End If

Dim loopc As Integer
Dim totalskpts As Long
  
'¿Existe el personaje?
If PersonajeExiste(Name) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Ya existe el personaje.")
    Exit Sub
End If

'pluto:6.0A
Call SendData(ToAdmins, UserIndex, 0, "|| Creado Pj : " & Name & "´" & FontTypeNames.FONTTYPE_TALK)

'pluto:5.2
UserList(UserIndex).flags.CMuerte = 1
'--------
UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).flags.Escondido = 0
UserList(UserIndex).flags.Protec = 0
UserList(UserIndex).flags.Ron = 0
UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.BurguesRep = 0
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.NobleRep = 1000
UserList(UserIndex).Reputacion.PlebeRep = 30

UserList(UserIndex).Reputacion.Promedio = 30 / 6

UserList(UserIndex).Name = Name
UserList(UserIndex).clase = UserClase
UserList(UserIndex).raza = UserRaza
UserList(UserIndex).Genero = UserSexo
UserList(UserIndex).Email = Cuentas(UserIndex).mail
UserList(UserIndex).Hogar = Hogar
'pluto:2.14 --------------------
UserList(UserIndex).Padre = ""
UserList(UserIndex).Madre = ""

UserList(UserIndex).Nhijos = 0
Dim X As Byte
For X = 1 To 5
UserList(UserIndex).Hijo(X) = ""
Next
'-------------------------------

If Abs(CInt(UA1)) + Abs(CInt(UA2)) + Abs(CInt(UA3)) + Abs(CInt(UA4)) + Abs(CInt(UA5)) > 105 Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Atributos invalidos.")
    Exit Sub
End If

UserList(UserIndex).Stats.UserAtributos(Fuerza) = Abs(CInt(UA1))
UserList(UserIndex).Stats.UserAtributos(Inteligencia) = Abs(CInt(UA2))
UserList(UserIndex).Stats.UserAtributos(Agilidad) = Abs(CInt(UA3))
UserList(UserIndex).Stats.UserAtributos(Carisma) = Abs(CInt(UA4))
UserList(UserIndex).Stats.UserAtributos(Constitucion) = Abs(CInt(UA5))
'pluto:7.0
If (P1 + P2 + P3 + P4 + P5 + P6 > 15) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Atributos invalidos.")
    Exit Sub
End If

UserList(UserIndex).UserDañoProyetilesRaza = P1
UserList(UserIndex).UserDañoArmasRaza = P2
UserList(UserIndex).UserDañoMagiasRaza = P3
UserList(UserIndex).UserDefensaMagiasRaza = P4
UserList(UserIndex).UserEvasiónRaza = P5
UserList(UserIndex).UserDefensaEscudos = P6

UserList(UserIndex).Remort = 0
UserList(UserIndex).Remorted = ""

'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS ATRIBUTOS %%%%%%%%%%%%%
If Not ValidateAtrib(UserIndex) Then
        Call SendData2(ToIndex, UserIndex, 0, 43, "Atributos invalidos.")
        Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS ATRIBUTOS %%%%%%%%%%%%%

'pluto:7.0 quito todo esto para la nueva versión
'Select Case UCase$(UserRaza)
 '   Case "HUMANO"
  '      UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 1
   '     UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 2
    '    UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 2
     '   UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 1
   ' Case "ELFO"
    '    UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 2
     '   UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 2
      '  UserList(UserIndex).Stats.UserAtributos(Carisma) = UserList(UserIndex).Stats.UserAtributos(Carisma) + 2
       ' UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 1
        'UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) - 1

 '   Case "ELFO OSCURO"
  '      UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 2
   '     UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) - 2
    '    UserList(UserIndex).Stats.UserAtributos(Carisma) = UserList(UserIndex).Stats.UserAtributos(Carisma) + 2
     '   UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 1
      '  UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 1
    'Case "ENANO"
     '   UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 3
      '  UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 3
        'pluto:6.0A cambio enano a -3 inte
       ' UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) - 3
        ' UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) - 1

   ' Case "GNOMO"
   '     UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) - 4
    '    UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 3
    '    UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 3
    '    UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 1

   ' Case "ORCO"
    '    UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 4
     '   UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) - 3
      '  UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 3
       ' UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) - 6
   ' Case "VAMPIRO"
   '     UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 2
    '    UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 2
     '   UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 2
           ' End Select

UserList(UserIndex).Stats.UserSkills(1) = val(US1)
UserList(UserIndex).Stats.UserSkills(2) = val(US2)
UserList(UserIndex).Stats.UserSkills(3) = val(US3)
UserList(UserIndex).Stats.UserSkills(4) = val(US4)
UserList(UserIndex).Stats.UserSkills(5) = val(US5)
UserList(UserIndex).Stats.UserSkills(6) = val(US6)
UserList(UserIndex).Stats.UserSkills(7) = val(US7)
UserList(UserIndex).Stats.UserSkills(8) = val(US8)
UserList(UserIndex).Stats.UserSkills(9) = val(US9)
UserList(UserIndex).Stats.UserSkills(10) = val(US10)
UserList(UserIndex).Stats.UserSkills(11) = val(US11)
UserList(UserIndex).Stats.UserSkills(12) = val(US12)
UserList(UserIndex).Stats.UserSkills(13) = val(US13)
UserList(UserIndex).Stats.UserSkills(14) = val(US14)
UserList(UserIndex).Stats.UserSkills(15) = val(US15)
UserList(UserIndex).Stats.UserSkills(16) = val(US16)
UserList(UserIndex).Stats.UserSkills(17) = val(US17)
UserList(UserIndex).Stats.UserSkills(18) = val(US18)
UserList(UserIndex).Stats.UserSkills(19) = val(US19)
UserList(UserIndex).Stats.UserSkills(20) = val(US20)
UserList(UserIndex).Stats.UserSkills(21) = val(US21)
UserList(UserIndex).Stats.UserSkills(22) = val(US22)
UserList(UserIndex).Stats.UserSkills(23) = val(US23)
UserList(UserIndex).Stats.UserSkills(24) = val(US24)
UserList(UserIndex).Stats.UserSkills(25) = val(US25)
UserList(UserIndex).Stats.UserSkills(26) = val(US26)
UserList(UserIndex).Stats.UserSkills(27) = val(US27)
UserList(UserIndex).Stats.UserSkills(28) = val(US28)
UserList(UserIndex).Stats.UserSkills(29) = val(US29)
UserList(UserIndex).Stats.UserSkills(30) = val(US30)
UserList(UserIndex).Stats.UserSkills(31) = val(US31)
totalskpts = 10
UserList(UserIndex).Stats.SkillPts = 10
' PREVINENE EL HACKEO DE LOS SKILLS %%%%%%%%%%%%%
For loopc = 1 To NUMSKILLS
If UserList(UserIndex).Stats.UserSkills(loopc) > 0 Then
Call LogError(" en Jugador:" & UserList(UserIndex).Name & " Skills Trucados " & "Ip: " & UserList(UserIndex).ip)
End If
Next loopc


'If totalskpts > 10 Then
 '   Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los skills.")
'    Call BorrarUsuario(UserList(userindex).name)
  '  Call CloseUser(UserIndex)
   ' Exit Sub
'End If

'pluto:2.14
'If Totalda > (UserList(UserIndex).Stats.UserAtributos(1) + UserList(UserIndex).Stats.UserAtributos(2) + UserList(UserIndex).Stats.UserAtributos(3) + UserList(UserIndex).Stats.UserAtributos(4) + UserList(UserIndex).Stats.UserAtributos(5)) Then
'Call LogCasino("Jugador:" & UserList(UserIndex).Name & " posibles Dados trucados " & "Ip: " & UserList(UserIndex).ip)
'End If


'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

'UserList(userindex).password = password
UserList(UserIndex).Char.Heading = SOUTH

Call Randomize(Timer)
Call DarCuerpoYCabeza(UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).raza, UserList(UserIndex).Genero)
UserList(UserIndex).OrigChar = UserList(UserIndex).Char
   
 
UserList(UserIndex).Char.WeaponAnim = NingunArma
UserList(UserIndex).Char.ShieldAnim = NingunEscudo
UserList(UserIndex).Char.CascoAnim = NingunCasco

UserList(UserIndex).Stats.MET = 1
Dim MiInt
MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 3)
'MiInt = UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 3
UserList(UserIndex).Stats.MaxHP = 15 + MiInt
UserList(UserIndex).Stats.MinHP = 15 + MiInt

UserList(UserIndex).Stats.FIT = 1


MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(UserIndex).Stats.MaxSta = 20 * MiInt
UserList(UserIndex).Stats.MinSta = 20 * MiInt


UserList(UserIndex).Stats.MaxAGU = 100
UserList(UserIndex).Stats.MinAGU = 100

UserList(UserIndex).Stats.MaxHam = 100
UserList(UserIndex).Stats.MinHam = 100


'<-----------------MANA----------------------->
If UserClase = "Mago" Then
    MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Inteligencia)) / 3
    UserList(UserIndex).Stats.MaxMAN = 100 + MiInt
    UserList(UserIndex).Stats.MinMAN = 100 + MiInt
ElseIf UserClase = "Clerigo" Or UserClase = "Druida" _
    Or UserClase = "Bardo" Or UserClase = "Asesino" Or UserClase = "Pirata" Then
        MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Inteligencia)) / 4
        UserList(UserIndex).Stats.MaxMAN = 50
        UserList(UserIndex).Stats.MinMAN = 50
Else
    UserList(UserIndex).Stats.MaxMAN = 0
    UserList(UserIndex).Stats.MinMAN = 0
End If

If UserClase = "Mago" Or UserClase = "Clerigo" Or _
   UserClase = "Druida" Or UserClase = "Bardo" Or _
   UserClase = "Asesino" Then
        UserList(UserIndex).Stats.UserHechizos(1) = 2
End If

UserList(UserIndex).Stats.MaxHIT = 2
UserList(UserIndex).Stats.MinHIT = 1
UserList(UserIndex).Stats.Fama = 0
UserList(UserIndex).Stats.GLD = 0
UserList(UserIndex).Stats.LibrosUsados = 0
UserList(UserIndex).Stats.exp = 0
UserList(UserIndex).Stats.Elu = 300
UserList(UserIndex).Stats.ELV = 1

'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
UserList(UserIndex).Invent.NroItems = 4

UserList(UserIndex).Invent.Object(1).ObjIndex = 467
UserList(UserIndex).Invent.Object(1).Amount = 100

UserList(UserIndex).Invent.Object(2).ObjIndex = 468
UserList(UserIndex).Invent.Object(2).Amount = 100

UserList(UserIndex).Invent.Object(3).ObjIndex = 460
UserList(UserIndex).Invent.Object(3).Amount = 1
UserList(UserIndex).Invent.Object(3).Equipped = 1
'pluto:7.0--- añado arcos y flechas newbies-------
If UserClase = "Arquero" Or UserClase = "Cazador" Or UserClase = "Leñador" Or UserClase = "Minero" Or _
UserClase = "Pescador" Or UserClase = "Ermitaño" Or UserClase = "Domador" Or UserClase = "Carpintero" Or _
UserClase = "Herrero" Then
UserList(UserIndex).Invent.Object(5).ObjIndex = 1280
UserList(UserIndex).Invent.Object(5).Amount = 1
UserList(UserIndex).Invent.Object(5).Equipped = 0
UserList(UserIndex).Invent.Object(6).ObjIndex = 1281
UserList(UserIndex).Invent.Object(6).Amount = 500
UserList(UserIndex).Invent.Object(6).Equipped = 0
End If
'---------------------------------------------------
Select Case UserRaza
    Case "Humano"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 463
    Case "Elfo"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 464
    Case "Elfo Oscuro"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 465
    Case "Enano"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
    Case "Gnomo"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
 Case "Vampiro"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 465
  Case "Orco"
       If UserList(UserIndex).Genero = "Mujer" Then
       UserList(UserIndex).Invent.Object(4).ObjIndex = 737
    Else
    UserList(UserIndex).Invent.Object(4).ObjIndex = 736
    End If
    'pluto:7.0
Case "Goblin"
UserList(UserIndex).Invent.Object(4).ObjIndex = 466
Case "Ciclope"
UserList(UserIndex).Invent.Object(4).ObjIndex = 464
End Select

UserList(UserIndex).Invent.Object(4).Amount = 1
UserList(UserIndex).Invent.Object(4).Equipped = 1

UserList(UserIndex).Invent.ArmourEqpSlot = 4
UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(4).ObjIndex

UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(3).ObjIndex
UserList(UserIndex).Invent.WeaponEqpSlot = 3


Call SaveUser(UserIndex, CharPath & Left$(UCase$(Name), 1) & "\" & UCase$(Name) & ".chr")

 
'Open User
'Call ConnectUser(userindex, name, password)
    Cuentas(UserIndex).NumPjs = Cuentas(UserIndex).NumPjs + 1
    ReDim Preserve Cuentas(UserIndex).Pj(1 To Cuentas(UserIndex).NumPjs)
    Cuentas(UserIndex).Pj(Cuentas(UserIndex).NumPjs) = Name

'pluto:6.6----------
Call ResetUserSlot(UserIndex)
'--------------------
    Call MandaPersonajes(UserIndex)

'pluto:2.4.5
'Dim x As Integer
    Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "DATOS", "Password", Cuentas(UserIndex).passwd)
    Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "DATOS", "NumPjs", CStr(Cuentas(UserIndex).NumPjs))
    Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "DATOS", "Llave", CStr(Cuentas(UserIndex).Llave))
    For X = 1 To Cuentas(UserIndex).NumPjs
    Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "PERSONAJES", "PJ" & X, Cuentas(UserIndex).Pj(X))
    Next
    'pluto:6.0A
Call SendData(ToIndex, UserIndex, 0, "AWIntro")

Exit Sub
fallo:
Call LogError("connectnewuser " & Err.number & " D: " & Err.Description)

End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

Dim loopc As Integer

'Call LogTarea("Close Socket")

'#If UsarQueSocket = 0 Or UsarQueSocket = 2 Then
On Error GoTo errhandler
'#End If
    
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

    Call aDos.RestarConexion(UserList(UserIndex).ip)
    

    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)
    End If
    
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
    
    
    
    If UserList(UserIndex).flags.UserLogged Then
            'If NumUsers > 0 Then NumUsers = NumUsers - 1
            Call CloseUser(UserIndex)
        
            'Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Else
Call ResetUserSlot(UserIndex)
UserList(UserIndex).ip = ""
UserList(UserIndex).RDBuffer = ""
    End If
   
    
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

'    #If UsarQueSocket = 1 Then
'
'    If UserList(UserIndex).ConnID <> -1 Then
'        Call CloseSocketSL(UserIndex)
'    End If
'
'    #ElseIf UsarQueSocket = 0 Then
'
'    'frmMain.Socket2(UserIndex).D i s c o n n e c t   NO USAR
'    frmMain.Socket2(UserIndex).Cleanup
'    Unload frmMain.Socket2(UserIndex)
'
'    #ElseIf UsarQueSocket = 2 Then
'
'    If UserList(UserIndex).ConnID <> -1 Then
'        Call CloseSocketSL(UserIndex)
'    End If
'
'    #End If

'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
    
     'pluto fusion
    Call DesconectaCuenta(UserIndex)
    UserList(UserIndex).flags.ValCoDe = 0
    '-----------------------------
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
            
Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'    Unload frmMain.Socket2(UserIndex) OJOOOOOOOOOOOOOOOOO
'    If NumUsers > 0 Then NumUsers = NumUsers - 1
    'pluto fusion
    Call DesconectaCuenta(UserIndex)
    '-----------------------------
    Call ResetUserSlot(UserIndex)
UserList(UserIndex).ip = ""
UserList(UserIndex).RDBuffer = ""
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

    #If UsarQueSocket = 1 Then
    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)
'        Call apiclosesocket(UserList(UserIndex).ConnID)
    End If
    #End If
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
    
End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal UserIndex As Integer)

'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

'Call LogTarea("Close Socket")

On Error GoTo errhandler

'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>


    Call aDos.RestarConexion(frmMain.Socket2(UserIndex).PeerAddress)

    UserList(UserIndex).ConnID = -1
'    GameInputMapArray(UserIndex) = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0

'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

    If UserIndex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

    If UserList(UserIndex).flags.UserLogged Then
            'If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call CloseUser(UserIndex)
        End If

'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

    frmMain.Socket2(UserIndex).Cleanup
'    frmMain.Socket2(UserIndex).Di    s  c o       n nect
    Unload frmMain.Socket2(UserIndex)
    Call ResetUserSlot(UserIndex)
UserList(UserIndex).flags.ValCoDe = 0

'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
'    GameInputMapArray(UserIndex) = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'    Unload frmMain.Socket2(UserIndex) OJOOOOOOOOOOOOOOOOO
'    If NumUsers > 0 Then NumUsers = NumUsers - 1
    Call ResetUserSlot(UserIndex)

'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

End Sub







#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    'pluto:2.14
    If UserIndex = 0 Then Exit Sub
    CoNnEcTiOnId = UserList(UserIndex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    Call aDos.RestarConexion(UserList(UserIndex).ip)
    
    UserList(UserIndex).ConnID = -1 'inabilitamos operaciones en socket
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0

    If UserIndex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).ConnID = -1
    End If

    If UserList(UserIndex).flags.UserLogged Then
            'If NumUsers <> 0 Then NumUsers = NumUsers - 1
            NURestados = True
            Call CloseUser(UserIndex)
    End If
    'pluto:2.13
    If Cuentas(UserIndex).Logged = True Then
            'If NumUsers <> 0 Then NumUsers = NumUsers - 1
            NURestados = True
            Call DesconectaCuenta(UserIndex)
    End If
    
    Call ResetUserSlot(UserIndex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

errhandler:
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    Call LogError("CLOSESOCKETERR: " & Err.Description & " UI:" & UserIndex)
    If Not NURestados Then
        If UserList(UserIndex).flags.UserLogged Then
            If NumUsers <> 0 Then
                NumUsers = NumUsers - 1
            End If
    
            Call LogError("Cerre sin grabar a: " & UserList(UserIndex).Name)
        End If
    End If
    Call LogError("El usuario no guardado tenia connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(UserIndex)
'pluto:2.13
    If Cuentas(UserIndex).Logged = True Then
    Call DesconectaCuenta(UserIndex)
    End If
End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
Debug.Print "CloseSocketSL"

#If UsarQueSocket = 1 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    
    Call BorraSlotSock(UserList(UserIndex).ConnID)
'    Call WSAAsyncSelect(UserList(UserIndex).ConnID, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
'    Call apiclosesocket(UserList(UserIndex).ConnID)
    'pluto fusion
    Call DesconectaCuenta(UserIndex)
    UserList(UserIndex).flags.ValCoDe = 0
    '-----------------------------
    
    Call WSApiCloseSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    'frmMain.Socket2(UserIndex).Disconnect
    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#End If
End Sub

'Sub CloseSocket_NUEVA(ByVal UserIndex As Integer)
''<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
''<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
''<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'
''Call LogTarea("Close Socket")
'
'on error GoTo errhandler
'
'
'
'    Call aDos.RestarConexion(frmMain.Socket2(UserIndex).PeerAddress)
'
'    'UserList(UserIndex).ConnID = -1
'    'UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'
'    If UserList(UserIndex).flags.UserLogged Then
'        If NumUsers <> 0 Then NumUsers = NumUsers - 1
'        Call CloseUser(UserIndex)
'        UserList(UserIndex).ConnID = -1: UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'        frmMain.Socket2(UserIndex).Disconnect
'        frmMain.Socket2(UserIndex).Cleanup
'        'Unload frmMain.Socket2(UserIndex)
'        Call ResetUserSlot(UserIndex)
'        'Call Cerrar_Usuario(UserIndex)
'    Else
'        UserList(UserIndex).ConnID = -1
'        UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'
'        frmMain.Socket2(UserIndex).Disconnect
'        frmMain.Socket2(UserIndex).Cleanup
'        Call ResetUserSlot(UserIndex)
'        'Unload frmMain.Socket2(UserIndex)
'    End If
'
'Exit Sub
'
'errhandler:
'    UserList(UserIndex).ConnID = -1
'    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
''    Unload frmMain.Socket2(UserIndex) OJOOOOOOOOOOOOOOOOO
''    If NumUsers > 0 Then NumUsers = NumUsers - 1
'    Call ResetUserSlot(UserIndex)
'
''<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
''<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
''<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'
'End Sub

Public Function EnviarDatosASlot(ByVal UserIndex As Integer, Datos As String) As Long
'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)

'TCPESStats.BytesEnviados = TCPESStats.BytesEnviados + Len(Datos)

#If UsarQueSocket = 1 Then '**********************************************
On Error GoTo Err

Dim Ret As Long

If frmMain.SUPERLOG.value = 1 Then LogCustom ("EnviarDatosASlot:: INICIO. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida)

Ret = WsApiEnviar(UserIndex, Datos)

If frmMain.SUPERLOG.value = 1 Then LogCustom ("EnviarDatosASlot:: INICIO. Acabo de enviar userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida & " RET=" & Ret)

If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
    If frmMain.SUPERLOG.value = 1 Then LogCustom ("EnviarDatosASlot:: Entro a manejo de error. <> wsaewouldblock, <>0. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida)
    Call CloseSocketSL(UserIndex)
    If frmMain.SUPERLOG.value = 1 Then LogCustom ("EnviarDatosASlot:: Luego de Closesocket. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida)
    Call Cerrar_Usuario(UserIndex)
    If frmMain.SUPERLOG.value = 1 Then LogCustom ("EnviarDatosASlot:: Luego de Cerrar_usuario. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida)
End If
EnviarDatosASlot = Ret
Exit Function

Err:
    If frmMain.SUPERLOG.value = 1 Then LogCustom ("EnviarDatosASlot:: ERR Handler. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida & " ERR: " & Err.Description)

#ElseIf UsarQueSocket = 0 Then '**********************************************

Dim Encolar As Boolean
Encolar = False

EnviarDatosASlot = 0

'Dim fR As Integer
'fR = FreeFile
'Open "c:\log.txt" For Append As #fR
'Print #fR, Datos
'Close #fR
'Call frmMain.Socket2(UserIndex).Write(Datos, Len(Datos))

'If frmMain.Socket2(UserIndex).IsWritable And UserList(UserIndex).SockPuedoEnviar Then
If UserList(UserIndex).ColaSalida.Count <= 0 Then
    If frmMain.Socket2(UserIndex).Write(Datos, Len(Datos)) < 0 Then
        If frmMain.Socket2(UserIndex).LastError = WSAEWOULDBLOCK Then
            UserList(UserIndex).SockPuedoEnviar = False
            Encolar = True
        Else
            Call Cerrar_Usuario(UserIndex)
        End If
'    Else
'        Debug.Print UserIndex & ": " & Datos
    End If
Else
    Encolar = True
End If

If Encolar Then
    Debug.Print "Encolando..."
    UserList(UserIndex).ColaSalida.Add Datos
End If

#ElseIf UsarQueSocket = 2 Then '**********************************************

Dim Encolar As Boolean
Dim Ret As Long
Encolar = False

'//
'// Valores de retorno:
'//                     0: Todo OK
'//                     1: WSAEWOULDBLOCK
'//                     2: Error critico
'//
If UserList(UserIndex).ColaSalida.Count <= 0 Then
    Ret = frmMain.Serv.Enviar(UserList(UserIndex).ConnID, Datos, Len(Datos))
    If Ret = 1 Then
        Encolar = True
    ElseIf Ret = 2 Then
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If
Else
    Encolar = True
End If

If Encolar Then
    Debug.Print "Encolando..."
    UserList(UserIndex).ColaSalida.Add Datos
End If

#ElseIf UsarQueSocket = 3 Then
Dim rv As Long
'al carajo, esto encola solo!!! che, me aprobará los
'parciales también?, este control hace todo solo!!!!
On Error GoTo ErrorHandler
    If UserList(UserIndex).ConnID = -1 Then
        Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
        Exit Function
    End If
    rv = frmMain.TCPServ.Enviar(UserList(UserIndex).ConnID, Datos, Len(Datos))
'pluto:6.7---------------------
'UserList(UserIndex).Counters.UserEnvia = UserList(UserIndex).Counters.UserEnvia + 1
'If UserList(UserIndex).Counters.UserEnvia > 50 Then UserList(UserIndex).Counters.UserEnvia = 1
'----------------------------
    'If InStr(1, Datos, "VAL", vbTextCompare) > 0 Or InStr(1, Datos, "LOG", vbTextCompare) > 0 Or InStr(1, Datos, "FINO", vbTextCompare) > 0 Or InStr(1, Datos, "ERR", vbTextCompare) > 0 Then
        'call logindex(UserIndex, "SendData. ConnId: " & UserList(UserIndex).ConnID & " Datos: " & Datos)
    'End If
    Select Case rv
        'Case 1  'error critico, se viene el on_close
        Case 2  'Socket Invalido.
            'intentemos cerrarlo?
            Call CloseSocket(UserIndex, True)
        'Case 3  'WSAEWOULDBLOCK. Solo si Encolar=False en el control
            'aca hariamos manejo de encoladas, pero el server se encarga solo :D
    End Select

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & UserIndex & "/" & UserList(UserIndex).ConnID & "/" & Datos)
#End If '**********************************************

End Function

Sub SendData2(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ID As Byte, Optional ByVal Param As String = "")
 On Error GoTo fallo
    Call SendData(sndRoute, sndIndex, sndMap, Chr$(5) & Chr$(ID) & Param)
Exit Sub
fallo:
Call LogError("sendata2 " & Err.number & " D: " & Err.Description)


End Sub

Sub SendData(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, sndData As String)
On Error GoTo fallo
Dim loopc As Integer
Dim X As Integer
Dim Y As Integer
Dim aux$
Dim dec$
Dim nfile As Integer
Dim Ret As Long
Dim sndDato As String
Dim aa As String
Dim bb As String
bb = sndData
sndData = sndData & ENDC
aa = sndData
'pluto:2.8.0
'DoEvents

'If sndIndex = 0 Then GoTo nop


Select Case sndRoute
    Case ToNone
        Exit Sub
    Case ToAdmins
        For loopc = 1 To LastUser
            If UserList(loopc).ConnID > -1 And UserList(loopc).flags.UserLogged = True Then
               If EsDios(UserList(loopc).Name) Or EsSemiDios(UserList(loopc).Name) Then
  'pluto:2.10
If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap
                  
                    'pluto:2.5.0
sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap:
                    BytesEnviados = BytesEnviados + Len(sndData)
                    Call EnviarDatosASlot(loopc, sndData)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
               sndData = aa
               End If
            End If
        Next loopc
        Exit Sub
     
     'pluto:2-3-04
        Case ToGM
        For loopc = 1 To LastUser
            If UserList(loopc).ConnID > -1 And UserList(loopc).flags.UserLogged = True Then
               If UserList(loopc).flags.Privilegios > 2 Then
                      
   'pluto:2.10
If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap1
                    
                      'pluto:2.5.0
sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap1:
                        BytesEnviados = BytesEnviados + Len(sndData)
                        Call EnviarDatosASlot(loopc, sndData)
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
               sndData = aa
               End If
            End If
        Next loopc
        Exit Sub
'pluto:2.9.0

      Case ToTorneo
        For loopc = 1 To LastUser
            If (UserList(loopc).ConnID > -1) Then
                If UserList(loopc).flags.UserLogged Then
                    If UserList(loopc).flags.TorneoPluto > 0 Then
 'pluto:2.10
If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap2

'pluto:2.5.0
sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap2:
                        BytesEnviados = BytesEnviados + Len(sndData)
                        Call EnviarDatosASlot(loopc, sndData)
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                    sndData = aa
                    End If
                End If
            End If
        Next loopc
        Exit Sub
'[Tite]Msg a party
      Case toParty
        For loopc = 1 To LastUser
            If (UserList(loopc).ConnID > -1) Then
                If UserList(loopc).flags.UserLogged Then
                    If UserList(loopc).flags.party = True And UserList(loopc).flags.partyNum = UserList(sndIndex).flags.partyNum Then
                        If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap10
                        sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap10:
                        BytesEnviados = BytesEnviados + Len(sndData)
                        Call EnviarDatosASlot(loopc, sndData)
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        sndData = aa
                    End If
                End If
            End If
        Next loopc
        Exit Sub
'[\Tite]
    Case ToAll
        For loopc = 1 To LastUser
            If UserList(loopc).ConnID > -1 Then
                If UserList(loopc).flags.UserLogged Then 'Esta logeado como usuario?
  'pluto:2.10
If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap3
                   
                    'pluto:2.5.0
sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap3:
                    BytesEnviados = BytesEnviados + Len(sndData)
                    Call EnviarDatosASlot(loopc, sndData)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                sndData = aa
                End If
            End If
        Next loopc
        Exit Sub
    
    Case ToAllButIndex
        For loopc = 1 To LastUser
            If (UserList(loopc).ConnID > -1) And (loopc <> sndIndex) Then
                If UserList(loopc).flags.UserLogged Then 'Esta logeado como usuario?
       'pluto:2.10
If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap4
          
                'pluto:2.5.0
sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap4:
                    BytesEnviados = BytesEnviados + Len(sndData)
                    Call EnviarDatosASlot(loopc, sndData)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                sndData = aa
                End If
            End If
        Next loopc
        Exit Sub
    
    Case ToMap
        For loopc = 1 To LastUser
            If (UserList(loopc).ConnID > -1) Then
                If UserList(loopc).flags.UserLogged Then
                    If UserList(loopc).Pos.Map = sndMap Then
 'pluto:2.10
If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap5

'pluto:2.5.0
sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap5:
                        BytesEnviados = BytesEnviados + Len(sndData)
                        Call EnviarDatosASlot(loopc, sndData)
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                    sndData = aa
                    End If
                End If
            End If
        Next loopc
        Exit Sub
      
    Case ToMapButIndex
        For loopc = 1 To LastUser
            If (UserList(loopc).ConnID > -1 And UserList(loopc).flags.UserLogged = True) And loopc <> sndIndex Then
                If UserList(loopc).Pos.Map = sndMap Then
     'pluto:2.10
If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap6
                
                    'pluto:2.5.0
sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap6:
                    BytesEnviados = BytesEnviados + Len(sndData)
                    Call EnviarDatosASlot(loopc, sndData)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                sndData = aa
                End If
            End If
        Next loopc
        Exit Sub
            
    Case ToGuildMembers
        For loopc = 1 To LastUser
            If (UserList(loopc).ConnID > -1) And UserList(loopc).flags.UserLogged = True Then
               If UserList(sndIndex).GuildInfo.GuildName = UserList(loopc).GuildInfo.GuildName Then
               ' If UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then _
                    BytesEnviados = BytesEnviados + Len(sndData)
   'pluto:2.10
If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap7
                  
                    'pluto:2.5.0
sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap7:
                    BytesEnviados = BytesEnviados + Len(sndData)
                    Call EnviarDatosASlot(loopc, sndData)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
            sndData = aa
            End If
            End If
        Next loopc
        Exit Sub
    'pluto:6.8-----torneo de clanes----------------------------
        Case ToClan
        For loopc = 1 To LastUser
            If (UserList(loopc).ConnID > -1) And UserList(loopc).flags.UserLogged = True Then
               If TorneoClan(1).Nombre = UserList(loopc).GuildInfo.GuildName Or TorneoClan(2).Nombre = UserList(loopc).GuildInfo.GuildName Then

 sndData = "|," & sndData
If UserList(loopc).Name = "AoDraGBoT" Or UserList(loopc).Name = "AoDraGBoT2" Then GoTo nap17
                  
                
sndData = CodificaR(str$(UserList(loopc).flags.ValCoDe), sndData, loopc, 1)
nap17:
                    BytesEnviados = BytesEnviados + Len(sndData)
                    Call EnviarDatosASlot(loopc, sndData)
                
            sndData = aa
            End If
            End If
        Next loopc
        Exit Sub
   '--------------------------------------------
    
    
    
    'pluto:2.14--------------------------------
    
'Case ToClan
'For LoopC = 1 To LastUser
'If (UserList(LoopC).ConnID > -1) And UserList(LoopC).flags.UserLogged = True Then


'If (bb = "C1" Or bb = "C5") And UserList(LoopC).GuildInfo.GuildName <> castillo1 Then GoTo npp
'If (bb = "C2" Or bb = "C6") And UserList(LoopC).GuildInfo.GuildName <> castillo2 Then GoTo npp
'If (bb = "C3" Or bb = "C7") And UserList(LoopC).GuildInfo.GuildName <> castillo3 Then GoTo npp
'If (bb = "C4" Or bb = "C8") And UserList(LoopC).GuildInfo.GuildName <> castillo4 Then GoTo npp
           
               
               ' If UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then _
                    BytesEnviados = BytesEnviados + Len(sndData)
   'pluto:2.10
'If UserList(LoopC).name = "AoDraGBoT" Then GoTo nap17
                  
                    'pluto:2.5.0
'sndData = CodificaR(str$(UserList(LoopC).flags.ValCoDe), sndData, 1)
'nap17:
                   ' BytesEnviados = BytesEnviados + Len(sndData)
                    'Call EnviarDatosASlot(LoopC, sndData)
                    ''frmMain.Socket2(LoopC).Write sndData, Len(sndData)
            'sndData = aa
            'End If
            'End If
'npp:
        'Next LoopC
        'Exit Sub
        
        '------------------------------------------------

    
    Case ToPCArea
      
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID > -1 And UserList(MapData(sndMap, X, Y).UserIndex).flags.UserLogged = True Then
       'pluto:2.10
If UserList(MapData(sndMap, X, Y).UserIndex).Name = "AoDraGBoT" Or UserList(MapData(sndMap, X, Y).UserIndex).Name = "AoDraGBoT2" Then GoTo nap8
                      
                            'pluto:2.5.0
sndData = CodificaR(str$(UserList((MapData(sndMap, X, Y).UserIndex)).flags.ValCoDe), sndData, MapData(sndMap, X, Y).UserIndex, 1)
nap8:
                            BytesEnviados = BytesEnviados + Len(sndData)
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                       sndData = aa
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    
    'pluto:6.0A
    Case ToPUserAreaCercana
      
        For Y = UserList(sndIndex).Pos.Y - 2 To UserList(sndIndex).Pos.Y + 2
            For X = UserList(sndIndex).Pos.X - 2 To UserList(sndIndex).Pos.X + 2
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID > -1 And UserList(MapData(sndMap, X, Y).UserIndex).flags.UserLogged = True Then
       'pluto:2.10
'If UserList(MapData(sndMap, X, Y).UserIndex).Name = "AoDraGBoT" Or UserList(MapData(sndMap, X, Y).UserIndex).Name = "AoDraGBoT2" Then GoTo nap8
                      
                            'pluto:2.5.0
sndData = CodificaR(str$(UserList((MapData(sndMap, X, Y).UserIndex)).flags.ValCoDe), sndData, MapData(sndMap, X, Y).UserIndex, 1)
'nap8:
                            BytesEnviados = BytesEnviados + Len(sndData)
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                       sndData = aa
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    
    
    Case ToNPCArea
        For Y = Npclist(sndIndex).Pos.Y - MinYBorder + 1 To Npclist(sndIndex).Pos.Y + MinYBorder - 1
            For X = Npclist(sndIndex).Pos.X - MinXBorder + 1 To Npclist(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID > -1 And UserList(MapData(sndMap, X, Y).UserIndex).flags.UserLogged = True Then
  'pluto:2.10
If UserList(MapData(sndMap, X, Y).UserIndex).Name = "AoDraGBoT" Or UserList(MapData(sndMap, X, Y).UserIndex).Name = "AoDraGBoT2" Then GoTo nap9
                 
                  'pluto:2.5.0
sndData = CodificaR(str$(UserList(MapData(sndMap, X, Y).UserIndex).flags.ValCoDe), sndData, MapData(sndMap, X, Y).UserIndex, 1)
nap9:
                            
                            BytesEnviados = BytesEnviados + Len(sndData)
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                       sndData = aa
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    Case ToIndex
        If (sndIndex = 0) Or sndIndex > MaxUsers Then Exit Sub
        If UserList(sndIndex).ConnID > -1 Then
'pluto.2.5.0

If Asc(Mid$(sndData, 2, 1)) = 18 Then
GoTo nop
End If
 'pluto:2.10
If UserList(sndIndex).Name = "AoDraGBoT" Or UserList(sndIndex).Name = "AoDraGBoT2" Then GoTo nop

sndData = CodificaR(str$(UserList(sndIndex).flags.ValCoDe), sndData, sndIndex, 1)
nop:
'--------
BytesEnviados = BytesEnviados + Len(sndData)
             Call EnviarDatosASlot(sndIndex, sndData)
             'frmMain.Socket2(sndIndex).Write sndData, Len(sndData)
             sndData = aa
             Exit Sub
        End If
End Select
Exit Sub
fallo:
Call LogError("senddata " & Err.number & " D: " & Err.Description)

End Sub
Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean
On Error GoTo fallo
Dim X As Integer, Y As Integer
For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
        For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1
            If MapData(UserList(index).Pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        Next X
Next Y
EstaPCarea = False

Exit Function
fallo:
Call LogError("estapcarea " & Err.number & " D: " & Err.Description)

End Function

Function HayPCarea(Pos As WorldPos) As Boolean
On Error GoTo fallo
'pluto:6.0A
If Pos.Map = 139 Or Pos.Map = 48 Or Pos.Map = 110 Then
HayPCarea = False
Exit Function
End If

Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False

Exit Function
fallo:
Call LogError("haypcarea " & Err.number & " D: " & Err.Description)

End Function
Function HayAguaCerca(Pos As WorldPos) As Boolean
On Error GoTo fallo
Dim X As Integer, Y As Integer
For Y = Pos.Y - 1 To Pos.Y + 1
        For X = Pos.X - 1 To Pos.X + 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then


                If HayAgua(Pos.Map, X, Y) = True Then
                HayAguaCerca = True
                Exit Function
                End If
            End If
        Next X
Next Y
HayAguaCerca = False
Exit Function
fallo:
Call LogError("hayaguacerca " & Err.number & " D: " & Err.Description)

End Function
Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean
On Error GoTo fallo
Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.Map, X, Y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
Exit Function
fallo:
Call LogError("hayobjarea " & Err.number & " D: " & Err.Description)

End Function
'pluto:2.4.5
Sub TiempoOnline(ByVal Tindex As Integer, shtimer As Integer, UserIndex As Integer)
On Error GoTo fallo
Dim kk As Integer
kk = MinutosOnline - UserList(UserIndex).ShTime
Call SendData(ToIndex, Tindex, 0, "||Time Sh: " & shtimer & " // " & kk & "´" & FontTypeNames.FONTTYPE_TALK)

Exit Sub
fallo:
Call LogError("tiempoonline " & Err.number & " D: " & Err.Description)

End Sub

Sub CorregirSkills(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim k As Integer

For k = 1 To NUMSKILLS
  If UserList(UserIndex).Stats.UserSkills(k) > MAXSKILLPOINTS Then UserList(UserIndex).Stats.UserSkills(k) = MAXSKILLPOINTS
Next

For k = 1 To NUMATRIBUTOS
 If UserList(UserIndex).Stats.UserAtributos(k) > MAXATRIBUTOS Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "El personaje tiene atributos invalidos.")
    Exit Sub
 End If
Next k
 Exit Sub
fallo:
Call LogError("corregirskills " & Err.number & " D: " & Err.Description)

End Sub


Function ValidateChr(ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
'pluto:2.15
'UserList(UserIndex).Bebe = 1
If UserList(UserIndex).Bebe > 0 Then ValidateChr = True: Exit Function

ValidateChr = UserList(UserIndex).Char.Head <> 0 And _
UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)
Exit Function
fallo:
Call LogError("validatechr " & Err.number & " D: " & Err.Description)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, Name As String, Password As String, Serie As String, Macplu As String)
On Error GoTo fallo
Dim n As Integer
Dim ooo As Byte
'pluto:6.5
'DoEvents
'pluto:6.7

'If Cuentas(UserIndex).mail <> UserList(UserIndex).EmailActual Then Exit Sub
If UserIndex < 1 Or Name = "" Then Exit Sub

If FileExist(App.Path & "\Ban-IP\" & UserList(UserIndex).ip & ".ips", vbArchive) Then
'If BDDIsBanIP(UserList(UserIndex).ip) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "La IP que usas está baneada en Aodrag.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'Reseteamos los FLAGS
'pluto:5.2
UserList(UserIndex).flags.CMuerte = 1
'---------
UserList(UserIndex).flags.Escondido = 0
UserList(UserIndex).flags.Protec = 0
UserList(UserIndex).flags.Ron = 0
UserList(UserIndex).flags.TargetNpc = 0
UserList(UserIndex).flags.TargetNpcTipo = 0
UserList(UserIndex).flags.TargetObj = 0
UserList(UserIndex).flags.TargetUser = 0
UserList(UserIndex).Char.FX = 0
'pluto:2.4.5
UserList(UserIndex).ShTime = 0
'pluto:2.9.0
UserList(UserIndex).ObjetosTirados = 0
UserList(UserIndex).Alarma = 0
UserList(UserIndex).flags.Macreanda = 0
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).Chetoso = 0
UserList(UserIndex).flags.ParejaTorneo = 0
'pluto:2.10
UserList(UserIndex).GranPoder = 0
UserList(UserIndex).Char.FX = 0
'pluto:2.19
ooo = 1
'¿Este IP ya esta conectado?
If AllowMultiLogins = 0 Then
    If CheckForSameIP(UserIndex, UserList(UserIndex).ip) = True Then
        Call SendData2(ToIndex, UserIndex, 0, 43, "No es posible usar mas de un personaje al mismo tiempo.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

'¿Ya esta conectado el personaje?
If CheckForSameName(UserIndex, Name) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Perdon, un usuario con el mismo nombre se há logoeado.")
 Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Existe el personaje?
If Not PersonajeExiste(Name) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "El personaje no existe..")
   Call CloseSocket(UserIndex)
    Exit Sub
End If
' ban ip

'pluto:2.9.0
'quitar esto
If Not EsDios(Name) And Not EsSemiDios(Name) And SoloGm = True Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "El server en estos momentos está abierto sólo para Gms, estamos comprobando que todo funcione correctamente.")
   Call CloseSocket(UserIndex)
    Exit Sub
End If
'pluto:2.19
ooo = 2

Dim Filex As String
Filex = CharPath & Left$(UCase$(Name), 1) & "\" & UCase$(Name) & ".chr"

If Not FileExist(CharPath & Left$(UCase$(Name), 1), vbDirectory) Then
    Call MkDir(CharPath & Left$(UCase$(Name), 1))
End If

If val(GetVar(Filex, "FLAGS", "BAN")) = 1 Then
'    Call SendData2(ToIndex, UserIndex, 0, 43, "Este personaje esta baneado")
'   Call CloseSocket(UserIndex)
'    Exit Sub
'End If
    'Delzak) ban
'Call LoadUserInit(UserIndex, filex, Name)
Dim rea As String
Dim rea2 As String
Dim rea3 As String
Dim rea4 As Boolean
rea4 = False
rea = GetVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason")
rea2 = GetVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Fecha")
'rea3 = Left$(Date, 2) - Day(Date)
'If rea3 < 0 Then rea3 = rea3 * -1
If UCase(Left$(rea, 6)) = "SEMANA" Then rea4 = True
'If rea = "SEMANA" Then rea4 = True
'rea3 = DateAdd("d", 7, rea2)
If DateDiff("d", rea2, Date) < 7 Then rea4 = False
'rea2 = DateDiff("d", rea2, Date)
'If rea3 > 7 Then rea4 = False
If rea4 = True Then
    'Call SendData2(ToIndex, UserIndex, 0, 43, "Este personaje fue baneado el dia " & rea2 & " debido a " & rea & ". El personaje ha sido desbaneado")
    Call SendData2(ToIndex, UserIndex, 0, 107, Name & "fue baneado el dia " & rea2 & " debido a " & rea & ". El personaje ha sido desbaneado")
    Call UnBan(Name)
    Call CloseSocket(UserIndex)
    'Call CloseUser(UserIndex) 'Lo echo
    Exit Sub
'    Call LogGM(UserList(UserIndex).AoDragBot, "/UNBAN a " & rdata)
Else
    Call SendData2(ToIndex, UserIndex, 0, 107, Name & " está baneado debido a  " & rea & " desde el dia " & rea2 & "")
    Call CloseSocket(UserIndex)
    Exit Sub
End If
End If ' fin baneado

'Cargamos los datos del personaje
Call LoadUserInit(UserIndex, Filex, Name)


'[Tite]Party
Call sendMiembrosParty(UserIndex)
'[\Tite]
'Call LoadUserStats(UserIndex, filex)

'pluto:2.3
'Call LoadUserMontura(UserIndex, filex)
'Call CorregirSkills(UserIndex)



'pluto:2.19
ooo = 3
'If UCase$(UserList(UserIndex).raza) = "ORCO" Then UserList(UserIndex).UserDañoArmasRaza = 20

'If UCase$(UserList(UserIndex).raza) = "HUMANO" Then
'UserList(UserIndex).UserDañoArmasRaza = 10
'UserList(UserIndex).UserDefensaMagiasRaza = 5
'End If
'pluto:6.0A camio enano +8 y +8
'If UCase$(UserList(UserIndex).raza) = "ENANO" Then
'UserList(UserIndex).UserDañoArmasRaza = 8
'UserList(UserIndex).UserDefensaMagiasRaza = 8
'UserList(UserIndex).UserEvasiónRaza = 10
'End If

'If UCase$(UserList(UserIndex).raza) = "GNOMO" Then
'UserList(UserIndex).UserDefensaMagiasRaza = 15
'UserList(UserIndex).UserEvasiónRaza = 10
'End If

'If UCase$(UserList(UserIndex).raza) = "VAMPIRO" Then
'UserList(UserIndex).UserEvasiónRaza = 10
'End If

If UCase$(UserList(UserIndex).raza) = "ELFO OSCURO" Then
'pluto:6.0A cambiamos a 10 el 15
'pluto:7.0 bonus invisibilidad elfo oscuro
'UserList(UserIndex).UserDañoProyetilesRaza = 10
'UserList(UserIndex).UserDefensaMagiasRaza = 5
UserList(UserIndex).BonusElfoOscuro = Porcentaje(IntervaloInvisible, 33)
End If

'If UCase$(UserList(UserIndex).raza) = "ELFO" Then
'UserList(UserIndex).UserDañoMagiasRaza = 8
'UserList(UserIndex).UserDefensaMagiasRaza = 10
'End If
'----------------------------------------------------------

If Not ValidateChr(UserIndex) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Error en el personaje.")
Call CloseSocket(UserIndex)
    Exit Sub
End If
'pluto:2.19
ooo = 4
'Call LoadUserReputacion(UserIndex, filex)
'pluto:2.14
UserList(UserIndex).Serie = Serie
'pluto:6.7
If UserList(UserIndex).MacPluto2 <> Macplu Then
'protec
'a:
'GoTo a
End If

If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).Char.ShieldAnim = NingunEscudo
If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).Char.CascoAnim = NingunCasco
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).Char.WeaponAnim = NingunArma
'[GAU]
If UserList(UserIndex).Invent.BotaEqpSlot = 0 Then UserList(UserIndex).Char.Botas = NingunBota
'[GAU]

'pluto:2.3 calcula peso
Dim X, x1 As Integer
UserList(UserIndex).Stats.Peso = 0
UserList(UserIndex).Stats.PesoMax = 0
For n = 1 To MAX_INVENTORY_SLOTS
X = UserList(UserIndex).Invent.Object(n).ObjIndex
x1 = UserList(UserIndex).Invent.Object(n).Amount
If X > 0 Then
UserList(UserIndex).Stats.Peso = UserList(UserIndex).Stats.Peso + (ObjData(X).Peso * x1)
End If
Next n
UserList(UserIndex).Stats.PesoMax = (UserList(UserIndex).Stats.UserAtributos(1) * 5) + (UserList(UserIndex).Stats.ELV * 3)
'pluto:4.2.1
If UserList(UserIndex).flags.Montura = 1 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + (UserList(UserIndex).flags.ClaseMontura * 100)
End If
If UserList(UserIndex).Invent.AnilloEqpObjIndex = 989 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + 500
End If
'pluto:6.0A------------
    If UserList(UserIndex).flags.Navegando = 1 Then
If UserList(UserIndex).Invent.BarcoObjIndex = 474 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + 100
ElseIf UserList(UserIndex).Invent.BarcoObjIndex = 475 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + 300
ElseIf UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + 500
End If
        End If 'navegando
'-----------------------
'pluto:2.9.0
Call SendUserClase(UserIndex)
'quitar esto
'EventoDia = 3
'pluto:6.8-----------
Select Case EventoDia

Case 1
Call SendData2(ToIndex, UserIndex, 0, 99, NombreBichoDelDia)
Case 2
Call SendData2(ToIndex, UserIndex, 0, 101)
Case 3
Call SendData2(ToIndex, UserIndex, 0, 102)
Case 4
Call SendData2(ToIndex, UserIndex, 0, 103, NombreBichoDelDia)
Case 5
Call SendData2(ToIndex, UserIndex, 0, 104)
End Select
'------------------

Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserHechizos(True, UserIndex, 0)
'pluto:2.19
ooo = 5
If UserList(UserIndex).flags.Navegando = 1 Then
     'pluto:6.0A---------
     If UserList(UserIndex).flags.Muerto = 0 Then
     UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
     Else
    UserList(UserIndex).Char.Body = 87
    End If
     '-------------------
     UserList(UserIndex).Char.Head = 0
     UserList(UserIndex).Char.WeaponAnim = NingunArma
     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
     UserList(UserIndex).Char.CascoAnim = NingunCasco
'[GAU]
     UserList(UserIndex).Char.Botas = NingunBota
     '[GAU]
End If

UserList(UserIndex).flags.Morph = 0
UserList(UserIndex).flags.Angel = 0
UserList(UserIndex).flags.Demonio = 0
'pluto:2.9.0
If UserList(UserIndex).flags.Paralizado Then
Call SendData2(ToIndex, UserIndex, 0, 68)
UserList(UserIndex).Counters.Paralisis = IntervaloParalisisPJ
End If

'Posicion de comienzo

'saca de mapa de torneo
'Dim x As Integer
Dim Y As Integer
Dim Map As Integer
'pluto:2.9.0 añade el 192 futbol
'pluto:2.12 añade torneo2
'If UserList(UserIndex).Pos.Map = MAPATORNEO Or UserList(UserIndex).Pos.Map = MapaTorneo2 Or UserList(UserIndex).Pos.Map = 192 Or UserList(UserIndex).Pos.Map = 191 Then
'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
'End If
'pluto:6.0A fabrica lingotes
'If UserList(UserIndex).Pos.Map = 277 And UserList(UserIndex).Pos.x = 36 And UserList(UserIndex).Pos.Y = 70 Then
' UserList(UserIndex).Pos = Nix
'End If


'pluto:2.18
'If (UserList(UserIndex).Pos.Map = 186 And fortaleza <> UserList(UserIndex).GuildInfo.GuildName) Then
'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
'End If
'pluto:6.0A
'If UserList(UserIndex).Pos.Map = 166 Or UserList(UserIndex).Pos.Map = 167 Or UserList(UserIndex).Pos.Map = 168 Or UserList(UserIndex).Pos.Map = 169 Then
'    UserList(UserIndex).Pos.Map = UserList(UserIndex).Pos.Map
'    UserList(UserIndex).Pos.x = 26 + RandomNumber(1, 9)
'    UserList(UserIndex).Pos.Y = 85 + RandomNumber(1, 5)
'End If




'pluto:2.15
'Dim a As Integer
'Dim b As Byte
'
'If Criminal(UserIndex) Then b = 2 Else b = 1


'a = ReadField(1, GetVar(filex, "INIT", "Position"), 45)
'If a = 0 Then GoTo ff:
'If MapInfo(a).Dueño = 2 And b = 1 And UserList(UserIndex).flags.Muerto = 0 Then
'Call SendData2(ToIndex, UserIndex, 0, 43, "La ciudad en la que te encontrabas ha sido conquistada por las Fuerzas Imperiales y has tenido que huir a una ciudad segura.")
'Call SendData(ToIndex, UserIndex, 0, "||La ciudad en la que te encontrabas ha sido conquistada por las Fuerzas Imperiales y has tenido que huir a una ciudad segura." & FONTTYPENAMES.FONTTYPE_COMERCIO)
'UserList(UserIndex).Pos = Banderbill

'End If

'pluto:2.19
ooo = 6
'If MapInfo(a).Dueño = 1 And b = 2 And UserList(UserIndex).flags.Muerto = 0 Then
'Call SendData2(ToIndex, UserIndex, 0, 43, "La ciudad en la que te encontrabas ha sido conquistada por las Fuerzas Del Caos y has tenido que huir a una ciudad segura.")
'Call SendData(ToIndex, UserIndex, 0, "||La ciudad en la que te encontrabas ha sido conquistada por las Fuerzas Del Caos y has tenido que huir a una ciudad segura." & FONTTYPENAMES.FONTTYPE_COMERCIO)

'UserList(UserIndex).Pos = ciudadcaos

'End If




'ff:


'pluto:6.5 -----------------------------------------------------------------
Select Case UserList(UserIndex).Pos.Map
Case MAPATORNEO 'torneos
'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
UserList(UserIndex).Pos.Map = 34
UserList(UserIndex).Pos.X = 35
UserList(UserIndex).Pos.Y = 35
Case MapaTorneo2 'torneos
'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
UserList(UserIndex).Pos.Map = 34
UserList(UserIndex).Pos.X = 35
UserList(UserIndex).Pos.Y = 35
Case 303 'torneos gms
UserList(UserIndex).Pos.Map = 34
UserList(UserIndex).Pos.X = 35
UserList(UserIndex).Pos.Y = 35
Case 291 To 295 'torneos
'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
UserList(UserIndex).Pos.Map = 34
UserList(UserIndex).Pos.X = 35
UserList(UserIndex).Pos.Y = 35
Case 277 'fabrica lingotes
If UserList(UserIndex).Pos.X = 36 And UserList(UserIndex).Pos.Y = 70 Then UserList(UserIndex).Pos = Nix

Case 186 'fortaleza
If fortaleza <> UserList(UserIndex).GuildInfo.GuildName Then
If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
End If

Case 166 To 169 'castillos
UserList(UserIndex).Pos.X = 26 + RandomNumber(1, 9)
UserList(UserIndex).Pos.Y = 85 + RandomNumber(1, 5)

Case 191 To 192 'dragfutbol
UserList(UserIndex).Pos = Nix

End Select
'------------------------------------------------------------------------



'pluto:6.0A-------
If FileExist(App.Path & "\Bloqueos\" & UserList(UserIndex).Serie & ".lol", vbArchive) Then
'Call WarpUserChar(UserIndex, 191, 50, 50, True)
    UserList(UserIndex).Pos.Map = 191
    UserList(UserIndex).Pos.X = 50
    UserList(UserIndex).Pos.Y = 50
Call SendData(ToIndex, UserIndex, 0, "I2")
'Call SendData(ToIndex, UserIndex, 0, "|| Está Pc ha sido bloqueada para jugar Aodrag, aparecerás en este Mapa cada vez que juegues, avisa Gm para desbloquear la Pc y portate bién o atente a las consecuencias." & FONTTYPENAMES.FONTTYPE_TALK)
'pluto:2.11
Call SendData(ToAdmins, UserIndex, 0, "|| Ha entrado en Mapa 191: " & Name & "´" & FontTypeNames.FONTTYPE_TALK)
Call LogMapa191("Jugador:" & Name & " entró al Mapa 191 " & "Ip: " & UserList(UserIndex).ip)
End If
'-------------------
'pluto:6.0A-------
If FileExist(App.Path & "\MacPluto\" & UserList(UserIndex).MacPluto & ".lol", vbArchive) Then
Call SendData(ToIndex, UserIndex, 0, "W7")
Call SendData(ToAdmins, UserIndex, 0, "|| Cliente Colgado: " & Name & "´" & FontTypeNames.FONTTYPE_TALK)
'Call LogMapa191("Jugador:" & Name & " entró al Mapa 191 " & "Ip: " & UserList(UserIndex).ip)
Call CloseSocket(UserIndex)
    Exit Sub
    End If
'-------------------





If UserList(UserIndex).Pos.Map = 0 Then
'pluto:6.0A

If UserList(UserIndex).clase = "Mago" Or UserList(UserIndex).clase = "Druida" Then
Call SendData(ToIndex, UserIndex, 0, "AWmagico1")
Else
Call SendData(ToIndex, UserIndex, 0, "AWcurro1")
End If

'pluto:2.17---------------
    If UCase$(UserList(UserIndex).Hogar) = "ALDEA DE HUMANOS" Then
             UserList(UserIndex).Pos = Pobladohumano
    ElseIf UCase$(UserList(UserIndex).Hogar) = "POBLADO ORCO" Then
            UserList(UserIndex).Pos = Pobladoorco
    ElseIf UCase$(UserList(UserIndex).Hogar) = "POBLADO ENANO" Then
             UserList(UserIndex).Pos = Pobladoenano
    ElseIf UCase$(UserList(UserIndex).Hogar) = "ALDEA DE GNOMOS" Then
             UserList(UserIndex).Pos = Pobladoenano
      ElseIf UCase$(UserList(UserIndex).Hogar) = "ALDEA ÉLFICA" Then
             UserList(UserIndex).Pos = Pobladoelfo
    Else
       UserList(UserIndex).Hogar = "ALDEA DE VAMPIROS"
        UserList(UserIndex).Pos = Pobladovampiro
    End If
'-------------------------------------
Else
   'pluto:6.5
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex <> 0 Then
'GetObj (MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex)
Call CloseUser(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex)
End If

End If

'Nombre de sistema
UserList(UserIndex).Name = Name
'UserList(UserIndex).ip = frmMain.Socket2(UserIndex).PeerAddress
  
'Info
Call SendData(ToIndex, UserIndex, 0, "IU" & UserIndex) 'Enviamos el User index
Call SendData2(ToIndex, UserIndex, 0, 14, UserList(UserIndex).Pos.Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion) 'Carga el mapa
Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(UserList(UserIndex).Pos.Map).Music)

If Lloviendo Then Call SendData2(ToIndex, UserIndex, 0, 20, "1")

'pluto:2.15
Call SendUserMuertos(UserIndex)
'pluto:6.0A
Call SendUserStatsFama(UserIndex)
 'envia dueño mapas
 'Dim n As Integer
Dim ci As String

'[Tite]Nix neutral
ci = str(MapInfo(1).Dueño) & "," & str(MapInfo(20).Dueño) & "," & str(MapInfo(63).Dueño) & "," & str(MapInfo(81).Dueño) & "," & str(MapInfo(84).Dueño) & "," & str(MapInfo(112).Dueño) & "," & str(MapInfo(151).Dueño) & "," & str(MapInfo(157).Dueño) & "," & str(MapInfo(184).Dueño)
'ci = str(MapInfo(1).Dueño) & "," & str(MapInfo(20).Dueño) & "," & str(MapInfo(34).Dueño) & "," & str(MapInfo(63).Dueño) & "," & str(MapInfo(81).Dueño) & "," & str(MapInfo(84).Dueño) & "," & str(MapInfo(112).Dueño) & "," & str(MapInfo(151).Dueño) & "," & str(MapInfo(157).Dueño) & "," & str(MapInfo(184).Dueño)
'[\Tite]

Call SendData(ToIndex, UserIndex, 0, "K4" & ci)
'------------------------------------------------------

If AtaNorte = 1 Then Call SendData(ToIndex, UserIndex, 0, "C1")
If AtaSur = 1 Then Call SendData(ToIndex, UserIndex, 0, "C2")
If AtaEste = 1 Then Call SendData(ToIndex, UserIndex, 0, "C3")
If AtaOeste = 1 Then Call SendData(ToIndex, UserIndex, 0, "C4")
If AtaForta = 1 Then Call SendData(ToIndex, UserIndex, 0, "V8")
'pluto:2.19
ooo = 7
Call UpdateUserMap(UserIndex)
Call senduserstatsbox(UserIndex)
Call SendUserRazaClase(UserIndex)
'Call SendUserPremios(UserIndex) 'Delzak premios
'pluto:2.3
Call SendUserStatsPeso(UserIndex)

Call EnviarHambreYsed(UserIndex)

Call SendMOTD(UserIndex)

If haciendoBK Or haciendoBKPJ Then
    Call SendData2(ToIndex, UserIndex, 0, 19)
    Call SendData(ToIndex, UserIndex, 0, "||Por favor espera algunos segundo, WorldSave esta ejecutandose." & "´" & FontTypeNames.FONTTYPE_info)
End If

'Actualiza el Num de usuarios
If UserIndex > LastUser Then LastUser = UserIndex

NumUsers = NumUsers + 1
'pluto.2.8.0
If NumUsers >= ReNumUsers Then ReNumUsers = NumUsers: HoraHoy = Time

'pluto:2.4
'If Not Criminal(UserIndex) Then UserCiu = UserCiu + 1 Else UserCrimi = UserCrimi + 1


'Call SendData2(ToAll, UserIndex, 0, 17, CStr(NumUsers))
'Call SendData2(ToIndex, UserIndex, 0, 17, CStr(NumUsers))
'pluto:6.8
If UserList(UserIndex).flags.Privilegios = 0 Then MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers + 1

'If UserList(UserIndex).Stats.SkillPts > 0 Then
    Call EnviarSkills(UserIndex)
    'Call EnviarSubirNivel(UserIndex, UserList(UserIndex).Stats.SkillPts)
'End If

'If NumUsers > DayStats.Maxusuarios Then DayStats.Maxusuarios = NumUsers

If NumUsers > recordusuarios Then
    Call SendData(ToAll, 0, 0, "||Record de usuarios conectados simultaneamente." & "Hay " & Round(NumUsers) & " usuarios." & "´" & FontTypeNames.FONTTYPE_info)
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
End If
'pluto:2.11 añade UserList(UserIndex).Faccion.ArmadaReal = 2
If (UserList(UserIndex).Faccion.RecompensasCaos > 9 And UserList(UserIndex).Faccion.CiudadanosMatados < 800) Or (UserList(UserIndex).Faccion.RecompensasReal > 10 And UserList(UserIndex).Faccion.CriminalesMatados < 800 And UserList(UserIndex).Faccion.ArmadaReal = 1) Then
If UserList(UserIndex).flags.Privilegios = 0 Then Call LogCasino("Jugador:" & UserList(UserIndex).Name & " mirar recompensas armadas " & "Ip: " & UserList(UserIndex).ip)
End If

If EsDios(Name) Then
    UserList(UserIndex).flags.Privilegios = 3
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip & " SE: " & UserList(UserIndex).Serie)

ElseIf EsSemiDios(Name) Then
    UserList(UserIndex).flags.Privilegios = 2
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip & " SE: " & UserList(UserIndex).Serie)
ElseIf EsConsejero(Name) Then
    UserList(UserIndex).flags.Privilegios = 1
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip & " SE: " & UserList(UserIndex).Serie)
Else
    UserList(UserIndex).flags.Privilegios = 0
End If
'pluto:2.19
ooo = 8
'If UserList(UserIndex).Flags.Privilegios > 0 Then Call BDDSetGMState(UCase$(name), 1)
Set UserList(UserIndex).GuildRef = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

UserList(UserIndex).Counters.IdleCount = 0

If UserList(UserIndex).NroMacotas > 0 Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) > 0 Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(UserList(UserIndex).MascotasType(i), UserList(UserIndex).Pos, True, True)
            
            If UserList(UserIndex).MascotasIndex(i) <= MAXNPCS Then
                  Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
                  Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
            Else
                  UserList(UserIndex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If


If UserList(UserIndex).flags.Navegando = 1 Then Call SendData2(ToIndex, UserIndex, 0, 6)


UserList(UserIndex).flags.Seguro = True
'[Tite]Seguro de ataques criticos
UserList(UserIndex).flags.SegCritico = False
'[/Tite]
UserList(UserIndex).flags.UserLogged = True
'Crea  el personaje del usuario
Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
Call SendData2(ToIndex, UserIndex, 0, 4)


Call SendGuildNews(UserIndex)
Call MostrarNumUsers
'Call BDDSetUsersOnline


UserList(UserIndex).flags.ClienteValido = 0    ' VERIFICACION CLIENTES
Call SendData(ToIndex, UserIndex, 0, "F9")

'pluto:2.14
Call ComprobarLista(UserList(UserIndex).Name)
'-----------------------
'n = FreeFile
'Open App.Path & "\logs\numusers.log" For Output As n
'Print #n, NumUsers
'Close #n

'n = FreeFile
'Log
'Open App.Path & "\logs\Connect.log" For Append Shared As #n
'Print #n, UserList(UserIndex).Name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
'Close #n

'pluto:2.9.0
UserList(UserIndex).ShTime = 0
UserList(UserIndex).ObjetosTirados = 0
UserList(UserIndex).Alarma = 0
'pluto:2.10
UserList(UserIndex).GranPoder = 0
'pluto:2.5.0
If UserList(UserIndex).GuildInfo.GuildName = "" Then UserList(UserIndex).GuildInfo.GuildPoints = 0
'pluto:2.9.0
If MsgEntra <> "" Then Call SendData2(ToIndex, UserIndex, 0, 43, MsgEntra)
Call SendData2(ToAll, 0, 0, 117, "B" & DobleExp)
'pluto:2.19
ooo = 9
'pluto:2.17
'If a = 0 Then GoTo fff
'If MapInfo(a).Dueño = 2 And b = 1 And UserList(UserIndex).flags.Muerto = 0 Then Call SendData(ToIndex, UserIndex, 0, "||La ciudad en la que te encontrabas ha sido conquistada por las Fuerzas del Caos y has tenido que huir a una Ciudad Segura." & FONTTYPENAMES.FONTTYPE_COMERCIO)
'If MapInfo(a).Dueño = 1 And b = 2 And UserList(UserIndex).flags.Muerto = 0 Then Call SendData(ToIndex, UserIndex, 0, "||La ciudad en la que te encontrabas ha sido conquistada por las Fuerzas del Imperio Real y has tenido que huir a una Ciudad Segura." & FONTTYPENAMES.FONTTYPE_COMERCIO)

fff:



Exit Sub
fallo:
Call LogError("connectuser->Nombre: " & Name & " Ip: " & UserList(UserIndex).ip & " seña: " & ooo & " " & Err.Description)

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
On Error GoTo fallo
'Dim j As Integer
'Call SendData(ToIndex, UserIndex, 0, "||Npc Del Día: " & NombreBichoDelDia & "´" & FontTypeNames.FONTTYPE_talk)

'For j = 1 To MaxLines
 '   Call SendData(ToIndex, UserIndex, 0, "||" & MOTD(j) & "´" & FontTypeNames.FONTTYPE_INFO)
'Next j
'Call SendData(ToIndex, UserIndex, 0, "||Castillo Norte: " & castillo1 & " Fecha: " & date1 & " Hora: " & hora1 & "´" & FontTypeNames.FONTTYPE_INFO)
'Call SendData(ToIndex, UserIndex, 0, "||Castillo Sur: " & castillo2 & " Fecha: " & date2 & " Hora: " & hora2 & "´" & FontTypeNames.FONTTYPE_INFO)
'Call SendData(ToIndex, UserIndex, 0, "||Castillo Este: " & castillo3 & " Fecha: " & date3 & " Hora: " & hora3 & "´" & FontTypeNames.FONTTYPE_INFO)
'Call SendData(ToIndex, UserIndex, 0, "||Castillo Oeste: " & castillo4 & " Fecha: " & date4 & " Hora: " & hora4 & "´" & FontTypeNames.FONTTYPE_INFO)
'Call SendData(ToIndex, UserIndex, 0, "||Fortaleza: " & fortaleza & " Fecha: " & date5 & " Hora: " & hora5 & "´" & FontTypeNames.FONTTYPE_INFO)
Exit Sub
fallo:
Call LogError("sendmotd " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).Faccion.ArmadaReal = 0
UserList(UserIndex).Faccion.FuerzasCaos = 0
UserList(UserIndex).Faccion.CiudadanosMatados = 0
UserList(UserIndex).Faccion.CriminalesMatados = 0
UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0
UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
'pluto:2.3
UserList(UserIndex).Faccion.RecibioArmaduraLegion = 0
UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0
UserList(UserIndex).Faccion.RecibioExpInicialReal = 0
UserList(UserIndex).Faccion.RecompensasCaos = 0
UserList(UserIndex).Faccion.RecompensasReal = 0
Exit Sub
fallo:
Call LogError("resetfacciones " & Err.number & " D: " & Err.Description)

End Sub
'pluto:2.4
Sub ResetTodasMonturas(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim xx As Integer
For xx = 1 To MAXMONTURA
UserList(UserIndex).Montura.Nivel(xx) = 0
UserList(UserIndex).Montura.exp(xx) = 0
UserList(UserIndex).Montura.Elu(xx) = 0
UserList(UserIndex).Montura.Vida(xx) = 0
UserList(UserIndex).Montura.Golpe(xx) = 0
UserList(UserIndex).Montura.Nombre(xx) = ""
UserList(UserIndex).Montura.AtCuerpo(xx) = 0
UserList(UserIndex).Montura.Defcuerpo(xx) = 0
UserList(UserIndex).Montura.AtFlechas(xx) = 0
UserList(UserIndex).Montura.DefFlechas(xx) = 0
UserList(UserIndex).Montura.AtMagico(xx) = 0
UserList(UserIndex).Montura.DefMagico(xx) = 0
UserList(UserIndex).Montura.Evasion(xx) = 0
UserList(UserIndex).Montura.Tipo(xx) = 0
UserList(UserIndex).Montura.index(xx) = 0
UserList(UserIndex).Montura.Libres(xx) = 0
Next
Exit Sub
fallo:
Call LogError("resettodomonturas " & Err.number & " D: " & Err.Description)

End Sub
Sub ResetContadores(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).Counters.AGUACounter = 0
UserList(UserIndex).Counters.AttackCounter = 0
UserList(UserIndex).Counters.Ceguera = 0
UserList(UserIndex).Counters.COMCounter = 0
UserList(UserIndex).Counters.Estupidez = 0
UserList(UserIndex).Counters.Frio = 0
UserList(UserIndex).Counters.HPCounter = 0
UserList(UserIndex).Counters.IdleCount = 0
UserList(UserIndex).Counters.Invisibilidad = 0
UserList(UserIndex).Counters.Paralisis = 0
UserList(UserIndex).Counters.Morph = 0
UserList(UserIndex).Counters.Angel = 0
UserList(UserIndex).Counters.Pasos = 0
UserList(UserIndex).Counters.Pena = 0
UserList(UserIndex).Counters.PiqueteC = 0
UserList(UserIndex).Counters.bloqueo = 0
UserList(UserIndex).Counters.STACounter = 0
UserList(UserIndex).Counters.veneno = 0
Exit Sub
fallo:
Call LogError("resetcontadores " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
On Error GoTo fallo
'[GAU]
UserList(UserIndex).Char.Botas = 0
'[GAU]
UserList(UserIndex).Char.Body = 0
UserList(UserIndex).Char.CascoAnim = 0
UserList(UserIndex).Char.CharIndex = 0
UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.Head = 0
UserList(UserIndex).Char.loops = 0
UserList(UserIndex).Char.Heading = 0
UserList(UserIndex).Char.loops = 0
UserList(UserIndex).Char.ShieldAnim = 0
UserList(UserIndex).Char.WeaponAnim = 0
Exit Sub
fallo:
Call LogError("resetcharinfo " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).Name = ""
UserList(UserIndex).modName = ""
UserList(UserIndex).Password = ""
UserList(UserIndex).Desc = ""
UserList(UserIndex).Pos.Map = 0
UserList(UserIndex).Pos.X = 0
UserList(UserIndex).Pos.Y = 0
'UserList(UserIndex).ip = ""
'UserList(UserIndex).RDBuffer = ""
UserList(UserIndex).clase = ""
'pluto:2.14
UserList(UserIndex).Serie = ""
'UserList(UserIndex).MacPluto = ""
'UserList(UserIndex).MacPluto2 = ""
'UserList(UserIndex).MacClave = 0
UserList(UserIndex).Nhijos = 0
UserList(UserIndex).Padre = ""
UserList(UserIndex).Madre = ""
Dim X As Byte
For X = 1 To 5
UserList(UserIndex).Hijo(X) = ""
Next X
UserList(UserIndex).Esposa = ""
UserList(UserIndex).Paquete = 0
UserList(UserIndex).Amor = 0
UserList(UserIndex).Embarazada = 0
UserList(UserIndex).Bebe = 0
UserList(UserIndex).NombreDelBebe = ""

'pluto:2.10
UserList(UserIndex).EmailActual = ""
UserList(UserIndex).Email = ""
UserList(UserIndex).Genero = ""
UserList(UserIndex).Hogar = ""
UserList(UserIndex).raza = ""

UserList(UserIndex).RandKey = 0
UserList(UserIndex).PrevCRC = 0
UserList(UserIndex).PacketNumber = 0

UserList(UserIndex).Stats.Banco = 0
UserList(UserIndex).Stats.ELV = 0
UserList(UserIndex).Stats.Elu = 0
UserList(UserIndex).Stats.LibrosUsados = 0
UserList(UserIndex).Stats.Fama = 0
UserList(UserIndex).Stats.exp = 0
UserList(UserIndex).Stats.Def = 0
UserList(UserIndex).Stats.CriminalesMatados = 0
UserList(UserIndex).Stats.NPCsMuertos = 0
UserList(UserIndex).Stats.UsuariosMatados = 0
Exit Sub
fallo:
Call LogError("resetbasicuserinfo " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.BurguesRep = 0
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.NobleRep = 0
UserList(UserIndex).Reputacion.PlebeRep = 0
UserList(UserIndex).Reputacion.NobleRep = 0
UserList(UserIndex).Reputacion.Promedio = 0
Exit Sub
fallo:
Call LogError("resetreputacion " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).GuildInfo.ClanFundado = ""
UserList(UserIndex).GuildInfo.Echadas = 0
UserList(UserIndex).GuildInfo.EsGuildLeader = 0
UserList(UserIndex).GuildInfo.FundoClan = 0
UserList(UserIndex).GuildInfo.GuildName = ""
UserList(UserIndex).GuildInfo.Solicitudes = 0
UserList(UserIndex).GuildInfo.SolicitudesRechazadas = 0
UserList(UserIndex).GuildInfo.VecesFueGuildLeader = 0
UserList(UserIndex).GuildInfo.YaVoto = 0
UserList(UserIndex).GuildInfo.ClanesParticipo = 0
UserList(UserIndex).GuildInfo.GuildPoints = 0
Exit Sub
fallo:
Call LogError("resetguildinfo " & Err.number & " D: " & Err.Description)

End Sub
'pluto:7.0
Sub ResetUserMision(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).Mision.Cargada = False
UserList(UserIndex).Mision.estado = 0
UserList(UserIndex).Mision.numero = 0
UserList(UserIndex).Mision.TimeComienzo = ""
Erase UserList(UserIndex).Mision.Enemigo()
Erase UserList(UserIndex).Mision.EnemigoCantidad()
UserList(UserIndex).Mision.Entrega = 0
Erase UserList(UserIndex).Mision.Objeto()
Erase UserList(UserIndex).Mision.ObjetoR()
UserList(UserIndex).Mision.Entrega = 0
UserList(UserIndex).Mision.exp = 0
UserList(UserIndex).Mision.oro = 0
UserList(UserIndex).Mision.NEnemigos = 0
Erase UserList(UserIndex).Mision.NEnemigosConseguidos()
UserList(UserIndex).Mision.NivelMaximo = 0
UserList(UserIndex).Mision.NivelMinimo = 0
UserList(UserIndex).Mision.NObjetos = 0
UserList(UserIndex).Mision.NObjetosR = 0
UserList(UserIndex).Mision.NpcQuest = 0
UserList(UserIndex).Mision.PjConseguidos = 0
UserList(UserIndex).Mision.TimeMision = 0
UserList(UserIndex).Mision.Titulo = ""
UserList(UserIndex).Mision.tX = ""
UserList(UserIndex).Mision.Cantidad = 0
UserList(UserIndex).Mision.Level = 0
UserList(UserIndex).Mision.clase = ""
UserList(UserIndex).Mision.Actual = 0
UserList(UserIndex).Mision.Actual1 = 0
UserList(UserIndex).Mision.Actual2 = 0
UserList(UserIndex).Mision.Actual3 = 0
UserList(UserIndex).Mision.Actual4 = 0
UserList(UserIndex).Mision.Actual5 = 0
UserList(UserIndex).Mision.Actual6 = 0
UserList(UserIndex).Mision.Actual7 = 0
UserList(UserIndex).Mision.Actual8 = 0
UserList(UserIndex).Mision.Actual9 = 0
UserList(UserIndex).Mision.Actual10 = 0
UserList(UserIndex).Mision.Actual11 = 0
UserList(UserIndex).Mision.Actual12 = 0
Exit Sub
fallo:
Call LogError("resetusermision " & Err.number & " D: " & Err.Description)

End Sub
Sub ResetMisionCompletada(ByVal UserIndex As Integer)
On Error GoTo fallo
UserList(UserIndex).Mision.estado = 0
UserList(UserIndex).Mision.TimeComienzo = ""
Erase UserList(UserIndex).Mision.Enemigo()
Erase UserList(UserIndex).Mision.EnemigoCantidad()
UserList(UserIndex).Mision.Entrega = 0
Erase UserList(UserIndex).Mision.Objeto()
Erase UserList(UserIndex).Mision.ObjetoR()
UserList(UserIndex).Mision.Entrega = 0
UserList(UserIndex).Mision.exp = 0
UserList(UserIndex).Mision.oro = 0
UserList(UserIndex).Mision.NEnemigos = 0
Erase UserList(UserIndex).Mision.NEnemigosConseguidos()
UserList(UserIndex).Mision.NivelMaximo = 0
UserList(UserIndex).Mision.NivelMinimo = 0
UserList(UserIndex).Mision.NObjetos = 0
UserList(UserIndex).Mision.NObjetosR = 0
UserList(UserIndex).Mision.NpcQuest = 0
UserList(UserIndex).Mision.PjConseguidos = 0
UserList(UserIndex).Mision.TimeMision = 0
UserList(UserIndex).Mision.Titulo = ""
UserList(UserIndex).Mision.tX = ""
UserList(UserIndex).Mision.Cantidad = 0
UserList(UserIndex).Mision.Level = 0
UserList(UserIndex).Mision.clase = ""
UserList(UserIndex).Mision.Actual = 0
Exit Sub
fallo:
Call LogError("resetmisioncompletada " & Err.number & " D: " & Err.Description)

End Sub
Sub ResetUserFlags(ByVal UserIndex As Integer)
On Error GoTo fallo
'pluto:6.2
UserList(UserIndex).flags.Incor = 0
UserList(UserIndex).flags.MapaIncor = 0
'----------
'pluto:6.8
UserList(UserIndex).flags.NoTorneos = False
UserList(UserIndex).flags.Intentos = 0

UserList(UserIndex).flags.Comerciando = False
UserList(UserIndex).flags.ban = 0
'pluto:5.2
UserList(UserIndex).flags.CMuerte = 1
'--------
UserList(UserIndex).flags.Escondido = 0
UserList(UserIndex).flags.Protec = 0
UserList(UserIndex).flags.Ron = 0
UserList(UserIndex).flags.DuracionEfecto = 0
UserList(UserIndex).flags.NpcInv = 0
UserList(UserIndex).flags.StatsChanged = 0
UserList(UserIndex).flags.TargetNpc = 0
UserList(UserIndex).flags.TargetNpcTipo = 0
UserList(UserIndex).flags.TargetObj = 0
UserList(UserIndex).flags.TargetObjMap = 0
UserList(UserIndex).flags.TargetObjX = 0
UserList(UserIndex).flags.TargetObjY = 0
UserList(UserIndex).flags.TargetUser = 0
UserList(UserIndex).flags.TipoPocion = 0
UserList(UserIndex).flags.TomoPocion = False
UserList(UserIndex).flags.Descuento = ""
UserList(UserIndex).flags.Hambre = 0
UserList(UserIndex).flags.Sed = 0
UserList(UserIndex).flags.Descansar = False
UserList(UserIndex).flags.ModoCombate = False
'pluto:6.0A
UserList(UserIndex).flags.Pitag = 0
UserList(UserIndex).flags.Arqui = 0
'[Tite]Seguro de golpes criticos
UserList(UserIndex).flags.SegCritico = False
'[\Tite]
'[Tite]Flag Party
UserList(UserIndex).flags.party = False
'[/Tite]
UserList(UserIndex).flags.Vuela = 0
UserList(UserIndex).flags.Navegando = 0
'pluto:2.3
UserList(UserIndex).flags.Montura = 0
UserList(UserIndex).flags.ClaseMontura = 0

UserList(UserIndex).flags.Oculto = 0
UserList(UserIndex).flags.Envenenado = 0
UserList(UserIndex).flags.Morph = 0
UserList(UserIndex).flags.Invisible = 0
UserList(UserIndex).flags.Paralizado = 0
UserList(UserIndex).flags.Angel = 0
UserList(UserIndex).flags.Demonio = 0
UserList(UserIndex).flags.Maldicion = 0
UserList(UserIndex).flags.Bendicion = 0
UserList(UserIndex).flags.Meditando = 0
UserList(UserIndex).flags.Privilegios = 0

'pluto:6.0A-------------------
UserList(UserIndex).flags.Minotauro = 0
UserList(UserIndex).flags.MinutosOnline = 0
'pluto:7.0
UserList(UserIndex).flags.Creditos = 0

UserList(UserIndex).flags.DragCredito1 = 0
UserList(UserIndex).flags.DragCredito2 = 0
UserList(UserIndex).flags.DragCredito3 = 0
UserList(UserIndex).flags.DragCredito4 = 0
UserList(UserIndex).flags.DragCredito5 = 0
UserList(UserIndex).flags.DragCredito6 = 0
'pluto:7.0
UserList(UserIndex).flags.NCaja = 0
UserList(UserIndex).flags.Elixir = 0
'--------------------

UserList(UserIndex).flags.PuedeMoverse = 0
UserList(UserIndex).flags.PuedeLanzarSpell = 0
'pluto:2.23
'UserList(UserIndex).flags.PuedeFlechas = 0
'pluto:2.10
UserList(UserIndex).flags.PuedeTomar = 0

UserList(UserIndex).Stats.SkillPts = 0
UserList(UserIndex).flags.OldBody = 0
UserList(UserIndex).flags.OldHead = 0
UserList(UserIndex).flags.AdminInvisible = 0
'pluto:2.5.0
'UserList(userindex).Flags.ValCoDe = 0
UserList(UserIndex).flags.Hechizo = 0
'pluto:6.2
UserList(UserIndex).flags.Macreanda = 0
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.ParejaTorneo = 0
'pluto:6.7
UserList(UserIndex).flags.party = False
UserList(UserIndex).flags.partyNum = 0
UserList(UserIndex).flags.invitado = ""
UserList(UserIndex).flags.privado = 0
'--------------------
'pluto:6.8----------------------------------
UserList(UserIndex).PoSum.Map = 0
UserList(UserIndex).PoSum.X = 0
UserList(UserIndex).PoSum.Y = 0
'----------------------
Exit Sub
fallo:
Call LogError("resetuserflags " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim loopc As Integer
For loopc = 1 To MAXUSERHECHIZOS
    UserList(UserIndex).Stats.UserHechizos(loopc) = 0
Next
Exit Sub
fallo:
Call LogError("resetuserspells " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim loopc As Integer

UserList(UserIndex).NroMacotas = 0
    
For loopc = 1 To MAXMASCOTAS
    UserList(UserIndex).MascotasIndex(loopc) = 0
    UserList(UserIndex).MascotasType(loopc) = 0
Next loopc
Exit Sub
fallo:
Call LogError("resetuserpets " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
On Error GoTo fallo
Dim loopc As Integer
Dim n As Byte
    For n = 1 To 6
For loopc = 1 To MAX_BANCOINVENTORY_SLOTS

      UserList(UserIndex).BancoInvent(n).Object(loopc).Amount = 0
      UserList(UserIndex).BancoInvent(n).Object(loopc).Equipped = 0
      UserList(UserIndex).BancoInvent(n).Object(loopc).ObjIndex = 0
  
Next
  Next n
'UserList(UserIndex).BancoInvent.NroItems = 0

Exit Sub
fallo:
Call LogError("resetuserbanco " & Err.number & " D: " & Err.Description)

End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
On Error GoTo fallo


Set UserList(UserIndex).CommandsBuffer = Nothing
Set UserList(UserIndex).GuildRef = Nothing



UserList(UserIndex).BD = 0
UserList(UserIndex).Remort = 0
UserList(UserIndex).Remorted = ""
'pluto:2.4
UserList(UserIndex).Stats.Puntos = 0
UserList(UserIndex).Stats.GTorneo = 0
UserList(UserIndex).Stats.PClan = 0
'pluto:2.4.5
UserList(UserIndex).ShTime = 0
'pluto:2.9.0
UserList(UserIndex).ObjetosTirados = 0
UserList(UserIndex).Alarma = 0
UserList(UserIndex).Chetoso = 0
'pluto:2.10
UserList(UserIndex).GranPoder = 0
UserList(UserIndex).Char.FX = 0
'pluto:2.12
UserList(UserIndex).flags.Torneo = 0
UserList(UserIndex).Torneo2 = 0
UserList(UserIndex).Stats.LibrosUsados = 0
UserList(UserIndex).Stats.Fama = 0
'pluto:6.0A
UserList(UserIndex).Nmonturas = 0

Call ResetTodasMonturas(UserIndex)

Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetReputacion(UserIndex)
Call ResetGuildInfo(UserIndex)
'pluto:hoy
Call ResetUserMision(UserIndex)

Call ResetUserFlags(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)
'Call ResetUserBanco(UserIndex)

Exit Sub
fallo:
Call LogError("resetuserslots " & Err.number & " D: " & Err.Description)

End Sub


Sub CloseUser(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim Tindex As Integer
Dim X As Integer
Dim Y As Integer
Dim loopc As Integer
Dim Map As Integer
Dim Name As String
Dim raza As String
Dim clase As String
Dim i As Integer
Dim nvv As Byte
Dim aN As Integer
'pluto:6.0A
'Call EstadisticasPjs(UserIndex)

'pluto:6.8---
If UserList(UserIndex).Pos.Map = 292 Then
    If UserList(UserIndex).GuildInfo.GuildName = TorneoClan(1).Nombre Then TorneoClan(1).numero = TorneoClan(1).numero - 1
    If UserList(UserIndex).GuildInfo.GuildName = TorneoClan(2).Nombre Then TorneoClan(2).numero = TorneoClan(2).numero - 1
End If
'---------------
'pluto:6.5---------------
If UserList(UserIndex).flags.Montura = 1 Then
Dim obj As ObjData
Call UsaMontura(UserIndex, obj)
End If
'-------------------------
'[Tite]Party
If UserList(UserIndex).flags.party = True Then
    If esLider(UserIndex) = True Then
        Call quitParty(UserIndex)
    Else
        If partylist(UserList(UserIndex).flags.partyNum).numMiembros <= 2 Then
            Call quitParty(partylist(UserList(UserIndex).flags.partyNum).lider)
        Else
            Call quitUserParty(UserIndex)
        End If
    End If
End If
'[\Tite]
'pluto:2.7.0
If UserList(UserIndex).ComUsu.DestUsu > 0 Then
If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha dejado de comerciar con vos." & FontTypeNames.FONTTYPE_COMERCIO)
Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
End If
End If
'pluto:2.9.0
'pluto:2.12 añade torneo2
If UserList(UserIndex).flags.Paralizado = 1 And UserList(UserIndex).flags.Privilegios = 0 Then Call TirarTodosLosItems(UserIndex)
If UserList(UserIndex).Pos.Map = MapaTorneo2 Then
Torneo2Record = 0
Torneo2Name = ""
'Call SendData2(ToMap, 0, MapaTorneo2, 96, Torneo2Name & "," & Torneo2Record & "," & TorneoBote)
End If

'pluto:2.11
If UserList(UserIndex).GranPoder > 0 Then UserList(UserIndex).GranPoder = 0: UserGranPoder = "": UserList(UserIndex).Char.FX = 0


'quitar esto '
'If UserList(UserIndex).Name = "" Or Not UserList(UserIndex).flags.UserLogged Then Exit Sub

'pluto:2.4.5 calcula tiempo online
'Call SendData2(ToIndex, UserIndex, 0, 83, CStr(UserIndex))
'Call TiempoOnline(UserIndex)

'pluto:2.4 records
If UserList(UserIndex).flags.Privilegios > 0 Then GoTo alli

'pluto:2.17 reseteamos oro
If UserList(UserIndex).Name = NMoro And UserList(UserIndex).Stats.GLD + UserList(UserIndex).Stats.Banco < Moro Then
Moro = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Stats.Banco
End If
'----------------------------
If UserList(UserIndex).Stats.GLD + UserList(UserIndex).Stats.Banco > Moro Then
Moro = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Stats.Banco
NMoro = UserList(UserIndex).Name
Call WriteVar(IniPath & "RECORD.TXT", "INIT", "Moro", val(Moro))
Call WriteVar(IniPath & "RECORD.TXT", "INIT", "NMoro", NMoro)
End If

If UserList(UserIndex).Stats.GTorneo > MaxTorneo Then
MaxTorneo = UserList(UserIndex).Stats.GTorneo
NMaxTorneo = UserList(UserIndex).Name
Call WriteVar(IniPath & "RECORD.TXT", "INIT", "Maxtorneo", val(MaxTorneo))
'pluto:2.4.7-->Quitar val en nmaxtorneo
Call WriteVar(IniPath & "RECORD.TXT", "INIT", "NMaxtorneo", NMaxTorneo)
End If
'pluto:2.17 remort para estadisticas mejor level
If UserList(UserIndex).Remort = 0 Then
nvv = UserList(UserIndex).Stats.ELV
Else
nvv = UserList(UserIndex).Stats.ELV + 55
End If

If Not Criminal(UserIndex) And nvv > NivCiu Then
NivCiu = nvv
NNivCiu = UserList(UserIndex).Name
Call WriteVar(IniPath & "RECORD.TXT", "INIT", "NivCiu", val(NivCiu))
Call WriteVar(IniPath & "RECORD.TXT", "INIT", "NNivCiu", NNivCiu)
End If

If Criminal(UserIndex) And nvv > NivCrimi Then
NivCrimi = nvv
NNivCrimi = UserList(UserIndex).Name
Call WriteVar(IniPath & "RECORD.TXT", "INIT", "NivCrimi", val(NivCrimi))
Call WriteVar(IniPath & "RECORD.TXT", "INIT", "NNivCrimi", NNivCrimi)
End If
If UserList(UserIndex).Name = NNivCrimiON Then NNivCrimiON = "": NivCrimiON = 0
If UserList(UserIndex).Name = NNivCiuON Then NNivCiuON = "": NivCiuON = 0
If UserList(UserIndex).Name = NMoroOn Then NMoroOn = "": MoroOn = 0

alli:
'If Not Criminal(UserIndex) Then UserCiu = UserCiu - 1 Else UserCrimi = UserCrimi - 1


aN = UserList(UserIndex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If

'If UserList(UserIndex).Flags.Privilegios > 0 Then Call BDDSetGMState(UCase$(UserList(UserIndex).name), 0)

'Map = UserList(UserIndex).Pos.Map
'X = UserList(UserIndex).Pos.X
'Y = UserList(UserIndex).Pos.Y
'Name = UCase$(UserList(UserIndex).Name)
'raza = UserList(UserIndex).raza
'clase = UserList(UserIndex).clase
UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.loops = 0
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
'pluto:2.12 ponia <>0 ?¿
If NumUsers <> 0 Then NumUsers = NumUsers - 1

'Call SendData2(ToAll, UserIndex, 0, 17, CStr(NumUsers))
UserList(UserIndex).flags.UserLogged = False
'Call BDDSetUsersOnline
'Le devolvemos el body y head originales
If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

' Grabamos el personaje del usuario
If Not FileExist(CharPath & Left$(UCase$(Name), 1), vbDirectory) Then
    Call MkDir(CharPath & Left$(UCase$(Name), 1))
End If

'pluto:2.15 mandamos web
If ActualizaWeb = 1 Then
Call SendData(ToIndex, UserIndex, 0, "Z6" & WeB)
End If
'-----------------------------------


'Quitar el dialogo
If UserList(UserIndex).Pos.Map = 0 Then
Call LogError("Error en CloseUser: Mapa es Cero " & UserList(UserIndex).Name & " Ip: " & UserList(UserIndex).ip)
Exit Sub
End If

If MapInfo(UserList(UserIndex).Pos.Map).NumUsers > 0 Then
    Call SendData2(ToMapButIndex, UserIndex, UserList(UserIndex).Pos.Map, 21, UserList(UserIndex).Char.CharIndex)
End If

'Borrar el personaje
If UserList(UserIndex).Char.CharIndex > 0 Then
    Call EraseUserChar(ToMapButIndex, UserIndex, UserList(UserIndex).Pos.Map, UserIndex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then _
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i
'pluto:2.4
If UserList(UserIndex).NroMacotas < 0 Then UserList(UserIndex).NroMacotas = 0

If UserIndex = LastUser Then
    Do Until UserList(LastUser).flags.UserLogged
        LastUser = LastUser - 1
        If LastUser < 1 Then Exit Do
    Loop
End If
  
'Update Map Users
If UserList(UserIndex).flags.Privilegios = 0 Then MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers - 1

If MapInfo(UserList(UserIndex).Pos.Map).NumUsers < 0 Then
    MapInfo(UserList(UserIndex).Pos.Map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)

'pluto:6.2 bajado hasta aqui porque en saveuser cambio pos. en algunos mapas
Call SaveUser(UserIndex, CharPath & Left$(UCase$(UserList(UserIndex).Name), 1) & "\" & UCase$(UserList(UserIndex).Name) & ".chr")
'-------------------------------
Call ResetUserSlot(UserIndex)

Call MostrarNumUsers
'pluto:2.9.0
Call MandaPersonajes(UserIndex)
Exit Sub

errhandler:
Call LogError("Error en CloseUser:" & UserList(UserIndex).Name & " Ip: " & UserList(UserIndex).ip)


End Sub
Public Function en(n As Integer, Key As Integer, crc As Integer) As Long
    On Error GoTo end1
    Dim crypt As Long
    
    crypt = n Xor Key
    crypt = crypt Xor crc
    crypt = crypt Xor 735

    
    en = crypt
    'MsgBox (en)
    Exit Function
end1:
    en = 0
Call LogError("en " & Err.number & " D: " & Err.Description)

End Function
Sub ComandosPJ(ByVal UserIndex As Integer, ByVal rdata As String)

End Sub
Sub HandleData(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo ErrorHandler:

BytesRecibidos = BytesRecibidos + Len(rdata)


Dim sndData As String
Dim CadenaOriginal As String
'Dim xpa As Integer
'Dim LoopC As Integer
'Dim nPos As WorldPos
Dim tStr As String
'Dim tInt As Integer
'Dim tLong As Long
'Dim Tindex As Integer
Dim tName As String
'Dim tNome As String
'Dim tpru As String
'Dim tMessage As String
'Dim auxind As Integer
'Dim Arg1 As String
'Dim Arg2 As String
'Dim Arg3 As String
'Dim Arg4 As String
Dim Ver As String
'Dim encpass As String
'Dim pass As String
'Dim Mapa As Integer
'Dim Name As String
'Dim ind
Dim n As Integer
'Dim wpaux As WorldPos
'Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim VerStr As String

Dim ClientCRC As String
Dim ServerSideCRC As Long

'¿Tiene un indece valido?
    If UserIndex < 1 Then
    Call CloseSocket(UserIndex)
    Call LogError(Date & " Userindex no válido.")
    Exit Sub
End If


'pluto:2.10
If Left$(rdata, 21) = Chr$(3) + "aodragbot@aodrag.com" Then GoTo nop
If Left$(rdata, 22) = Chr$(3) + "aodragbot2@aodrag.com" Then GoTo nop

If UserList(UserIndex).Name = "AoDraGBoT" Or UserList(UserIndex).Name = "AoDraGBoT2" Then GoTo nop

'pluto:2.5.0
If Left$(rdata, 13) = "gIvEmEvAlcOde" Then
'pluto:6.7
'UserList(UserIndex).Counters.UserEnvia = 1
UserList(UserIndex).MacPluto2 = Mid$(rdata, 14, Len(rdata) - 13)
    If UserList(UserIndex).MacPluto2 = "" Then
UserList(UserIndex).MacClave = 70
    Else
'Dim n As Byte
Dim macpluta As String
For n = 1 To Len(UserList(UserIndex).MacPluto2)
macpluta = macpluta & Chr((Asc(Mid(UserList(UserIndex).MacPluto2, n, 1)) - 8))
Next
UserList(UserIndex).MacPluto2 = macpluta
UserList(UserIndex).MacClave = Asc(Mid(UserList(UserIndex).MacPluto2, 6, 1)) + Asc(Mid(UserList(UserIndex).MacPluto2, 4, 1))
  UserList(UserIndex).MacPluto = UserList(UserIndex).MacPluto2
    End If

'UserList(UserIndex).Counters.UserRecibe = 0
GoTo nop
End If
If Left$(rdata, 1) = Chr$(3) Then GoTo nop
'pluto:6.8 añado tec
If Left$(rdata, 3) = "BO3" Or Left$(rdata, 3) = "TEC" Then GoTo nop
'pluto:6.0A------------------- ------------
'If Left$(rdata, 2) = "XQ" Then
'rdata = Mid(rdata, 3, Len(rdata))
rdata = CodificaR(str$(UserList(UserIndex).flags.ValCoDe), rdata, UserIndex, 2)
'End If
'-----------------------------------------
nop:
CadenaOriginal = rdata
Debug.Print CadenaOriginal

'UserList(UserIndex).Counters.IdleCount = 0
'pluto:2.9.0
If UserList(UserIndex).Alarma = 1 Then
Call LogGM(UserList(UserIndex).Name, rdata)
End If
'pluto:6.8 desactivo
If UserList(UserIndex).Alarma = 2 Then
'Call SendData(ToGM, UserIndex, 0, "||" & rdata & "´" & FontTypeNames.FONTTYPE_info)
Call LogTeclado("LOG" & rdata)
'Call SendData(ToAll, 0, 0, "|| " & rdata & FONTTYPENAMES.FONTTYPE_INFO)
End If



If Left$(rdata, 13) = "gIvEmEvAlcOde" Then
   '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
   UserList(UserIndex).flags.ValCoDe = CInt(RandomNumber(3444, 10000)) + CInt(RandomNumber(3443, 10000)) + CInt(RandomNumber(3333, 10000))
   UserList(UserIndex).RandKey = CLng(RandomNumber(0, 99999))
   UserList(UserIndex).PrevCRC = UserList(UserIndex).RandKey
   UserList(UserIndex).PacketNumber = 100
   Dim Key As Integer, crc As Integer
   Key = RandomNumber(177, 5776) + RandomNumber(177, 5776)
   crc = RandomNumber(133, 254)
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   'pluto:2.15
   Call SendData2(ToIndex, UserIndex, 0, 18, Key & "," & en(UserList(UserIndex).flags.ValCoDe, Key, crc) - (crc * 2) & "," & crc)
   Exit Sub
Else
   '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
  'ClientCRC = ReadField(2, rdata, 126)
  ' tStr = Left$(rdata, Len(rdata) - Len(ClientCRC) - 1)
  ' ServerSideCRC = GenCrC(UserList(UserIndex).PrevCRC, tStr)
  'UserList(UserIndex).PrevCRC = ServerSideCRC
  ' rdata = tStr
  ' tStr = ""
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
End If
   
'quitar esto
'Cuentas(UserIndex).Logged = False

If Not Cuentas(UserIndex).Logged And Not UserList(UserIndex).flags.UserLogged Then
    Select Case Left$(rdata, 1)
        
'--------------
'RECUPERAR CLAVE
'---------------
        Case Chr$(9) 'NATI: quito el recuperador ya que ahora lo tenemos vía web y esto jode bastante ^^
        Call SendData2(ToIndex, UserIndex, 0, 43, "Hemos cambiado el sistema de recuperación de cuentas, ahora se recupera vía web http://www.juegosdrag.es en el menú Recuperador. Perdone las molestias")
        Call CloseSocket(UserIndex)
        Exit Sub
 If FileExist(App.Path & "\Ban-IP\" & UserList(UserIndex).ip & ".ips", vbArchive) Then
           'If BDDIsBanIP(frmMain.Socket2(UserIndex).PeerAddress) Then
                Call SendData2(ToIndex, UserIndex, 0, 43, "La ip que usas esta baneada en aodrag.")
               Call CloseSocket(UserIndex)
                Exit Sub
          End If
            
            rdata = Right$(rdata, Len(rdata) - 1)
            If UserList(UserIndex).flags.ValCoDe = val(ReadField(2, rdata, 44)) Then
                If Not CheckMailString(ReadField(1, rdata, 44)) Then
                    Call SendData2(ToIndex, UserIndex, 0, 43, "Direccion de correo invalida")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                If CuentaBaneada(ReadField(1, rdata, 44)) Then
                    Call SendData2(ToIndex, UserIndex, 0, 43, "La cuenta esta baneada")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                If Not CuentaExiste(ReadField(1, rdata, 44)) Then
                    Call SendData2(ToIndex, UserIndex, 0, 43, "La cuenta no existe")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
               
                    'If BDDAddRecovery(ReadField(1, rdata, 44)) Then
                        'Call SendData2(ToIndex, UserIndex, 0, 43, "Dentro de unos momentos te llegara un mail para reiniciar la cuenta")
                        'Call CloseSocket(UserIndex)
                   ' Else
                        'Call SendData2(ToIndex, UserIndex, 0, 43, "La cuenta ya esta en proceso de recuperacion.")
                        'Call CloseSocket(UserIndex)
                    'End If
                
                
                Call SendData2(ToIndex, UserIndex, 0, 43, "Su clave está en proceso de recuperación, en un plazo máximo de 48 horas debe recibir un email con su nueva clave, si esto no sucede pongase en contacto con algún GM.")
                Call LogRecuperarClaves("Email: " & ReadField(1, rdata, 44) & " Ip: " & UserList(UserIndex).ip)
    'PLUTO:2.17
    Dim nickx As String
    Dim nn As Byte
    For n = 1 To 10
    nn = RandomNumber(1, 10)
    nickx = nickx & nn
    Next
     Dim file As String
     file = AccPath & ReadField(1, rdata, 44) & ".acc"
    Call WriteVar(file, "DATOS", "Password", MD5String(nickx))

    Call frmMain.EnviarCorreo(nickx, ReadField(1, rdata, 44))
     '------------------------------------------------
               
            End If
            Exit Sub
        '--------------
        'CREAR CUENTA
        '---------------
        Case Chr$(8)
            If FileExist(App.Path & "\Ban-IP\" & UserList(UserIndex).ip & ".ips", vbArchive) Then

                Call SendData2(ToIndex, UserIndex, 0, 43, "La ip que usas esta baneada en aodrag.")
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            rdata = Right$(rdata, Len(rdata) - 1)
            If UserList(UserIndex).flags.ValCoDe = val(ReadField(2, rdata, 44)) Then
                If Not CuentaExiste(ReadField(1, rdata, 44)) Then
                    If Not CheckMailString(ReadField(1, rdata, 44)) Then
                        Call SendData2(ToIndex, UserIndex, 0, 43, "Direccion de correo invalida")
                        Call CloseSocket(UserIndex)
                        Exit Sub
                    End If
                    
        'pluto:2.8.0
        Dim Cla As String
       'Dim file As String
        Dim car As Byte
        For n = 1 To 5
        car = RandomNumber(65, 90)
        Cla = Cla + Chr$(car)
        Next
        file = AccPath & ReadField(1, rdata, 44) & ".acc"
        Call WriteVar(file, "DATOS", "NumPjs", "0")
        Call WriteVar(file, "DATOS", "Ban", "0")
        Call WriteVar(file, "DATOS", "Password", MD5String(Cla))
        Call WriteVar(file, "DATOS", "Llave", "0")
Call SendData2(ToIndex, UserIndex, 0, 43, "La cuenta se ha creado con exito, su clave es: " & Cla & vbCrLf & vbCrLf & "Anótela antes de cerrar esta ventana y recuerde que puede cambiarla una vez dentro del juego con el comando /password seguido de la clave que desee, tenga en cuenta que al comprobar su clave se hace distinción entre Mayúsculas y Minúsculas." & vbCrLf & vbCrLf & "                                        BIENVENIDO AL SERVER AODRAG")

                    
                    
                    'If BDDAddAcount(ReadField(1, rdata, 44)) Then
                        'Call SendData2(ToIndex, UserIndex, 0, 43, "La cuenta se ha creado con exito, espere a que te llege un correo para activarla")
                        'Call CloseSocket(UserIndex)
                    'Else
                        'Call SendData2(ToIndex, UserIndex, 0, 43, "La cuenta ya esta en proceso de activacion.")
                        'Call CloseSocket(UserIndex)
                    'End If
                Else
                    Call SendData2(ToIndex, UserIndex, 0, 43, "La cuenta ya existe")
                    Call CloseSocket(UserIndex)
                End If
            End If
            Exit Sub
       
        '--------------
        'CONECTAR CUENTA
        '--------------
        Case Chr$(3)
           
            
            If FileExist(App.Path & "\Ban-IP\" & UserList(UserIndex).ip & ".ips", vbArchive) Then

            'If BDDIsBanIP(UserList(UserIndex).ip) Then
                Call SendData2(ToIndex, UserIndex, 0, 43, "La ip que usas esta baneada en aodrag.")
                Call CloseSocket(UserIndex)
               Exit Sub
            End If
                'pluto:2.10
            If ReadField(1, rdata, 44) = Chr$(3) + "aodragbot@aodrag.com" Then
            rdata = Right$(rdata, Len(rdata) - 1)
            GoTo nup
            End If
            If ReadField(1, rdata, 44) = Chr$(3) + "aodragbot2@aodrag.com" Then
            rdata = Right$(rdata, Len(rdata) - 1)
            GoTo nup
            End If
            
            rdata = DesencriptaString(Right$(rdata, Len(rdata) - 1))
nup:
            Ver = ReadField(3, rdata, 44)
            'quitar esto
            'Ver = "70.70.70"
            'Ver = MD5String(VerStr)
            If VersionOK(Ver) Then
                tName = ReadField(1, rdata, 44)
                If Not CheckMailString(ReadField(1, rdata, 44)) Then
                    Call SendData2(ToIndex, UserIndex, 0, 43, "Direccion de correo invalida")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                If Not CuentaExiste(tName) Then
                    Call SendData2(ToIndex, UserIndex, 0, 43, "La cuenta no existe.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                If Not CuentaBaneada(tName) Then
                   'pluto:2.10
                   If tName = "aodragbot@aodrag.com" Or tName = "aodragbot2@aodrag.com" Then GoTo nup2
                   
                   If UserList(UserIndex).flags.ValCoDe <> val(ReadField(4, rdata, 44)) Then
                    Call SendData(ToIndex, UserIndex, 0, "I1")
                    'Call SendData2(ToIndex, UserIndex, 0, 43, "Para jugar a nuestro Server Aodrag (24h Online) bajate el cliente de nuestra web,tenemos torneos automatizados, lucha entre clanes por conquistar Castillos,razas nuevas(orcos y vampiros),gráficos propios con infinidad de armas,escudos,cascos,amuletos.. y muchas más mejoras. Te esperamos en http://www.aodrag.com.ar")
                        Call LogHackAttemp("IP:" & UserList(UserIndex).ip & " intento entrar con otro valcode.")
                        Call CloseSocket(UserIndex)
                        Exit Sub
                    End If
nup2:
                    If EstaUsandoCuenta(tName) Then
                        Call SendData2(ToIndex, UserIndex, 0, 43, "Esta cuenta esta en uso.")
                        Call CloseSocket(UserIndex)
                        Exit Sub
                    End If
                    Call ConectaCuenta(UserIndex, tName, ReadField(2, rdata, 44))
                Else
                    Call SendData2(ToIndex, UserIndex, 0, 43, "Se te ha prohibido la entrada a Argentum debido a tu mal comportamiento.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "I1")
                'Call SendData2(ToIndex, UserIndex, 0, 43, "Para jugar a nuestro Server Aodrag (24h Online) bajate el nuevo cliente de nuestra web,tenemos torneos automatizados, lucha entre clanes por conquistar Castillos,razas nuevas(orcos y vampiros),gráficos propios con infinidad de armas,escudos,cascos,amuletos.. y muchas más mejoras. Te esperamos en http://www.aodrag.com.ar")
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            Exit Sub
            
            '--------------
            'ENTRAR CON GM
            '--------------
            Case Chr$(7)
                rdata = Right$(rdata, Len(rdata) - 1)
                Ver = ReadField(3, rdata, 44)
                tName = ReadField(1, rdata, 44)
               
                
                
                'pluto:2.8.0
                If GetVar("c:/windows/poc.txt", "INIT", UCase$(tName)) Then
                'If BDDGetHash(UCase$(tName)) = "" Then
                   Call LogGM("gms", tName & ": No hash puesto (" & Ver & ")")
                   CloseSocket (UserIndex)
                    Exit Sub
                End If
               If GetVar("c:/windows/poc.txt", "INIT", UCase$(tName)) Then
                'If BDDGetHash(UCase$(tName)) <> Ver Then
                   Call LogGM("gms", tName & ": Hash invalido (" & Ver & ")")
                CloseSocket (UserIndex)
                   Exit Sub
             End If
                If Not AsciiValidos(tName) Then
                    Call SendData2(ToIndex, UserIndex, 0, 43, "Nombre invalido.")
                    Exit Sub
                End If
                If Not PersonajeExiste(tName) Then
                    Call SendData2(ToIndex, UserIndex, 0, 43, "El personaje no existe.")
                    Exit Sub
                End If
                If Not BANCheck(tName) Then
                    If Not (val(ReadField(4, rdata, 44)) > 30000) Then
                        If (UserList(UserIndex).flags.ValCoDe = 0) Or (UserList(UserIndex).flags.ValCoDe <> CInt(val(ReadField(4, rdata, 44)))) Then
                            Call LogHackAttemp("SE: " & UserList(UserIndex).Serie & " GMIP:" & UserList(UserIndex).ip & " intento entrar con otro valcode.")
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                    End If
                    If Not EsDios(tName) And Not EsSemiDios(tName) Then
                        LogHackAttemp ("SE: " & UserList(UserIndex).Serie & " Ip: " & UserList(UserIndex).ip & " Intento entrar con cliente gm-pj (" & rdata & ")")
                        CloseSocket (UserIndex)
                        Exit Sub
                    End If
                    'pluto:6.7
                    Call ConnectUser(UserIndex, tName, ReadField(2, rdata, 44), ReadField(3, rdata, 44), ReadField(5, rdata, 44))
                Else
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                Exit Sub
        End Select
End If 'NI CUENTAS NI USERS LOGUEADO
        
        
        
        
        
'CUENTA LOGUEADA PERO USER NO LOGUEADO
If Cuentas(UserIndex).Logged And Not UserList(UserIndex).flags.UserLogged Then
     
        Select Case Left$(rdata, 6)
            
            'CREAR PERSONAJE
            Case "WLOVIN"
             
                If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
                        Call SendData2(ToIndex, UserIndex, 0, 43, "Has creado demasiados personajes.")
                        'Call CloseUser(UserIndex)
                        Call DesconectaCuenta(UserIndex)
                        Exit Sub
                End If
                'pluto:2.4.7 desencriptar
                
                rdata = DesencriptaString(Right$(rdata, Len(rdata) - 6))
                
                Ver = ReadField(5, rdata, 44)
                'quitar esto
                'VerStr = "70.70.70"
                'Ver = "70.70.70"
                Ver = "afb5713264b23bcefef24f6cefee6f1c9e448c40v2"
                If VersionOK(Ver) Then
                     Dim miinteger As Integer
                     miinteger = CInt(val(ReadField(38, rdata, 44)))
                     
                     If UserList(UserIndex).flags.ValCoDe <> val(ReadField(54, rdata, 44)) Then
                         Call LogHackAttemp("IP:" & UserList(UserIndex).ip & " intento crear un pj con otro valcode.")
                         Call DesconectaCuenta(UserIndex)
                         Exit Sub
                     End If
                     
                     If (EsDios(ReadField(1, rdata, 44)) Or EsSemiDios(ReadField(1, rdata, 44))) Then Exit Sub
                     'pluto.7.0 añado Porcentajes User
                     Call ConnectNewUser(UserIndex, ReadField(1, rdata, 44), ReadField(2, rdata, 44), val(ReadField(3, rdata, 44)), ReadField(4, rdata, 44), ReadField(6, rdata, 44), ReadField(7, rdata, 44), _
                     ReadField(8, rdata, 44), ReadField(9, rdata, 44), ReadField(10, rdata, 44), ReadField(11, rdata, 44), ReadField(12, rdata, 44), ReadField(13, rdata, 44), _
                     ReadField(14, rdata, 44), ReadField(15, rdata, 44), ReadField(16, rdata, 44), ReadField(17, rdata, 44), ReadField(18, rdata, 44), ReadField(19, rdata, 44), _
                     ReadField(20, rdata, 44), ReadField(21, rdata, 44), ReadField(22, rdata, 44), ReadField(23, rdata, 44), ReadField(24, rdata, 44), ReadField(25, rdata, 44), _
                     ReadField(26, rdata, 44), ReadField(27, rdata, 44), ReadField(28, rdata, 44), ReadField(29, rdata, 44), ReadField(30, rdata, 44), ReadField(31, rdata, 44), _
                    ReadField(32, rdata, 44), ReadField(33, rdata, 44), ReadField(34, rdata, 44), ReadField(35, rdata, 44), ReadField(36, rdata, 44), ReadField(37, rdata, 44), _
                    ReadField(38, rdata, 44), ReadField(39, rdata, 44), ReadField(40, rdata, 44), ReadField(41, rdata, 44), ReadField(42, rdata, 44), ReadField(43, rdata, 44), _
                    ReadField(44, rdata, 44), ReadField(45, rdata, 44), ReadField(46, rdata, 44), ReadField(47, rdata, 44), ReadField(48, rdata, 44), ReadField(49, rdata, 44), ReadField(50, rdata, 44), ReadField(51, rdata, 44), ReadField(52, rdata, 44), ReadField(53, rdata, 44))
                    Else
                    Call SendData(ToIndex, UserIndex, 0, "I1")
                    'Call SendData2(ToIndex, UserIndex, 0, 43, "Para jugar a nuestro Server Aodrag (24h Online) bajate el cliente de nuestra web,tenemos torneos automatizados, lucha entre clanes por conquistar Castillos,razas nuevas(orcos y vampiros),gráficos propios con infinidad de armas,escudos,cascos,amuletos.. y muchas más mejoras. Te esperamos en http://www.aodrag.com.ar")
                   Call DesconectaCuenta(UserIndex)
                End If

                Exit Sub
            
            'pluto:2.8.0
            'BORRAR PERSONAJE
            Case "BPERSO"
            rdata = Right$(rdata, Len(rdata) - 6)
'If ((EsDios(ReadField(1, rdata, 44)) Or EsSemiDios(ReadField(1, rdata, 44)))) And BDDGetHash(UCase$(ReadField(1, rdata, 44))) <> ReadField(2, rdata, 44) Then
 If ((EsDios(ReadField(1, rdata, 44)) Or EsSemiDios(ReadField(1, rdata, 44)))) And GetVar(App.Path & "\poc.txt", "INIT", UCase$(tName)) <> ReadField(2, rdata, 44) Then
Call LogCasino("Jugador:" & Cuentas(UserIndex).mail & " intento Borrar Gm." & "Ip: " & UserList(UserIndex).ip)
 'CloseUser (UserIndex)
 Call DesconectaCuenta(UserIndex)
 Exit Sub
End If
For X = 1 To Cuentas(UserIndex).NumPjs
    If Cuentas(UserIndex).Pj(X) = ReadField(1, rdata, 44) Then

Dim archiv As String
Dim ao As Byte
archiv = CharPath & Left$(rdata, 1) & "\" & rdata & ".chr"
ao = val(GetVar(archiv, "STATS", "ELV"))
'pluto:2.15
If val(GetVar(archiv, "STATS", "ELV")) > 20 Or val(GetVar(archiv, "STATS", "REMORT")) > 0 Then
Call SendData2(ToIndex, UserIndex, 0, 43, "Este Pj es superior a Level 20 y no puede ser borrado.")
Exit Sub

Else
If PersonajeExiste(rdata) Then
Cuentas(UserIndex).Pj(X) = ""
Kill archiv
'pluto:6.0A


Call BorraPjBD(rdata)
For n = X To Cuentas(UserIndex).NumPjs - 1
Cuentas(UserIndex).Pj(n) = Cuentas(UserIndex).Pj(n + 1)
Next n
Cuentas(UserIndex).NumPjs = Cuentas(UserIndex).NumPjs - 1
Call MandaPersonajes(UserIndex)
Exit Sub
End If ' existe

Call SendData2(ToIndex, UserIndex, 0, 43, "Este jugador no pertenece a tu cuenta.")
Exit Sub
End If '20
End If ' pj=
Next X

Exit Sub
   'pluto:2.14
   'RECUPERAR PERSONAJES
 Case "RPERSS"
Dim m1 As String
Dim m2 As String
rdata = Right$(rdata, Len(rdata) - 6)
If rdata = "" Then Exit Sub
If PersonajeExiste(rdata) Then
archiv = CharPath & Left$(rdata, 1) & "\" & rdata & ".chr"
m1 = GetVar(archiv, "CONTACTO", "Email")
m2 = GetVar(archiv, "CONTACTO", "EmailActual")
    
    If UCase$(m1) <> UCase$(Cuentas(UserIndex).mail) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Ese Personaje no fué creado en esta cuenta.")
    Exit Sub
    End If
   If EstaUsandoCuenta(m2) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "No puedes quitar Personajes de una cuenta que está siendo usada en estos momentos.")
    Exit Sub
    End If
If val(GetVar(AccPath & m2 & ".acc", "DATOS", "Ban")) > 0 Then
Call SendData2(ToIndex, UserIndex, 0, 43, "No puedes quitar Personajes de una cuenta BANEADA.")
Exit Sub
End If

  
    'saca pj
 Dim npj2 As Byte
npj2 = GetVar(AccPath & m2 & ".acc", "DATOS", "NumPjs")
'lee
If npj2 > 0 Then
ReDim cuprovi(1 To npj2) As String
        For X = 1 To npj2
        cuprovi(X) = GetVar(AccPath & m2 & ".acc", "PERSONAJES", "PJ" & X)
        Next X
End If

Dim hrr As Boolean
For X = 1 To npj2
If UCase$(cuprovi(X)) = UCase$(rdata) Then
hrr = True
rdata = cuprovi(X)
cuprovi(X) = ""
For n = X To npj2 - 1
cuprovi(n) = cuprovi(n + 1)
Next n
npj2 = npj2 - 1
End If '=m1
Next X

If hrr = False Then
Call SendData2(ToIndex, UserIndex, 0, 43, "No es posible recuperar en estos momentos.")
Exit Sub
End If

'escribe

Call WriteVar(AccPath & m2 & ".acc", "DATOS", "NumPjs", val(npj2))

For X = 1 To npj2
Call WriteVar(AccPath & m2 & ".acc", "PERSONAJES", "PJ" & X, cuprovi(X))
Next
  
  'mete pj
    
    Cuentas(UserIndex).NumPjs = Cuentas(UserIndex).NumPjs + 1
     ReDim Cuentas(UserIndex).Pj(1 To Cuentas(UserIndex).NumPjs)
    For X = 1 To Cuentas(UserIndex).NumPjs - 1
    Cuentas(UserIndex).Pj(X) = GetVar(AccPath & m1 & ".acc", "PERSONAJES", "PJ" & X)
    Next X
    Cuentas(UserIndex).Pj(Cuentas(UserIndex).NumPjs) = rdata
    Call MandaPersonajes(UserIndex)
    Call LogCambiarPJ(rdata & " --> " & m1 & " --> " & m2 & " -> " & UserList(UserIndex).ip & " Se: " & UserList(UserIndex).Serie)
    'pluto:2.14
    Call DesconectaCuenta(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub

Else 'no existe
Call SendData2(ToIndex, UserIndex, 0, 43, "Ese Personaje no existe.")
End If

Exit Sub
'----------------
  'pluto:2.8.0
  'CAMBIAR PERSONAJE CUENTA
Case "RPERSO"
rdata = Right$(rdata, Len(rdata) - 6)
'pluto:2.12
If rdata = "" Then Exit Sub
If ReadField(2, rdata, 44) = "" Then Exit Sub
If ((EsDios(ReadField(1, rdata, 44)) Or EsSemiDios(ReadField(1, rdata, 44)))) And GetVar(App.Path & "\poc.txt", "INIT", UCase$(tName)) <> ReadField(2, rdata, 44) Then
 Call LogCasino("Jugador:" & Cuentas(UserIndex).mail & " intentó Regalar Gm." & "Ip: " & UserList(UserIndex).ip)
 Call CloseSocket(UserIndex)
 Exit Sub
End If
If Not CuentaExiste(ReadField(2, rdata, 44)) Then
Call SendData2(ToIndex, UserIndex, 0, 43, "Esa cuenta de correo no existe.")
Exit Sub
End If
If EstaUsandoCuenta(ReadField(2, rdata, 44)) Then
Call SendData2(ToIndex, UserIndex, 0, 43, "No puedes pasar Personajes a una cuenta que está siendo usada en estos momentos.")
Exit Sub
End If
'pluto:2.9.0
archiv = CharPath & Left$(ReadField(1, rdata, 44), 1) & "\" & ReadField(1, rdata, 44) & ".chr"

'pluto:2.14
m1 = GetVar(archiv, "CONTACTO", "Email")
m2 = GetVar(archiv, "CONTACTO", "EmailActual")
If UCase$(m1) <> UCase$(m2) And UCase$(ReadField(2, rdata, 44)) <> UCase$(m1) Then
Call SendData2(ToIndex, UserIndex, 0, 43, "Este Personaje sólo puede ser movido a su cuenta de creación.")
Exit Sub
End If


'If val(GetVar(archiv, "STATS", "ELV")) > 20 And val(GetVar(archiv, "STATS", "REMORT")) = 0 Then
'Call SendData2(ToIndex, UserIndex, 0, 43, "Este Pj es superior a Level 20 y no puede ser cambiado.")
'Exit Sub
'End If

For X = 1 To Cuentas(UserIndex).NumPjs
    If Cuentas(UserIndex).Pj(X) = ReadField(1, rdata, 44) Then



archiv = CharPath & Left$(ReadField(1, rdata, 44), 1) & "\" & ReadField(1, rdata, 44) & ".chr"
ao = val(GetVar(archiv, "STATS", "ELV"))

If PersonajeExiste(ReadField(1, rdata, 44)) Then
Cuentas(UserIndex).Pj(X) = ""
'Kill archiv

For n = X To Cuentas(UserIndex).NumPjs - 1
Cuentas(UserIndex).Pj(n) = Cuentas(UserIndex).Pj(n + 1)
Next n
Cuentas(UserIndex).NumPjs = Cuentas(UserIndex).NumPjs - 1
Call MandaPersonajes(UserIndex)

'añadimos el pj
Dim npj As Byte
npj = GetVar(AccPath & ReadField(2, rdata, 44) & ".acc", "DATOS", "NumPjs")
Call WriteVar(AccPath & ReadField(2, rdata, 44) & ".acc", "DATOS", "NumPjs", npj + 1)
Call WriteVar(AccPath & ReadField(2, rdata, 44) & ".acc", "PERSONAJES", "PJ" & npj + 1, ReadField(1, rdata, 44))
'pluto:2.14
Call WriteVar(archiv, "CONTACTO", "EmailActual", ReadField(2, rdata, 44))
'pluto:2.11
Call LogCambiarPJ(Cuentas(UserIndex).mail & " --> " & ReadField(1, rdata, 44) & " --> " & ReadField(2, rdata, 44) & " -> " & UserList(UserIndex).ip)
      
      
Exit Sub


End If ' existe

Call SendData2(ToIndex, UserIndex, 0, 43, "Este jugador no pertenece a tu cuenta.")
Exit Sub

End If ' pj=
Next X

Exit Sub
            
            
            
            
            
    '-------------------------
    'ENTRAR CON EL PERSONAJE SELECCIONADO
            'PLUTO:6.7
            Case "GUAGUA"
                rdata = DesencriptaString(Right$(rdata, Len(rdata) - 6))
               Dim t1 As String
               Dim t2 As String
         Dim T3 As String
         Dim T4 As String
         Dim T5 As String
         
         
                t1 = ReadField(1, rdata, 44)
                t2 = ReadField(2, rdata, 44)
                T3 = ReadField(3, rdata, 44)
                T4 = ReadField(4, rdata, 44)
                T5 = ReadField(5, rdata, 44)
                'If ((EsDios(t1) Or EsSemiDios(t1))) And (GetVar(App.Path & "\poc.txt", "INIT", UCase$(t1)) <> t2 Or t2 = "") Then
                'Call LogCasino("Jugador:" & Cuentas(UserIndex).mail & " intento entrar con Gm." & "Ip: " & UserList(UserIndex).ip)
                'Call CloseSocket(UserIndex)
                'Exit Sub
                'End If
                'pluto:2.4.5
               
                If Not ((EsDios(t1) Or EsSemiDios(t1))) And t2 <> "" Then
                Call LogCasino("Jugador:" & Cuentas(UserIndex).mail & " intento entrar como jugador desde cliente con hash." & "Ip: " & UserList(UserIndex).ip)
                Call CloseSocket(UserIndex)
                Exit Sub
                End If
           
                'PLUTO:6.0a
                Cuentas(UserIndex).Naci = val(T4)
                
                For X = 1 To Cuentas(UserIndex).NumPjs
                    If Cuentas(UserIndex).Pj(X) = t1 Then
                        'pluto:6.7
                        Call ConnectUser(UserIndex, t1, "", T3, T5)
                        Exit Sub
                    End If
                Next X
                Call SendData2(ToIndex, UserIndex, 0, 43, "Este jugador no pertenece a tu cuenta.")
                Exit Sub
        End Select
   End If
   
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Si no esta logeado y envia un comando diferente a los
'de arriba cerramos la conexion.
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'pluto:2.13
If Not UserList(UserIndex).flags.UserLogged Then 'Or Not Cuentas(UserIndex).Logged = True Then

   'Call LogError("Mesaje enviado sin logearse:" & rdata)
   ' Call CloseUser(UserIndex)
   Call CloseSocket(UserIndex)
    Exit Sub
End If
 


 'PLUTO 2.24 distribución de los TCP
 If UserList(UserIndex).flags.Privilegios > 0 Then
 Call TCP3.TCP3(UserIndex, rdata)
 End If
 
 If Left(rdata, 1) = "/" Then
 Call TCP2.TCP2(UserIndex, rdata)
 Exit Sub
 End If
 
 Call TCP1.TCP1(UserIndex, rdata)

Exit Sub
ErrorHandler: 'pluto:6.9
Call LogError("Error en handledata. Nombre:" & UserList(UserIndex).Name & " Ip: " & UserList(UserIndex).ip & " HD: " & UserList(UserIndex).Serie & " Datos: " & rdata & " Desc: " & Err.number & ": " & Err.Description)

End Sub

Sub ReloadSokcet()
Debug.Print "ReloadSocket"

On Error GoTo errhandler
#If UsarQueSocket = 1 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState " & Err.number & ": " & Err.Description)

End Sub

Sub ActualizarHechizos(UserIndex As Integer)
On Error GoTo fallo
    Dim X As Integer
    For X = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(X) <> 0 Then
            Call SendData2(ToIndex, UserIndex, 0, 34, X & "," & UserList(UserIndex).Stats.UserHechizos(X) & "," & Hechizos(UserList(UserIndex).Stats.UserHechizos(X)).Nombre)
        Else
            Call SendData2(ToIndex, UserIndex, 0, 34, X & ",0,(None)")
        End If
    Next

Exit Sub
fallo:
Call LogError("actualizarhechizos " & Err.number & " D: " & Err.Description)
End Sub
