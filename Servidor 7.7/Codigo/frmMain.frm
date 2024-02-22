VERSION 5.00
Object = "{3007DCB1-D1F6-4A56-873C-0798895C1EC9}#1.0#0"; "CtlServidorTCP.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server AodraG v.6.2"
   ClientHeight    =   3645
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   7140
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3645
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer tesorocaballero 
      Interval        =   60000
      Left            =   4440
      Top             =   2640
   End
   Begin VB.Timer timermomia 
      Interval        =   60000
      Left            =   4440
      Top             =   2160
   End
   Begin VB.Timer Limpiado 
      Interval        =   20000
      Left            =   6360
      Top             =   3240
   End
   Begin VB.Timer Torneo 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5880
      Top             =   3240
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws_server 
      Left            =   2760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Auditoria 
      Interval        =   1000
      Left            =   5400
      Top             =   3240
   End
   Begin VB.Timer ContadorBytes 
      Interval        =   1000
      Left            =   4920
      Top             =   3240
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1080
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   360
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.ListBox lstLog 
      Height          =   900
      Left            =   1920
      TabIndex        =   10
      Top             =   2400
      Width           =   2535
   End
   Begin CTLSERVIDORTCPLib.CtlServidorTCP TCPServ 
      Left            =   1080
      Top             =   1800
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "Log (No Marcar)"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5880
      Top             =   2160
   End
   Begin VB.Timer CmdExec 
      Interval        =   10
      Left            =   6360
      Top             =   2880
   End
   Begin VB.Timer GameTimer 
      Interval        =   40
      Left            =   5880
      Top             =   2520
   End
   Begin VB.Timer tPiqueteC 
      Interval        =   6000
      Left            =   5400
      Top             =   2160
   End
   Begin VB.Timer tTraficStat 
      Interval        =   6000
      Left            =   5400
      Top             =   2520
   End
   Begin VB.Timer tLluviaEvent 
      Interval        =   60000
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Timer tLluvia 
      Interval        =   500
      Left            =   6360
      Top             =   2160
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6360
      Top             =   2520
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   6855
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Timer FX 
      Interval        =   4000
      Left            =   4920
      Top             =   2160
   End
   Begin VB.Timer npcataca 
      Interval        =   4000
      Left            =   4920
      Top             =   2520
   End
   Begin VB.Timer KillLog 
      Interval        =   60000
      Left            =   5400
      Top             =   2880
   End
   Begin VB.Timer TIMER_AI 
      Interval        =   100
      Left            =   5880
      Top             =   2880
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000003&
      Caption         =   "Minutos Online:"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000001&
      Caption         =   "Trafico Entrante:"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000011&
      Caption         =   "Trafico Saliente:"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCierra 
         Caption         =   "&Cerrar Servidor sin perder datos."
      End
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pluto:6.2
'*        Servidor                                                     *
'*        =======                                                      *
'*      CODE BY MaLkAvIaN_NeT  jav_025@hotmail.com,xmalkavianx@gmail.com*
'*           Comunidad de programación de proyectos 'OPEN SOURCE':
'*
'*--         http://groups.msn.com/SencicoNetOnlinepromo007/           *
'***********************************************************************
'***********************************************************************

Public str_ruta As String, str_archivo_temporal As String
Dim lng_tamaño_archivo As Long


'--------------
    'PLUTO:2.15
    Dim aa As Long
    Dim bb As Long
    '-----------------
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
Dim iuserindex As Integer

For iuserindex = 1 To MaxUsers
   
   'Conexion activa? pluto:2.12 añado user logged y privilegios
   If UserList(iuserindex).ConnID <> -1 And UserList(iuserindex).flags.UserLogged And UserList(iuserindex).flags.Privilegios = 0 Then
        'Actualiza el contador de inactividad
        
         UserList(iuserindex).Counters.IdleCount = UserList(iuserindex).Counters.IdleCount + 1
        If UserList(iuserindex).Counters.IdleCount >= IdleLimit Then
            'Call SendData(ToIndex, iuserindex, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
            Call CloseSocket(iuserindex)
        End If
  End If
  
Next iuserindex

End Sub

 Public Sub EnviarCorreo(ByVal Clave As String, ByVal Nombre As String)
On Error GoTo fuera
 'pluto:2.24--------------------
Dim caa As Byte
caa = val(GetVar(App.Path & "\DraG-Email\Emails.txt", "CORREOS", "Numero"))
Call WriteVar(App.Path & "\DraG-Email\Emails.txt", "CORREOS", "Numero", caa + 1)
Call WriteVar(App.Path & "\DraG-Email\Emails.txt", "CORREOS", "E" & caa + 1, " " & Nombre)
Call WriteVar(App.Path & "\DraG-Email\Emails.txt", "CORREOS", "C" & caa + 1, " " & Clave)
Exit Sub
'-------------------------------
fuera:
 Call LogError("Error EnviarCorreo")
End Sub


Private Sub Auditoria_Timer()
Static segundos As Byte

'segundos = segundos + 1
'If segundos > 60 Then segundos = 1
'Call SendData(ToAll, 0, 0, "|| Segundo: " & segundos & "´" & FontTypeNames.FONTTYPE_info)

'Dim k As Integer
'For k = 1 To LastUser
   ' If UserList(k).ConnID <> -1 Then
      '  DayStats.Segundos = DayStats.Segundos + 1
    'End If
'Next k

'Call PasarSegundo

End Sub

Private Sub AutoSave_Timer()

On Error GoTo errhandler
'fired every minute
'pluto:2.14
'Static MinutosPoder As Byte
 Dim tt As Byte
 Dim rn As Integer
 Dim ii As Integer
Static Minutos As Long
Static MinutosLatsClean As Long

Static MinsSocketReset As Long
Static MinutosNumUsersCheck As Long
Dim ero As Byte
Dim num As Long

MinsRunning = MinsRunning + 1
'Call SendData(ToAll, 0, 0, "|| Ahora!! " & "´" & FontTypeNames.FONTTYPE_info)
Dim i As Integer
    
 'se edito tema de la carcel acalele
    
'For i = 1 To MaxUsers
    'If UserList(i).ShTime > 19 Then
        'If UserList(i).Counters.Pena = 0 Then
            'Call Encarcelar(i, 30)
            'Call SendData(ToIndex, i, 0, "||Has sido encarcelado, deberas permanecer en la carcel 30 minutos por uso de speed hack." & "´" & FontTypeNames.FONTTYPE_info)
        'End If
    'End If
    'UserList(i).ShTime = 0
'Next i


If MinsRunning = 60 Then
  'pluto:2.15
  'yaya = 0
    
    Horas = Horas + 1
    If Horas = 24 Then
       ' Call SaveDayStats
        DayStats.Maxusuarios = 0
        DayStats.segundos = 0
        DayStats.Promedio = 0
        'pluto:2.4.1 quitar dias eleccion lider
        'Call DayElapsed
        Dias = Dias + 1
        Horas = 0
    End If
    MinsRunning = 0
End If

'pluto:2.23----------------------
'Call ModAreas.AreasOptimizacion
'-------------------------------
Minutos = Minutos + 1


'EstadoMinotauro = 1
'Minotauro = "MaChoTe"
'MinutosMinotauro = 2
'PLUTO:6.0a
If EstadoMinotauro = 1 Then
MinutosMinotauro = MinutosMinotauro - 1
'MinutosMinotauro = 4
End If

'pluto:2.17
If castillo1 <> "" Then MinutosCastilloNorte = MinutosCastilloNorte + 1
If castillo2 <> "" Then MinutosCastilloSur = MinutosCastilloSur + 1
If castillo3 <> "" Then MinutosCastilloEste = MinutosCastilloEste + 1
If castillo4 <> "" Then MinutosCastilloOeste = MinutosCastilloOeste + 1
If fortaleza <> "" Then MinutosFortaleza = MinutosFortaleza + 1
ero = 1
'-----------------------------
'pluto:2.14
If DobleExp > 0 Then
DobleExp = DobleExp - 1
Call SendData2(ToAll, 0, 0, 117, "B@" & DobleExp)
Else
MsgEntra = ""
DobleExp = 0
End If

'pluto:2.12
If MapInfo(MapaTorneo2).NumUsers > 1 Then MinutoSinMorir = MinutoSinMorir + 1
If MinutoSinMorir > 4 Then MinutoSinMorir = 0
If MinutoSinMorir = 3 Then
Call SendData(ToMap, 0, 194, "||Maten a sus rivales o serán todos expulsados de la sala." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
End If
'pluto:2.4.5
MinutosOnline = MinutosOnline + 1
frmMain.Label4.Caption = "Minutos Online: " & MinutosOnline
'nati: agrego esto para los usuarios online web
Call WriteVar(App.Path & "\web.txt", "ESP", "Online", str$(NumUsers))
Call WriteVar(App.Path & "\php.txt", "INIT", "Online", str$(NumUsers))
'pluto:2.24-----------
Call BDDSetUsersOnline
'If ServerPrimario = 1 Then

'Call WriteVar(IniPath & "php.txt", "INIT", "Online", str(NumUsers))
'Call WriteVar(IniPath & "php.txt", "INIT", "RecordHoy", str(ReNumUsers))
'Call WriteVar(IniPath & "php.txt", "INIT", "RecordAyer", str(AyerReNumUsers))
'Else
'Call WriteVar(IniPath & "php.txt", "INIT", "Online2", str(NumUsers))
'Call WriteVar(IniPath & "php.txt", "INIT", "RecordHoy2", str(ReNumUsers))
'Call WriteVar(IniPath & "php.txt", "INIT", "RecordAyer2", str(AyerReNumUsers))
'End If
'---------------------
ero = 2

MinutosNumUsersCheck = MinutosNumUsersCheck + 1
If MinutosNumUsersCheck >= 2 Then
    MinutosNumUsersCheck = 0
    num = 0
'pluto:6.2
CodigoMacro = RandomNumber(100, 999)
'---------------------------------------------------------------------
'-------Todos los usuarios cada 2 minutos---pluto:6.0A----------------
'---------------------------------------------------------------------
'Dim TodosUser As Integer
'TodosUser = LastUser
    
    For i = 1 To MaxUsers
 'users no logueados cada 2 minutos comprobamos
    If UserList(i).ConnID <> -1 And UserList(i).flags.UserLogged = False And Cuentas(i).Logged = False Then
    Call CloseSocket(i, True) 'añado true

        Dim Tindex As Integer
        Tindex = NameIndex("AoDraGBoT")
        If Tindex > 0 Then
        Call SendData(ToIndex, Tindex, 0, "|| C.Slot: " & i & "´" & FontTypeNames.FONTTYPE_TALK)
        End If

    End If


'users logueados cada 2 minutos
    If UserList(i).ConnID <> -1 And UserList(i).flags.UserLogged Then
num = num + 1

tt = RandomNumber(1, NumUsers)
Dim cct As Integer
cct = RandomNumber(1, 200)
'pluto:6.8
If cct = i And EventoDia = 3 And UserList(i).Pos.Map <> 165 Then
Call SpawnNpc(718, UserList(i).Pos, True, False)
End If

'For i = 1 To LastUser
'pluto:2.11.0
'If UserList(i).ConnID <> -1 Then
      '¿User valido?
        'If UserList(i).flags.UserLogged Then
ero = 3
'pluto:6.0A
UserList(i).flags.MinutosOnline = UserList(i).flags.MinutosOnline + 2

If Minotauro = UserList(i).Name Then
Call SendData(ToIndex, i, 0, "|| Debes matar al Minotauro antes de " & MinutosMinotauro & " minutos." & "´" & FontTypeNames.FONTTYPE_GUILD)
End If

'pluto:5.2------------
UserList(i).flags.CMuerte = 0
'---------------------
If (UserList(i).Pos.Map = 171 Or UserList(i).Pos.Map = 177) Then UserList(i).ObjetosTirados = 0
'pluto:2.12 añado alarma=1 para evitar la del /log
If UserList(i).Alarma = 1 Then UserList(i).Alarma = 0
UserList(i).MuertesTime = 0
If UserList(i).flags.Muerto = 0 And UserList(i).ObjetosTirados > 15 And UserList(i).ObjetosTirados < 26 Then
Call Encarcelar(i, 10)
Call LogCasino("/CARCEL AUTOMATICO " & UserList(i).Name & " IP:" & UserList(i).ip)
End If

UserList(i).ObjetosTirados = 0

'pluto:2.14
If UserList(i).Counters.Pena > 0 Then
Call SendData(ToIndex, i, 0, "||Te quedan " & UserList(i).Counters.Pena & " Minutos de Cárcel." & "´" & FontTypeNames.FONTTYPE_info)
End If
'pluto:6.8
If UserList(i).Alarma = 2 Then
Call SendData(ToGM, i, 0, "||LoG activo sobre " & UserList(i).Name & "´" & FontTypeNames.FONTTYPE_TALK)
End If


'pluto:2.17

'If UserList(i).flags.Navegando > 0 And ((UserList(i).Stats.UserSkills(Navegacion) >= 40 And ModNavegacion(UserList(i).clase) = 2.3) Or (UserList(i).Stats.UserSkills(Navegacion) >= 20 And ModNavegacion(UserList(i).clase) <> 2.3)) Then Call SubirSkill(i, Navegacion)
 
 '--------


Dim aa As Integer
aa = RandomNumber(1, 360)
'pluto:6.2
If UserList(i).flags.Macreanda > 0 And aa > 300 Then
'COMPROBANDOMACRO = True
UserList(i).flags.ComproMacro = 15
'Call SendData(ToIndex, i, 0, "O4")
'UserList(i).flags.ComproMacro = 1
End If


'Dim tindex As Integer
'TimeEmbarazo = 2
'TimeAborto = 7
'ProbEmbarazo = 35
'pluto:2.17 añade max 5 hijos

'pluto:6.8 añade marido max 5 hijos
If UserList(i).Esposa = "" Then GoTo fuera3
Tindex = NameIndex(UserList(i).Esposa)
If Tindex < 1 Then GoTo fuera3

'quitar testeo 1 por 95
If aa > ProbEmbarazo And UserList(i).Amor > 1 And UserList(i).Embarazada = 0 And UCase$(UserList(i).Genero) = "MUJER" And UserList(i).Nhijos < 5 And UserList(Tindex).Nhijos < 5 Then
UserList(i).Embarazada = 95
Call SendData(ToIndex, i, 0, "||Sientes que una vida nueva crece en tu interior. Enhorabuena! estás embarazada." & "´" & FontTypeNames.FONTTYPE_TALK)
End If
'pluto:2.15
'UserList(i).Embarazada = 1
ero = 4



If aa > 250 Then
'If UserList(i).Esposa = "" Then GoTo fuera2
'Tindex = NameIndex(UserList(i).Esposa)
'If Tindex < 1 Then GoTo fuera2

If Distancia(UserList(i).Pos, UserList(Tindex).Pos) < 8 Then
UserList(i).Amor = UserList(i).Amor + 1
If UserList(i).Amor > 100 Then UserList(i).Amor = 100
Call SendData2(ToPCArea, i, UserList(i).Pos.Map, 22, UserList(i).Char.CharIndex & "," & 88 & "," & 5)
Else
If UserList(i).Amor > 0 Then UserList(i).Amor = UserList(i).Amor - 1
If UserList(i).Amor < 0 Then UserList(i).Amor = 0
ero = 6
End If 'distancia

End If 'aa


fuera3:
ero = 5
'pluto.6.3 +2 embarazo
If UserList(i).Embarazada > 0 Then UserList(i).Embarazada = UserList(i).Embarazada + 2

If UserList(i).Embarazada >= TimeAborto Then
UserList(i).Embarazada = 0
Call SendData(ToIndex, i, 0, "||Has perdido al bebé!!" & "´" & FontTypeNames.FONTTYPE_TALK)
End If

If UserList(i).Embarazada >= TimeEmbarazo And UserList(i).NombreDelBebe = "" Then
Call SendData(ToIndex, i, 0, "||Estás a punto de tener un hijo, ve cuanto antes a una Matrona." & "´" & FontTypeNames.FONTTYPE_TALK)
End If
'----------





If UserList(i).GranPoder > 0 Then
    'pluto:6.2 Añadimos minas fortaleza y salas clan
    If UserList(i).Pos.Map = 186 Or MapInfo(UserList(i).Pos.Map).Pk = False Or MapInfo(UserList(i).Pos.Map).Zona = "CLAN" Then
        MinutosPoder = MinutosPoder + 1
        If MinutosPoder > 3 Then
        UserList(i).GranPoder = 0
        UserGranPoder = ""
        UserList(i).Char.FX = 0
        Call SendData2(ToMap, i, UserList(i).Pos.Map, 22, UserList(i).Char.CharIndex & "," & 68 & "," & 0)
        MinutosPoder = 0
        Else
        Call SendData(ToIndex, i, 0, "|| Estás en zona segura, te quedan " & 4 - MinutosPoder & " minutos para perder el Gran Poder." & "´" & FontTypeNames.FONTTYPE_GUILD)
        End If
    End If
End If 'granpoder>0
ero = 7

'pluto:2.11 Quitar esto
If UserGranPoder = "" And i = tt And UserList(i).flags.Muerto = 0 And UserList(i).flags.Privilegios = 0 And NumUsers > NumeroGranPoder Then
UserGranPoder = UserList(i).Name
UserList(i).GranPoder = 1
End If

fuera:
'End If 'connID
'End If  'logged
'Next
End If

'-------------------------------------
'comprobamos exp de castillos
'-------------------------------------
'pluto.2.17 norte
If MinutosCastilloNorte > 59 Then

'pluto:2.4.1 exp por rangos, cambio 20000 por rn

'For i = 1 To LastUser
If UserList(i).GuildInfo.GuildName = castillo1 And UserList(i).Stats.ELV > 25 Then
rn = Int(UserList(i).GuildInfo.GuildPoints * 2) + 15000
UserList(i).Stats.exp = UserList(i).Stats.exp + rn
'UserList(i).Stats.GLD = UserList(i).Stats.GLD + Int(rn / 5)
Call AddtoVar(UserList(i).Stats.GLD, Int(rn / 5), MAXORO)
'pluto:2.4.5

Call SendData(ToIndex, i, 0, "|| Has obtenido " & rn & " de Exp y " & Int(rn / 5) & " de Oro por Mantener el Castillo Norte" & "´" & FontTypeNames.FONTTYPE_GUILD)
Call SendData(ToIndex, i, 0, "TW" & 105)
Call SendUserStatsOro(i)
Call CheckUserLevel(i)
End If
'Next i

End If 'norte


ero = 10
'pluto.2.17 SUR
If MinutosCastilloSur > 59 Then

'Dim rn As Integer
'For i = 1 To LastUser
If UserList(i).GuildInfo.GuildName = castillo2 And UserList(i).Stats.ELV > 25 Then
rn = Int(UserList(i).GuildInfo.GuildPoints * 2) + 15000
UserList(i).Stats.exp = UserList(i).Stats.exp + rn
'UserList(i).Stats.GLD = UserList(i).Stats.GLD + Int(rn / 5)
Call AddtoVar(UserList(i).Stats.GLD, Int(rn / 5), MAXORO)
Call SendData(ToIndex, i, 0, "|| Has obtenido " & rn & " de Exp y " & Int(rn / 5) & " de Oro por Mantener el Castillo Sur" & "´" & FontTypeNames.FONTTYPE_GUILD)
Call SendData(ToIndex, i, 0, "TW" & 105)
Call SendUserStatsOro(i)
Call CheckUserLevel(i)
End If
'Next i

End If 'Sur
ero = 11
'pluto.2.17 ESTE
If MinutosCastilloEste > 59 Then



'For i = 1 To LastUser
If UserList(i).GuildInfo.GuildName = castillo3 And UserList(i).Stats.ELV > 25 Then
rn = Int(UserList(i).GuildInfo.GuildPoints * 2) + 15000
UserList(i).Stats.exp = UserList(i).Stats.exp + rn
'UserList(i).Stats.GLD = UserList(i).Stats.GLD + Int(rn / 5)
Call AddtoVar(UserList(i).Stats.GLD, Int(rn / 5), MAXORO)

Call SendData(ToIndex, i, 0, "|| Has obtenido " & rn & " de Exp y " & Int(rn / 5) & " de Oro por Mantener el Castillo Este" & "´" & FontTypeNames.FONTTYPE_GUILD)
Call SendData(ToIndex, i, 0, "TW" & 105)
Call SendUserStatsOro(i)
Call CheckUserLevel(i)
End If
'Next i

End If 'este

'pluto.2.17 oESTE
If MinutosCastilloOeste > 59 Then
'suma puntos clan

'Dim rn As Integer
'For i = 1 To LastUser
If UserList(i).GuildInfo.GuildName = castillo4 And UserList(i).Stats.ELV > 25 Then
rn = Int(UserList(i).GuildInfo.GuildPoints * 2) + 15000
UserList(i).Stats.exp = UserList(i).Stats.exp + rn
'UserList(i).Stats.GLD = UserList(i).Stats.GLD + Int(rn / 5)
Call AddtoVar(UserList(i).Stats.GLD, Int(rn / 5), MAXORO)
Call SendData(ToIndex, i, 0, "|| Has obtenido " & rn & " de Exp y " & Int(rn / 5) & " de Oro por Mantener el Castillo Oeste" & "´" & FontTypeNames.FONTTYPE_GUILD)
Call SendData(ToIndex, i, 0, "TW" & 105)
Call SendUserStatsOro(i)
Call CheckUserLevel(i)
End If
'Next i

End If 'oeste

'pluto.2.17 fortale
If MinutosFortaleza > 59 Then
'suma puntos clan

ero = 12
'Dim rn As Integer
'For i = 1 To LastUser
If UserList(i).GuildInfo.GuildName = fortaleza And UserList(i).Stats.ELV > 25 Then
rn = Int(UserList(i).GuildInfo.GuildPoints * 2) + 15000
UserList(i).Stats.exp = UserList(i).Stats.exp + rn
'UserList(i).Stats.GLD = UserList(i).Stats.GLD + Int(rn / 2)
Call AddtoVar(UserList(i).Stats.GLD, Int(rn / 2), MAXORO)
Call SendData(ToIndex, i, 0, "|| Has obtenido " & rn & " de Exp y " & Int(rn / 2) & " de Oro por Mantener la Fortaleza" & "´" & FontTypeNames.FONTTYPE_GUILD)
Call SendData(ToIndex, i, 0, "TW" & 105)
Call SendUserStatsOro(i)
Call CheckUserLevel(i)
End If
'Next i

End If 'fortaleza

'fin exp castillos

'fin users logueados cada 2 minutos
Next i
'------FIN todos los usuarios cada dos minutos------------
    
    If num <> NumUsers Then
        NumUsers = num
        'Call SendData(ToAdmins, 0, 0, "Servidor> Error en NumUsers. Contactar a algun Programador." & FONTTYPE_SERVER)
        Call LogCriticEvent("Num <> NumUsers")
    End If

End If 'minutoscheck = 2 -----------------------------------------------

ero = 8

'nati: modifico esto, Nuevo sistema de puntuaje de recompensas.
'pluto:6.0A---- SUMAMOS PUNTOS CLANES
Dim OnlineClanRecompensa As Integer
Dim LoopUserClan As Integer
OnlineClanRecompensa = 0
If MinutosCastilloNorte > 59 Then
For ii = 1 To Guilds.Count
If UCase$(castillo1) = UCase$(Guilds(ii).GuildName) Then
'Guilds(ii).Reputation = Guilds(ii).Reputation + 5
'nati:Calcula la gente online en el Clan
For LoopUserClan = 1 To LastUser
If UCase$(Guilds(ii).GuildName) = UCase$(UserList(LoopUserClan).GuildInfo.GuildName) Then
OnlineClanRecompensa = OnlineClanRecompensa + 1
End If
Next
'nati: Calcula la gente online en el Clan
Guilds(ii).Reputation = Guilds(ii).Reputation + (NumUsers - OnlineClanRecompensa / 2) + 10
End If
'>> ANTIGUO >> If UCase$(fortaleza) = UCase$(Guilds(ii).GuildName) Then Guilds(ii).Reputation = Guilds(ii).Reputation + 5
Next ii
MinutosCastilloNorte = 0
End If

If MinutosCastilloSur > 59 Then
For ii = 1 To Guilds.Count
If UCase$(castillo2) = UCase$(Guilds(ii).GuildName) Then
'Guilds(ii).Reputation = Guilds(ii).Reputation + 5
'nati:Calcula la gente online en el Clan
For LoopUserClan = 1 To LastUser
If UCase$(Guilds(ii).GuildName) = UCase$(UserList(LoopUserClan).GuildInfo.GuildName) Then
OnlineClanRecompensa = OnlineClanRecompensa + 1
End If
Next
'nati: Calcula la gente online en el Clan
Guilds(ii).Reputation = Guilds(ii).Reputation + (NumUsers - OnlineClanRecompensa / 2) + 10
End If
'>> ANTIGUO >> If UCase$(fortaleza) = UCase$(Guilds(ii).GuildName) Then Guilds(ii).Reputation = Guilds(ii).Reputation + 5
Next ii
MinutosCastilloSur = 0
End If

If MinutosCastilloEste > 59 Then
For ii = 1 To Guilds.Count
If UCase$(castillo3) = UCase$(Guilds(ii).GuildName) Then
'Guilds(ii).Reputation = Guilds(ii).Reputation + 5
'nati:Calcula la gente online en el Clan
For LoopUserClan = 1 To LastUser
If UCase$(Guilds(ii).GuildName) = UCase$(UserList(LoopUserClan).GuildInfo.GuildName) Then
OnlineClanRecompensa = OnlineClanRecompensa + 1
End If
Next
'nati: Calcula la gente online en el Clan
Guilds(ii).Reputation = Guilds(ii).Reputation + (NumUsers - OnlineClanRecompensa / 2) + 10
End If
'>> ANTIGUO >> If UCase$(fortaleza) = UCase$(Guilds(ii).GuildName) Then Guilds(ii).Reputation = Guilds(ii).Reputation + 5
Next ii
MinutosCastilloEste = 0
End If

If MinutosCastilloOeste > 59 Then
For ii = 1 To Guilds.Count
If UCase$(castillo1) = UCase$(Guilds(ii).GuildName) Then
'Guilds(ii).Reputation = Guilds(ii).Reputation + 5
'nati:Calcula la gente online en el Clan
For LoopUserClan = 1 To LastUser
If UCase$(Guilds(ii).GuildName) = UCase$(UserList(LoopUserClan).GuildInfo.GuildName) Then
OnlineClanRecompensa = OnlineClanRecompensa + 1
End If
Next
'nati: Calcula la gente online en el Clan
Guilds(ii).Reputation = Guilds(ii).Reputation + (NumUsers - OnlineClanRecompensa / 2) + 10
End If
If UCase$(castillo4) = UCase$(Guilds(ii).GuildName) Then Guilds(ii).Reputation = Guilds(ii).Reputation + 5
'>> ANTIGUO >> If UCase$(fortaleza) = UCase$(Guilds(ii).GuildName) Then Guilds(ii).Reputation = Guilds(ii).Reputation + 5
Next ii
MinutosCastilloOeste = 0
End If

If MinutosFortaleza > 59 Then
For ii = 1 To Guilds.Count
If UCase$(castillo1) = UCase$(Guilds(ii).GuildName) Then
'Guilds(ii).Reputation = Guilds(ii).Reputation + 5
'nati:Calcula la gente online en el Clan
For LoopUserClan = 1 To LastUser
If UCase$(Guilds(ii).GuildName) = UCase$(UserList(LoopUserClan).GuildInfo.GuildName) Then
OnlineClanRecompensa = OnlineClanRecompensa + 1
End If
Next
'nati: Calcula la gente online en el Clan
Guilds(ii).Reputation = Guilds(ii).Reputation + (NumUsers - OnlineClanRecompensa / 2) + 10
End If
'>> ANTIGUO >> If UCase$(fortaleza) = UCase$(Guilds(ii).GuildName) Then Guilds(ii).Reputation = Guilds(ii).Reputation + 5
Next ii
MinutosFortaleza = 0
End If
'-------------------FIN SUMA PUNTOS CLAN----------

If Minutos = 60 Or Minutos = 120 Or Minutos = 180 Then
'pluto:2.4.7 --> Cambia pregunta trivial cada hora.
Call Loadtrivial
'pluto:6.8 añado evento
If aa > 330 And Caballero = False And EventoDia <> 2 Then
Dim Caba As WorldPos
Caba.X = 50
Caba.Y = 50
Caba.Map = 250
Call SpawnNpc(633, Caba, False, True)
Caballero = True
End If
ero = 9
End If '=60 or =120... mintuos
'-------------------------------------
'EstadoMinotauro = 0

If Minutos = 5 Then
    'pluto:6.5 reponemos raids-----------
    Dim nxx As Byte
    
    For nxx = 1 To 6 'cambiamos a 6
    If RaidVivos(nxx).Activo = 0 Then Call RespawnRaids(nxx)
    Next
    '----------------------------------
    If EstadoMinotauro = 2 Then
    If RandomNumber(1, 100) > 80 Then EstadoMinotauro = 0
    End If
End If


'end if
ero = 13

If Minutos Mod 40 = 0 Then
    Call grabaPJ
End If

If Minutos >= MinutosWs Then
    Call DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
    'pluto:6.9
    frmMain.Limpiado.Enabled = True
    
End If
ero = 14
'pluto:2.11
'UserGranPoder = ""
If (MinutosLatsClean = 7 Or MinutosLatsClean = 15) And UserGranPoder <> "" Then
Dim Podercito As Integer
Podercito = NameIndex(UserGranPoder)
    If Podercito > 0 Then
Call SendData(ToAll, 0, 0, "|| Gran Poder: " & UserGranPoder & " en Mapa " & UserList(Podercito).Pos.Map & "´" & FontTypeNames.FONTTYPE_COMERCIO)
    End If
End If
ero = 15
'pluto:2.15
If MinutosLatsClean = 15 Then
MediaVez = MediaVez + 1
MediaUser = MediaUser + NumUsers
MediaUsers = Round((MediaUser / MediaVez), 1)
End If

ero = 16

If MinutosLatsClean >= 15 Then

        MinutosLatsClean = 0
        'pluto:6.0a quitamos el respawn cada 15, ahora solo con cada worlsave
        'Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Call LimpiarMundo
Else
        MinutosLatsClean = MinutosLatsClean + 1
End If

Call PurgarPenas
Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim n As Integer
n = FreeFile(1)
Open App.Path & "\logs\numusers.log" For Output Shared As n
Print #n, NumUsers
Close #n
'<<<<<-------- Log the number of users online ------>>>
ero = 17
Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave" & ero & " " & i)

End Sub
'pluto fusion
Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub Cerrar_Timer()

End Sub

Private Sub CmdExec_Timer()
Dim i As Integer
Static n As Long
On Error Resume Next ':(((
n = n + 1
'pluto:2.12 igual a la z
For i = 1 To MaxUsers
'For i = 1 To LastUser

   If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
     If Not UserList(i).CommandsBuffer.IsEmpty Then
'pluto:6.7------
'UserList(MiDato).Counters.UserRecibe = 2
'UserList(i).Counters.UserRecibe = UserList(i).Counters.UserRecibe + 1
'If UserList(i).Counters.UserRecibe > 35 Then UserList(i).Counters.UserRecibe = 1
'---------------------
     Call HandleData(i, UserList(i).CommandsBuffer.Pop)
     End If
   If n >= 10 Then
            If UserList(i).ColaSalida.Count > 0 Then ' And UserList(i).SockPuedoEnviar Then
    #If UsarQueSocket = 1 Then
                Call IntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 0 Then
    '            Call WrchIntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 2 Then
    '            Call ServIntentarEnviarDatosEncolados(i)
    #ElseIf UsarQueSocket = 3 Then
        'NADA, el control deberia ocuparse de esto!!!
        'si la cola se llena, dispara un on close
    #End If
            End If
        End If
    End If
Next i

If n >= 10 Then
    n = 0
End If

Exit Sub
hayerror:

End Sub

Private Sub Command1_Click()
Call SendData(ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub




Private Sub ContadorBytes_Timer()
     
     'PLUTO.2.15---------------------------

    aa = Round(BytesEnviados / 1024, 3)
    bb = Round(BytesRecibidos / 1024, 3)
    Label2.Caption = "Trafico Saliente: " & aa & "kb/s"
    Label3.Caption = "Trafico Entrante: " & bb & "kb/s"
    'pluto:2.14
    TotalBytesEnviados = TotalBytesEnviados + aa
    TotalBytesRecibidos = TotalBytesRecibidos + bb

    BytesEnviados = 0
    BytesRecibidos = 0
End Sub

Private Sub Form_Load()
Dim ix As Integer
'ix = Inet1.OpenURL("http://www.juegosdrag.es/update/updateseguridad.txt")
'If ix <> 8 Then End

End Sub

'Private Sub Cuentas_Timer()
    'Call BDDCmpCuentas
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hwnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Cancel = 1
'Me.Hide
End Sub

Private Sub Form_Resize()
'If WindowState = vbMinimized Then Command2_Click
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi(frmMain.hwnd)
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If


Call DescargaNpcsDat

Dim loopc As Integer

For loopc = 1 To MaxUsers
    If UserList(loopc).ConnID <> -1 Then Call CloseSocket(loopc)
Next

'Log
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " server cerrado."
Close #n

End

End Sub




Private Sub FX_Timer()
On Error GoTo hayerror
Dim MapIndex As Integer
Dim n As Integer
For MapIndex = 1 To NumMaps
    Randomize
    If RandomNumber(1, 150) < 12 Then
        If MapInfo(MapIndex).NumUsers > 0 Then

                Select Case MapInfo(MapIndex).Terreno
                   'Bosque
                   Case Bosque
                        n = RandomNumber(1, 100)
                        Select Case MapInfo(MapIndex).Zona
                            Case Campo
                              If Not Lloviendo Then
                                If n < 25 And n >= 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE)
                                ElseIf n < 25 And n < 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE2)
                                ElseIf n >= 25 And n <= 30 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO)
                                ElseIf n >= 30 And n <= 35 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO2)
                                ElseIf n >= 40 And n <= 45 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE3)
                                End If
                               End If
                            Case Ciudad
                               If Not Lloviendo Then
                                If n < 30 And n >= 25 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE)
                                ElseIf n < 30 And n < 25 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE2)
                                ElseIf n >= 30 And n <= 35 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO)
                                ElseIf n >= 35 And n <= 40 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO2)
                                ElseIf n >= 40 And n <= 45 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE3)
                                End If
                               End If
                        End Select
                            'pluto:sonidos casa
                                Case Casa
                              
                                Dim n2 As Integer
                                n2 = RandomNumber(1, 300)
                                If n2 < 20 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_CASA1)
                                ElseIf n2 > 19 And n2 < 40 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_CASA2)
                                ElseIf n2 > 39 And n2 < 60 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_CASA3)
                                ElseIf n2 > 59 And n2 < 80 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_CASA4)
                                ElseIf n2 > 279 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_CASA5)
                                End If
                             
                             'pluto:2-3-04 sonido alcantarillas
                            Case ALCANTARILLA
                              n2 = RandomNumber(1, 300)
                                If n2 < 20 Then
                                Call SendData(ToMap, 0, MapIndex, "TW" & 137)
                                End If
                            
                            
                         'pluto:sonidos bosque terror
                                Case BOSQUETERROR
                              
                                n2 = RandomNumber(1, 300)
                                If n2 < 20 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & 107)
                                ElseIf n2 > 19 And n2 < 40 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & 114)
                                ElseIf n2 > 39 And n2 < 60 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & 116)
                                ElseIf n2 > 59 And n2 < 80 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & 117)
                                ElseIf n2 > 150 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & 116)
                    
                        End If
                End Select

        End If
    End If
Next
Exit Sub
hayerror:
End Sub
Private Sub GameTimer_Timer()
Dim iuserindex As Integer
Dim bEnviarStats As Boolean
Dim bEnviarAyS As Boolean
Dim iNpcIndex As Integer
Dim nvv As Byte
'Static lTirarBasura As Long
'Static lPermiteAtacar As Long
'Static lPermiteCast As Long
'Static lPermiteTrabajar As Long
'pluto:2.8.0
'Static lPermiteFlechas As Long
'pluto:2.10
'Static lPermiteTomar As Long

'If lPermiteTomar < IntervaloUserPuedeTomar Then
 '   lPermiteTomar = lPermiteTomar + 1
'End If
'--------------------------------
'If lPermiteAtacar < IntervaloUserPuedeAtacar Then
 '   lPermiteAtacar = lPermiteAtacar + 1
'End If

'If lPermiteCast < IntervaloUserPuedeCastear Then
 '   lPermiteCast = lPermiteCast + 1
'End If

'If lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
 '    lPermiteTrabajar = lPermiteTrabajar + 1
'End If
'pluto:2.8
'If lPermiteFlechas < IntervaloUserPuedeFlechas Then
     'lPermiteFlechas = lPermiteFlechas + 1
'End If


'[/Alejo]
On Error GoTo hayerror
 '<<<<<< Procesa eventos de los usuarios >>>>>>
 For iuserindex = 1 To MaxUsers
   'Conexion activa?
    If UserList(iuserindex).ConnID <> -1 Then
      '¿User valido?
        'pluto:6.0A
        If UserList(iuserindex).ConnIDValida And UserList(iuserindex).flags.UserLogged Then
  'pluto:6.0A
   If UserList(iuserindex).Name = "" Or Cuentas(iuserindex).Logged = False Then
   CloseUser (iuserindex)
   Call LogError("Quitado Clon: " & UserList(iuserindex).Pos.Map & "," & UserList(iuserindex).Pos.X & "," & UserList(iuserindex).Pos.Y)
   Exit Sub
   End If
   '---------
         bEnviarStats = False
         bEnviarAyS = False
        
         UserList(iuserindex).NumeroPaquetesPorMiliSec = 0

         '<<<<<<<<<<<< Allow attack >>>>>>>>>>>>>
         'If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
             '   UserList(iUserIndex).flags.PuedeAtacar = 1
         'End If
         '<<<<<<<<<<<< Allow attack >>>>>>>>>>>>>

         '<<<<<<<<<<<< Allow Cast spells >>>>>>>>>>>
         
         'If Not lPermiteCast < IntervaloUserPuedeCastear Then
            '  UserList(iUserIndex).flags.PuedeLanzarSpell = 1
         'End If
         '<<<<<<<<<<<< Allow Cast spells >>>>>>>>>>>

         '<<<<<<<<<<<< Allow Work >>>>>>>>>>>
         
         'If Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
             ' UserList(iUserIndex).flags.PuedeTrabajar = 1
         'End If
         '<<<<<<<<<<<< Allow Work >>>>>>>>>>>
'pluto:2.8.0
'If Not lPermiteFlechas < IntervaloUserPuedeFlechas Then
              'UserList(iUserIndex).flags.PuedeFlechas = 1
         'End If
'pluto:2.10
'If Not lPermiteTomar < IntervaloUserPuedeTomar Then
             ' UserList(iUserIndex).flags.PuedeTomar = 1
         'End If
'---------------------------------------------------
'pluto.2.3
'UserList(iUserIndex).Stats.PesoMax = (UserList(iUserIndex).Stats.UserAtributosBackUP(1) * 5) + (UserList(iUserIndex).Flags.ClaseMontura * 100)

'pluto:2.4
'If UserList(iUserIndex).flags.Privilegios > 0 Then GoTo alli
'If UserList(iUserIndex).Stats.GLD + UserList(iUserIndex).Stats.Banco > MoroOn Then
'MoroOn = UserList(iUserIndex).Stats.GLD + UserList(iUserIndex).Stats.Banco
'NMoroOn = UserList(iUserIndex).Name
'pluto:2.17
'If MoroOn > Moro Then
'Moro = MoroOn
'NMoro = NMoroOn
'End If
'-------------
'End If

'If UserList(iUserIndex).Remort = 0 Then
'nvv = UserList(iUserIndex).Stats.ELV
'Else
'nvv = UserList(iUserIndex).Stats.ELV + 55
'End If

'If Not Criminal(iUserIndex) And nvv > NivCiuON Then NivCiuON = nvv: NNivCiuON = UserList(iUserIndex).Name
'If Criminal(iUserIndex) And nvv > NivCrimiON Then NivCrimiON = nvv: NNivCrimiON = UserList(iUserIndex).Name
'alli:
            'pluto:6.5---quitamos el dotileevents por controlar salidas y vigilareventos
            
         'Call ControlaSalidas(iuserindex, UserList(iuserindex).Pos.Map, UserList(iuserindex).Pos.X, UserList(iuserindex).Pos.Y)
                
If CuentaRegresiva Then
If CuentaRegresiva = 200 Then
Call SendData(ToMap, 0, UserList(indexCuentaRegresiva).Pos.Map, "||Cuenta regresiva: 3 ´" & FontTypeNames.FONTTYPE_TALK)
End If
If CuentaRegresiva = 100 Then
Call SendData(ToMap, 0, UserList(indexCuentaRegresiva).Pos.Map, "||Cuenta regresiva: 2 ´" & FontTypeNames.FONTTYPE_TALK)
End If
If CuentaRegresiva = 50 Then
Call SendData(ToMap, 0, UserList(indexCuentaRegresiva).Pos.Map, "||Cuenta regresiva: 1 ´" & FontTypeNames.FONTTYPE_TALK)
End If
If CuentaRegresiva = 1 Then
MapaSeguro = UserList(indexCuentaRegresiva).Pos.Map
Call SendData(ToMap, 0, UserList(indexCuentaRegresiva).Pos.Map, "||Mapa Inseguro ´" & FontTypeNames.FONTTYPE_TALK)
End If
CuentaRegresiva = CuentaRegresiva - 1
End If
                
If UserList(iuserindex).flags.Muerto = 0 Then
         
          'pluto:6.2
         If UserList(iuserindex).flags.Macreanda > 0 And UserList(iuserindex).Stats.MinSta > 0 Then Call EfectoMacrear(iuserindex, UserList(iuserindex).flags.Macreanda)
                 
         If UserList(iuserindex).flags.Paralizado = 1 Then Call EfectoParalisisUser(iuserindex)
         'pluto:2.15
         If UserList(iuserindex).flags.Protec > 0 Then Call EfectoProtec(iuserindex)
         'nati:Ron
         If UserList(iuserindex).flags.Ron > 0 Then Call EfectoRon(iuserindex)
         
         If UserList(iuserindex).flags.Ceguera = 1 Or _
            UserList(iuserindex).flags.Estupidez Then Call EfectoCegueEstu(iuserindex)
          'pluto:2.14 torneo2
          'If UserList(iuserindex).flags.Morph > 0 And UserList(iuserindex).Pos.Map <> MapaTorneo2 Then Call EfectoMorphUser(iuserindex)
If UserList(iuserindex).flags.Morph > 0 Then Call EfectoMorphUser(iuserindex)

         
               If UserList(iuserindex).flags.Angel Or UserList(iuserindex).flags.Demonio Then Call QuitarSta(iuserindex, 2)
               'pluto:2.9.0
               If UserList(iuserindex).Char.Body = 9 Or UserList(iuserindex).Char.Body = 260 Or UserList(iuserindex).Char.Body = 380 Then
               Call QuitarSta(iuserindex, 2)
               If UserList(iuserindex).Stats.MinSta < 1 Then UserList(iuserindex).Counters.Morph = 0
               End If
               If UserList(iuserindex).Char.Body = 214 Then
               Call QuitarSta(iuserindex, 5)
               If UserList(iuserindex).Stats.MinSta < 1 Then UserList(iuserindex).Counters.Morph = 0
               End If
               
               If UserList(iuserindex).flags.Desnudo Then Call EfectoFrio(iuserindex)
               If UserList(iuserindex).flags.Meditando Then Call DoMeditar(iuserindex)
               If UserList(iuserindex).flags.Envenenado > 1 Then Call EfectoVeneno(iuserindex, bEnviarStats)
               If UserList(iuserindex).flags.AdminInvisible <> 1 And UserList(iuserindex).flags.Invisible = 1 Then Call EfectoInvisibilidad(iuserindex)
          'pluto:6.2
          If UserList(iuserindex).flags.Incor = True Then Call EfectoIncor(iuserindex)
               Call DuracionPociones(iuserindex)
               Call HambreYSed(iuserindex, bEnviarAyS)

               'pluto:7.0 vampiros regeneran como si descansasen - nati:los vampiros solo regeneraran vida.
                If UserList(iuserindex).raza = "Vampiro" Then
                    If UserList(iuserindex).Char.Body <> 9 And UserList(iuserindex).Char.Body <> 260 Then
                                'nati: cambio el metodo de sanar.
                If Not UserList(iuserindex).flags.Descansar Then
                Call Sanar(iuserindex, bEnviarStats, IntervaloRegeneraVampiro)
                                          Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloSinDescansar)
                                 Else
                                          Call Sanar(iuserindex, bEnviarStats, IntervaloRegeneraVampiro)
                                          Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloDescansar)
                                          If UserList(iuserindex).Stats.MaxHP = UserList(iuserindex).Stats.MinHP And _
                                             UserList(iuserindex).Stats.MaxSta = UserList(iuserindex).Stats.MinSta Then
                                                    Call SendData2(ToIndex, iuserindex, 0, 41)
                                                    Call SendData(ToIndex, iuserindex, 0, "||Has terminado de descansar." & "´" & FontTypeNames.FONTTYPE_info)
                                                    UserList(iuserindex).flags.Descansar = False
                'Call Sanar(iuserindex, bEnviarStats, SanaIntervaloDescansar)
                'Call Sanar(iuserindex, bEnviarStats, 10)
                'Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloDescansar)
                'Call RecStamina(iuserindex, bEnviarStats, 5)
                GoTo vampi
                End If
                End If
                    End If
                End If

               If Lloviendo Then
                    If Not Intemperie(iuserindex) Then
                                 If Not UserList(iuserindex).flags.Descansar And (UserList(iuserindex).flags.Hambre = 0 And UserList(iuserindex).flags.Sed = 0) Then
                                 'No esta descansando
                                          Call Sanar(iuserindex, bEnviarStats, SanaIntervaloSinDescansar)
                                If UserList(iuserindex).clase = "Pirata" Then
                                If UserList(iuserindex).flags.Ron > 0 Then
                                          Call RecStamina(iuserindex, bEnviarStats, (StaminaIntervaloSinDescansar * 2))
                                Else
                                          Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloSinDescansar)
                                End If
                                End If
                                If Not UserList(iuserindex).clase = "Pirata" Then
                                           Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloSinDescansar)
                                End If
                                 ElseIf UserList(iuserindex).flags.Descansar Then
                                 'esta descansando
                                          Call Sanar(iuserindex, bEnviarStats, SanaIntervaloDescansar)
                                          Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloDescansar)
                                          'termina de descansar automaticamente
                                          If UserList(iuserindex).Stats.MaxHP = UserList(iuserindex).Stats.MinHP And _
                                             UserList(iuserindex).Stats.MaxSta = UserList(iuserindex).Stats.MinSta Then
                                                    Call SendData2(ToIndex, iuserindex, 0, 41)
                                                    Call SendData(ToIndex, iuserindex, 0, "||Has terminado de descansar." & "´" & FontTypeNames.FONTTYPE_info)
                                                    UserList(iuserindex).flags.Descansar = False
                                          End If
                                 End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
                    End If
               Else
                    If Not UserList(iuserindex).flags.Descansar And (UserList(iuserindex).flags.Hambre = 0 And UserList(iuserindex).flags.Sed = 0) Then
                    'No esta descansando
                             Call Sanar(iuserindex, bEnviarStats, SanaIntervaloSinDescansar)
                                                    If UserList(iuserindex).clase = "Pirata" Then
                                If UserList(iuserindex).flags.Ron > 0 Then
                                          Call RecStamina(iuserindex, bEnviarStats, (StaminaIntervaloSinDescansar / 2))
                                Else
                                          Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloSinDescansar)
                                End If
                                End If
                                If Not UserList(iuserindex).clase = "Pirata" Then
                                           Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloSinDescansar)
                                End If
                                           'Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloSinDescansar)
                    ElseIf UserList(iuserindex).flags.Descansar Then
                    'esta descansando
                             Call Sanar(iuserindex, bEnviarStats, SanaIntervaloDescansar)
                             Call RecStamina(iuserindex, bEnviarStats, StaminaIntervaloDescansar)
                             'termina de descansar automaticamente
                             If UserList(iuserindex).Stats.MaxHP = UserList(iuserindex).Stats.MinHP And _
                                UserList(iuserindex).Stats.MaxSta = UserList(iuserindex).Stats.MinSta Then
                                     Call SendData2(ToIndex, iuserindex, 0, 41)
                                     Call SendData(ToIndex, iuserindex, 0, "||Has terminado de descansar." & "´" & FontTypeNames.FONTTYPE_info)
                                     UserList(iuserindex).flags.Descansar = False
                             End If
                    End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
               End If
vampi:
               If bEnviarStats Then
               Call SendUserStatsVida(iuserindex)
               Call SendUserStatsEnergia(iuserindex)
               End If
               If bEnviarAyS Then Call EnviarHambreYsed(iuserindex)

               If UserList(iuserindex).NroMacotas > 0 Then Call TiempoInvocacion(iuserindex)
       
Else 'ESTA MUERTO
       'PLUTO:6.2
       If MapInfo(UserList(iuserindex).Pos.Map).Terreno = "TORNEO" Or (UserList(iuserindex).Pos.Map = MapaTorneo2 And MinutoSinMorir > 3) Then
        UserList(iuserindex).flags.Torneo = 0
        UserList(iuserindex).Torneo2 = 0
       Call WarpUserChar(iuserindex, 296, 71, 65, True)
       ElseIf UserList(iuserindex).Pos.Map = mapi Then
       Call WarpUserChar(iuserindex, 110, 50, 50, True)
       ElseIf UserList(iuserindex).Pos.Map = 250 Then
       Call WarpUserChar(iuserindex, 247, 20, 59, True)
       'pluto:2-3-04 muere en egipto
        ElseIf UserList(iuserindex).Pos.Map > 177 And UserList(iuserindex).Pos.Map < 183 And UserList(iuserindex).Pos.Map <> 181 And UserList(iuserindex).flags.Muerto = 1 Then
        Call WarpUserChar(iuserindex, 181, 42, 66, True)
        'pluto:6.8 muere castillos sin puntos
        ElseIf UserList(iuserindex).Pos.Map > 165 And UserList(iuserindex).Pos.Map < 170 And UserList(iuserindex).Stats.PClan < 0 And UserList(iuserindex).flags.Muerto = 1 Then
        'nati: Ahora al morir en castillos no volvera a Nix.
        'Call WarpUserChar(iuserindex, 34, 35, 35, True)
        End If
       
End If 'NO ESTA MUERTO
     Else 'no esta logeado?
        'UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        'If UserList(iUserIndex).Counters.IdleCount > IntervaloParaConexion Then
        '      UserList(iUserIndex).Counters.IdleCount = 0
        '      Call CloseSocket(iUserIndex)
        'End If
     End If 'UserLogged
        
   End If

   Next iuserindex
    'If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
     '   lPermiteAtacar = 0
    'End If
'pluto:2.10
'If Not lPermiteTomar < IntervaloUserPuedeTomar Then
 '       lPermiteTomar = 0
  '  End If
 '-------------
   ' If Not lPermiteCast < IntervaloUserPuedeCastear Then
    '    lPermiteCast = 0
    'End If

    'If Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     '    lPermiteTrabajar = 0
    'End If
'pluto:2.8
'If Not lPermiteFlechas < IntervaloUserPuedeFlechas Then
         'lPermiteFlechas = 0
    'End If
  Exit Sub
hayerror:
  'DoEvents
End Sub

Private Sub Limpiado_Timer()
Static Vez As Byte
Dim X As Byte
Dim Y As Byte
Vez = Vez + 1
Select Case Vez
Case 1
Call SendData(ToAll, 0, 0, "||%%% NO TIREN OBJETOS AL SUELO, LIMPIADO DE MAPAS EN 15 SEGUNDOS.%%%" & "´" & FontTypeNames.FONTTYPE_TALK)
Case 2
Call SendData(ToAll, 0, 0, "||%%% NO TIREN OBJETOS AL SUELO, LIMPIADO DE MAPAS EN 10 SEGUNDOS.%%%" & "´" & FontTypeNames.FONTTYPE_TALK)
Case 3
Call SendData(ToAll, 0, 0, "||%%% NO TIREN OBJETOS AL SUELO, LIMPIADO DE MAPAS EN 5 SEGUNDOS.%%%" & "´" & FontTypeNames.FONTTYPE_TALK)
Case 4

Vez = 0
 Dim Mapasatu As Integer
 Call SendData(ToAll, 0, 0, "||Limpiando Mapas no seguros: Por favor espere ...." & "´" & FontTypeNames.FONTTYPE_info)


    For Mapasatu = 1 To NumMaps
 If MapaValido(Mapasatu) Then
 'pluto:6.9 añade casas arghal y mapa gm torneo
 If MapInfo(Mapasatu).Pk = True And Mapasatu <> 151 And Mapasatu <> 303 Then
  For Y = 1 To 100
  For X = 1 To 100
  If X > 0 And Y > 0 And X < 101 And Y < 101 Then
If MapData(Mapasatu, X, Y).OBJInfo.ObjIndex > 0 And MapData(Mapasatu, X, Y).Blocked = 0 Then
If ObjData(MapData(Mapasatu, X, Y).OBJInfo.ObjIndex).Agarrable = 0 Then
Call EraseObj(ToMap, 0, Mapasatu, 10000, Mapasatu, X, Y)
End If 'blocked
End If 'AGARRABLE
    End If 'x>0
    Next X
    Next Y
    
End If ' mapa inseguro pk=0
End If 'MAPAVALIDO

    Next ' mapasatu
'Call LogGM(UserList(UserIndex).Name, "/LIMPINOSEGURO")
Call SendData(ToAll, 0, 0, "||Limpiado de Mapas Completado." & "´" & FontTypeNames.FONTTYPE_info)


frmMain.Limpiado.Enabled = False

End Select
End Sub

'Private Sub Macrear_Timer()
'pluto:6.2-----
'Dim iuserindex As Integer

'on error GoTo errorh
'For iuserindex = 1 To MaxUsers
   'Conexion activa?
    'If UserList(iuserindex).ConnID <> -1 Then
      '¿User valido?
        'pluto:6.0A
        'If UserList(iuserindex).ConnIDValida And UserList(iuserindex).flags.UserLogged Then

        'If UserList(iuserindex).flags.Macreanda > 0 And UserList(iuserindex).Stats.MinSta > 0 Then Call EfectoMacrear(iuserindex, UserList(iuserindex).flags.Macreanda)

        'End If
    'End If
'Next

'-----------------
'Exit Sub
'errorh:

'End Sub

Private Sub mnuCerrar_Click()
Cerrando.Visible = True
Call WriteVar(IniPath & "eventodia.txt", "INIT", "Evento", val(EventoDia))
Dim X As Integer

Call SaveGuildsDB
DoEvents
Cerrando.Label1.Caption = "Grabando Datos ...."
Call grabaPJ
Call DoBackUp


'If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f



For X = 1 To MaxUsers
CloseSocket (X)
Next X
   
   
   
   
    For Each f In Forms
        Unload f
    Next


End Sub

Private Sub mnuCierra_Click()
Call WriteVar(IniPath & "eventodia.txt", "INIT", "Evento", val(EventoDia))
Call SendData(ToAll, 0, 0, "||%%%% EL SERVICIO TÉCNICO VA REINICIAR LA PC...%%%%" & "´" & FontTypeNames.FONTTYPE_TALK)

Cerrando.Show
Cerrando.Label1.Caption = "Preparando para Cerrar ...."
Call Sleep(2000)
Call SendData(ToAll, 0, 0, "||%%%% EL SERVICIO TÉCNICO VA REINICIAR LA PC...%%%%" & "´" & FontTypeNames.FONTTYPE_TALK)

Dim X As Integer

Call SaveGuildsDB
DoEvents
Cerrando.Label1.Caption = "Grabando Datos ...."
Call Sleep(2000)
Call SendData(ToAll, 0, 0, "||%%%% EL SERVICIO TÉCNICO VA REINICIAR LA PC...%%%%" & "´" & FontTypeNames.FONTTYPE_TALK)

Dim f
For X = 1 To MaxUsers
CloseSocket (X)
Next X

Call DoBackUp


'If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then

   
   
   
   
    For Each f In Forms
        Unload f
    Next
   
End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
On Error Resume Next
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
'If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
'pluto fusión
If Not FileExist(App.Path & "\logs\nokillwsapi.txt", vbNormal) Then
    If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
End If
End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim s As String
Dim nid As NOTIFYICONDATA

s = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub npcataca_Timer()

Dim npc As Integer

For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc


End Sub








Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler

Dim NpcIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim Mapa As Integer

If Not haciendoBK And Not haciendoBKPJ Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
'pluto:2.22-----------------------------------------------------
'If Npclist(NpcIndex).Name = "NPC SIN INICIAR" Then Npclist(NpcIndex).flags.NPCActive = False
'-------------------------------------------------------------
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
           
           'pluto:6.0A---------
           If Npclist(NpcIndex).NPCtype = 99 Then
            If MinutosMinotauro <= 0 And EstadoMinotauro = 1 Then
        EstadoMinotauro = 2
           Minotauro = ""
           Call SendData(ToAll, 0, 0, "||El Minotauro no ha sido encontrado y desaparece." & "´" & FontTypeNames.FONTTYPE_PARTY)
            Call LogCasino("Quitado " & Npclist(NpcIndex).Name & " del mapa: " & Npclist(NpcIndex).Pos.Map & " x:" & Npclist(NpcIndex).Pos.X & " y:" & Npclist(NpcIndex).Pos.Y)
             Call QuitarNPC(NpcIndex)
            GoTo ala:
            End If
           End If
           '---------------------
           
           'pluto:2.4.1
           If Npclist(NpcIndex).flags.Paralizado > 0 And Npclist(NpcIndex).NPCtype <> 6 Then
                 Call EfectoParalisisNpc(NpcIndex)
           Else
                'Usamos AI si hay algun user en el mapa
                Mapa = Npclist(NpcIndex).Pos.Map
                If Mapa > 0 Then
                     If MapInfo(Mapa).NumUsers > 0 Then
                             If Npclist(NpcIndex).Movement <> ESTATICO Then
                                   Call NPCAI(NpcIndex)
                             End If
       'pluto:2.18
       'If Npclist(NpcIndex).NPCtype = 83 Then
       'Call HablaPirata(NpcIndex)
       'End If
       '-----------------------
                     End If 'mapinfo
                End If 'mapa>0
                
           End If ' paralizado
                   
        End If 'active
ala:
    Next NpcIndex
    
End If


Exit Sub

ErrorHandler:
 Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
 Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub Timer1_Timer()

Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).flags.UserLogged Then _
        If UserList(i).flags.Oculto = 1 Then Call DoPermanecerOculto(i)
           
        
Next i

End Sub
Private Sub tesorocaballero_Timer()
Static cabaoscuro As Integer

cabaoscuro = cabaoscuro + 1

If cabaoscuro = 120 Then
    tesorocaballero = 0
    cabaoscuro = 0
    End If

End Sub

Private Sub timermomia_Timer()
Static momia As Integer

momia = momia + 1
If momia = 60 Then
Tesoromomia = 0
momia = 0
End If


End Sub

Private Sub tLluvia_Timer()
On Error GoTo errhandler

Dim iCount As Integer

If Lloviendo Then
   For iCount = 1 To LastUser
    Call EfectoLluvia(iCount)
   Next iCount
End If

Exit Sub
errhandler:
Call LogError("tLluvia")
End Sub

Private Sub tLluviaEvent_Timer()

On Error GoTo ErrorHandler

Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    MinutosSinLluvia = MinutosSinLluvia + 1
    If MinutosSinLluvia >= 60 And MinutosSinLluvia < 1440 Then
            If RandomNumber(1, 100) <= 3 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData2(ToAll, 0, 0, 20, "1")
            End If
    ElseIf MinutosSinLluvia >= 1440 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData2(ToAll, 0, 0, 20, "1")
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 3 Then
            Lloviendo = False
            Call SendData2(ToAll, 0, 0, 20, "0")
            MinutosLloviendo = 0
    Else
            If RandomNumber(1, 100) <= 50 Then
                Lloviendo = False
                MinutosLloviendo = 0
                Call SendData2(ToAll, 0, 0, 20, "0")
                '[\END]
            End If
    End If
End If


Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")



End Sub


Private Sub Torneo_Timer()
Static Vez As Byte
Static GenteClan2(1 To 50) As String
Static GenteClan1(1 To 50) As String
If TClanOcupado = 0 Then Exit Sub

If TClanOcupado = 1 Then

Vez = Vez + 1
Call SendData(ToAll, 0, 0, "||" & MsgTorneo & "´" & FontTypeNames.FONTTYPE_TALK)
    If Vez = 5 Then
    Call SendData(ToAll, 0, 0, "||Nadie aceptó el desafío para el Duelo de Clanes!!" & "´" & FontTypeNames.FONTTYPE_pluto)
    Vez = 0
    TClanOcupado = 0
    TorneoClan(1).Nombre = ""
    TorneoClan(2).Nombre = ""
    TorneoClan(1).numero = 0
    TorneoClan(2).numero = 0
    End If
End If

If TClanOcupado = 2 Then
Vez = Vez + 1
If Vez = 6 Then
TClanOcupado = 3
Vez = 0
End If
Call SendData(ToClan, 0, 0, MsgTorneo & "´" & FontTypeNames.FONTTYPE_pluto)
End If

If TClanOcupado = 3 Then
Dim i As Integer, E As Integer
'Dim GenteClan1(1 To 20) As String
Dim GuildP1(1 To 50) As Integer
'Dim GenteClan2(1 To 20) As String
Dim GuildP2(1 To 50) As Integer
Dim a1 As Byte
Dim a2 As Byte
Dim nomaux1 As String
Dim NomAux2 As String
Dim dniaux As Integer

Dim n As Integer
'revisamos los users de clan
a1 = 0
a2 = 0
For n = 1 To LastUser
    'clan 1
    If UserList(n).GuildInfo.GuildName = TorneoClan(1).Nombre And UserList(n).flags.NoTorneos = False Then
    a1 = a1 + 1
    GenteClan1(a1) = UserList(n).Name
    GuildP1(a1) = UserList(n).GuildInfo.GuildPoints
    End If
    'clan 2
    If UserList(n).GuildInfo.GuildName = TorneoClan(2).Nombre And UserList(n).flags.NoTorneos = False Then
    a2 = a2 + 1
    GenteClan2(a2) = UserList(n).Name
    GuildP2(a2) = UserList(n).GuildInfo.GuildPoints
    End If

Next n


'ordenamos clan 1


            For E = 1 To a1
                For i = 1 To a1
    
           If GuildP1(i) < GuildP1(E) Then
              nomaux1 = GenteClan1(i)
              GenteClan1(i) = GenteClan1(E)
              GenteClan1(E) = nomaux1
              'pluto:6.0A
              'LevelAux = GuildLevel(i)
              'GuildLevel(i) = GuildLevel(e)
              'GuildLevel(e) = LevelAux
           
              dniaux = GuildP1(i)
              GuildP1(i) = GuildP1(E)
              GuildP1(E) = dniaux
           End If
   
                Next i
            Next E

            
'ordenamos clan 2
            For E = 1 To a2
                For i = 1 To a2
    
           If GuildP2(i) < GuildP2(E) Then
              NomAux2 = GenteClan2(i)
              GenteClan2(i) = GenteClan2(E)
              GenteClan2(E) = NomAux2
              'pluto:6.0A
              'LevelAux = GuildLevel(i)
              'GuildLevel(i) = GuildLevel(e)
              'GuildLevel(e) = LevelAux
           
                
              dniaux = GuildP2(i)
              GuildP2(i) = GuildP2(E)
              GuildP2(E) = dniaux
           End If
      
                Next i
            Next E



TClanOcupado = 4

End If 'tclanocupado=3

If TClanOcupado = 4 Then
frmMain.Torneo.Interval = 30000
    Vez = Vez + 1
    'mostramos participantes clan 1
MsgTorneo = ""
    For n = 1 To TClanNumero
    MsgTorneo = MsgTorneo & " " & GenteClan1(n) & ","
    Next
Call SendData(ToClan, 0, 0, "Clan " & TorneoClan(1).Nombre & ": " & MsgTorneo & "´" & FontTypeNames.FONTTYPE_pluto)
'mostramos participantes clan 2
MsgTorneo = ""
    For n = 1 To TClanNumero
    MsgTorneo = MsgTorneo & " " & GenteClan2(n) & ","
    Next
Call SendData(ToClan, 0, 0, "Clan " & TorneoClan(2).Nombre & ": " & MsgTorneo & "´" & FontTypeNames.FONTTYPE_pluto)
Call SendData(ToClan, 0, 0, "Vayan equipandose en unos instantes serán Teletransportados." & "´" & FontTypeNames.FONTTYPE_pluto)

If Vez = 6 Then
Vez = 0
TClanOcupado = 5
End If

End If



'teleportamos usuarios................
        If TClanOcupado = 5 Then
 Dim Tindex As Integer
 For n = 1 To TClanNumero
'clan1.....
If GenteClan1(n) = "" Then GoTo otro
Tindex = NameIndex(GenteClan1(n))
If Tindex = 0 Then GoTo otro
'pluto:6.8 staminia cero para transformados
If UserList(Tindex).flags.Angel > 0 Or UserList(Tindex).flags.Demonio > 0 Or UserList(Tindex).flags.Morph > 0 Then UserList(Tindex).Stats.MinSta = 0
Call WarpUserChar(Tindex, 292, 53, 50 + n, True)
TorneoClan(1).numero = TorneoClan(1).numero + 1
otro:
 'clan2.....
 If GenteClan2(n) = "" Then GoTo otro2
Tindex = NameIndex(GenteClan2(n))
If Tindex = 0 Then GoTo otro2
'pluto:6.8 staminia cero para transformados
If UserList(Tindex).flags.Angel > 0 Or UserList(Tindex).flags.Demonio > 0 Or UserList(Tindex).flags.Morph > 0 Then UserList(Tindex).Stats.MinSta = 0
Call WarpUserChar(Tindex, 292, 70, 50 + n, True)
TorneoClan(2).numero = TorneoClan(2).numero + 1
otro2:
 Next
 TClanOcupado = 6
            End If 'clanocupado=5
'----------------------------------------------
If TClanOcupado = 6 Then
Dim Mensa As String
Dim ii As Integer
Dim Punto1 As Long
Dim Punto2 As Long
Dim i1 As Integer
Dim i2 As Integer
Dim Puntazos1 As Integer
Dim Puntazos2 As Integer

  If TorneoClan(1).numero < 1 Then
   
   Select Case TorneoClan(2).numero
   Case Is <= 1
   Mensa = " ha ganado un duelo muy igualado contra el clan "
   Case 2
   Mensa = " ha obtenido una gran victoria en su duelo contra el clan "
   Case 3
   Mensa = " ha arrasado en su duelo contra el clan "
   Case 4
   Mensa = " ha dado una tremenda paliza en su duelo al clan "
   Case 5
   Mensa = " ha destrozado sin esfuerzo al clan "
   Case 6
   Mensa = " ha aniquilado y humillado al clan "
   End Select
   

'buscamos los puntos actuales
For ii = 1 To Guilds.Count
If UCase$(TorneoClan(1).Nombre) = UCase$(Guilds(ii).GuildName) Then
Punto1 = Guilds(ii).Reputation
i1 = ii
End If
If UCase$(TorneoClan(2).Nombre) = UCase$(Guilds(ii).GuildName) Then
Punto2 = Guilds(ii).Reputation
i2 = ii
End If
Next ii

'sumamos los obtenidos por derrota
Puntazos1 = Int(Guilds(i1).Reputation / 500)
Puntazos2 = Int(Guilds(i2).Reputation / 500)
If Puntazos1 < 1 Then Puntazos1 = 1
If Puntazos2 < 1 Then Puntazos2 = 1
Guilds(i2).Reputation = Guilds(i2).Reputation + Puntazos1
Guilds(i1).Reputation = Guilds(i1).Reputation - Puntazos1
'pluto:6.9
Guilds(i1).PuntosTorneos = Guilds(i1).PuntosTorneos - Puntazos1
Guilds(i2).PuntosTorneos = Guilds(i2).PuntosTorneos + Puntazos1

'If Guilds(i2).PuntosTorneos > PClan1Torneo Then
'PClan1Torneo = Guilds(i2).PuntosTorneos
'Clan1Torneo = TorneoClan(2).Nombre
'ElseIf Guilds(i2).PuntosTorneos > PClan2Torneo Then
'PClan2Torneo = Guilds(i2).PuntosTorneos
'Clan2Torneo = TorneoClan(2).Nombre
'End If


Call SendData(ToAll, 0, 0, "||El Clan " & TorneoClan(2).Nombre & Mensa & TorneoClan(1).Nombre & "!!" & "´" & FontTypeNames.FONTTYPE_pluto)
Call SendData(ToClan, 0, 0, "El Clan " & TorneoClan(2).Nombre & " ha ganado " & Puntazos1 & " Puntos de Clan" & "´" & FontTypeNames.FONTTYPE_pluto)
Call SendData(ToClan, 0, 0, "El Clan " & TorneoClan(1).Nombre & " ha pérdido " & Puntazos1 & " Puntos de Clan" & "´" & FontTypeNames.FONTTYPE_pluto)
Call SendData(ToClan, 0, 0, "En unos instantes serán expulsados de la sala." & "´" & FontTypeNames.FONTTYPE_pluto)
   TClanOcupado = 7
    Exit Sub
  End If
  
  If TorneoClan(2).numero < 1 Then
  
  Select Case TorneoClan(1).numero
   Case Is <= 1
   Mensa = " ha ganado un duelo muy igualado contra el clan "
   Case 2
   Mensa = " ha obtenido una gran victoria en su duelo contra el clan "
   Case 3
   Mensa = " ha arrasado en su duelo contra el clan "
   Case 4
   Mensa = " ha dado una tremenda paliza en su duelo al clan "
   Case 5
   Mensa = " ha destrozado sin esfuerzo al clan "
   Case 6
   Mensa = " ha aniquilado y humillado al clan "
   End Select

'buscamos los puntos actuales
For ii = 1 To Guilds.Count
    If UCase$(TorneoClan(1).Nombre) = UCase$(Guilds(ii).GuildName) Then
    Punto1 = Guilds(ii).Reputation
    i1 = ii
    End If
    If UCase$(TorneoClan(2).Nombre) = UCase$(Guilds(ii).GuildName) Then
    Punto2 = Guilds(ii).Reputation
    i2 = ii
    End If
Next ii

'sumamos los obtenidos por derrota
Dim joder As String
'joder = Guilds(i1).Reputation
Puntazos1 = Int(Guilds(i1).Reputation / 500)
Puntazos2 = Int(Guilds(i2).Reputation / 500)
If Puntazos1 < 1 Then Puntazos1 = 1
If Puntazos2 < 1 Then Puntazos2 = 1
Guilds(i2).Reputation = Guilds(i2).Reputation - Puntazos2
Guilds(i1).Reputation = Guilds(i1).Reputation + Puntazos2
'pluto:6.9
Guilds(i1).PuntosTorneos = Guilds(i1).PuntosTorneos + Puntazos2
Guilds(i2).PuntosTorneos = Guilds(i2).PuntosTorneos - Puntazos2

'If Guilds(i1).PuntosTorneos > PClan1Torneo Then
'PClan1Torneo = Guilds(i1).PuntosTorneos
'Clan1Torneo = TorneoClan(1).Nombre
'ElseIf Guilds(i1).PuntosTorneos > PClan2Torneo Then
'PClan2Torneo = Guilds(i1).PuntosTorneos
'Clan2Torneo = TorneoClan(1).Nombre
'End If

  
   Call SendData(ToAll, 0, 0, "||El Clan " & TorneoClan(1).Nombre & Mensa & TorneoClan(2).Nombre & "!!" & "´" & FontTypeNames.FONTTYPE_pluto)
Call SendData(ToClan, 0, 0, "El Clan " & TorneoClan(1).Nombre & " ha ganado " & Puntazos2 & " Puntos de Clan" & "´" & FontTypeNames.FONTTYPE_pluto)
Call SendData(ToClan, 0, 0, "El Clan " & TorneoClan(2).Nombre & " ha pérdido " & Puntazos2 & " Puntos de Clan" & "´" & FontTypeNames.FONTTYPE_pluto)
Call SendData(ToClan, 0, 0, "En unos instantes serán expulsados de la sala." & "´" & FontTypeNames.FONTTYPE_pluto)
TClanOcupado = 7
    Exit Sub
    End If
End If 'tclanocupado=6
  
  If TClanOcupado = 7 Then
   TClanOcupado = 0
   'pluto:6.9
  
  
  'ordenamos puntostorneos
ReDim PuntClan(1 To Guilds.Count) As Integer
ReDim NomClan(1 To Guilds.Count) As String
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
  
    TorneoClan(1).Nombre = ""
    TorneoClan(1).numero = 0
    TorneoClan(2).Nombre = ""
    TorneoClan(2).numero = 0
    'pluto:6.9
    For n = 1 To MaxUsers
        If UserList(n).flags.UserLogged = True Then
    If UserList(n).Pos.Map = 292 Then Call WarpUserChar(n, 296, 59, 54, True)
        End If
    Next
    
    For n = 1 To 50
    GenteClan1(n) = ""
    GenteClan2(n) = ""
    Next
  

 'pluto:6.9
 Dim Mapasatu As Integer
 Mapasatu = 292
 Dim X As Byte
 Dim Y As Byte
  For Y = 1 To 100
  For X = 1 To 100
  If X > 0 And Y > 0 And X < 101 And Y < 101 Then
If MapData(Mapasatu, X, Y).OBJInfo.ObjIndex > 0 And MapData(Mapasatu, X, Y).Blocked = 0 Then
If ObjData(MapData(Mapasatu, X, Y).OBJInfo.ObjIndex).Agarrable = 0 Then
Call EraseObj(ToMap, 0, Mapasatu, 10000, Mapasatu, X, Y)
End If 'blocked
End If 'AGARRABLE
    End If 'x>0
    Next X
    Next Y
    'Call LogGM(UserList(UserIndex).Name, "/LIMPI Mapa: " & Mapasatu)
    'Call SendData(ToIndex, UserIndex, 0, "||Limpiado mapa: " & Mapasatu & "´" & FontTypeNames.FONTTYPE_talk)

    Exit Sub

End If
  
  
 
 



End Sub

Private Sub tPiqueteC_Timer()
On Error Resume Next
Dim MiObj As obj
Static segundos As Integer
Dim nvv As Byte
segundos = segundos + 6

Dim i As Integer
'pluto:2.4.5
UserCiu = 0: UserCrimi = 0
'pluto:2.15------------------------
'If Left$(Time, 6) = "0:00:0" And yaya = 0 Then
'Call Logrenumusers(str$(ReNumUsers), str$(MediaUsers))
'AyerMediaUsers = MediaUsers
'AyerReNumUsers = ReNumUsers
'Horayer = HoraHoy
'ReNumUsers = 0
'MediaUser = 0
'MediaVez = 0
'yaya = 1
'End If
'pluto:6.8
'-----------------------------------------
'pluto:6.8--------------------------------
If Date$ <> HOYESDIA Then
Call Logrenumusers(str$(ReNumUsers), str$(MediaUsers))
AyerMediaUsers = MediaUsers
AyerReNumUsers = ReNumUsers
Horayer = HoraHoy
ReNumUsers = 0
MediaUser = 0
MediaVez = 0
'yaya = 1
HOYESDIA = Date$
    'pluto:6.8
    If EventoDia > 0 Then
    i = RandomNumber(1, 100)
    RecordEventoDia = EventoDia
    'porcentajes de eventodia-----------
    Select Case i
    Case Is < 40
        EventoDia = 1
    Case 40 To 60
        EventoDia = 4
    Case 61 To 80
        EventoDia = 3
    Case 81 To 90
        EventoDia = 2
    Case Is > 90
        EventoDia = 5
    End Select
    '--------------------------------------
If RecordEventoDia = EventoDia Then
    i = RandomNumber(1, 100)
    'porcentajes de eventodia-----------
    Select Case i
    Case Is < 40
        EventoDia = 1
    Case 40 To 60
        EventoDia = 4
    Case 61 To 80
        EventoDia = 3
    Case 81 To 90
        EventoDia = 2
    Case Is > 90
        EventoDia = 5
    End Select
    '--------------------------------------
End If
    
    'EventoDia = i
    'EventoDia = 5
    Select Case EventoDia
    Case 1
    Call CargarDiaEspecial
    Case 4
    Call CargarDiaEspecial
    End Select

    Call WriteVar(IniPath & "eventodia.txt", "INIT", "Evento", val(EventoDia))
    End If
    Select Case EventoDia
Case 1
Call SendData2(ToIndex, i, 0, 99, NombreBichoDelDia)
Case 2
Call SendData2(ToIndex, i, 0, 101)
Case 3
Call SendData2(ToIndex, i, 0, 102)
Case 4
Call SendData2(ToIndex, i, 0, 103, NombreBichoDelDia)
Case 5
Call SendData2(ToIndex, i, 0, 104)
End Select
End If
'---------------------------------------
For i = 1 To LastUser

    
    If UserList(i).flags.UserLogged Then
 'pluto:6.0A--------
If UserList(i).flags.Sed > 0 Or UserList(i).flags.Hambre > 0 Then Call QuitarSta(i, 40)

'------------------



'pluto:6.0A----------------esto lo traigo de doevents-----------------------
Dim Mapcasa As Byte
Dim Mapcasa2 As Byte

Dim Map As Integer
Mapcasa = 171
Mapcasa2 = 177
Map = UserList(i).Pos.Map

'Casa = 52




If UserList(i).flags.Privilegios > 0 Or UserList(i).flags.Muerto > 0 Then GoTo yupi2
If UserList(i).Pos.Map <> Mapcasa And UserList(i).Pos.Map <> Mapcasa2 Then GoTo yupi
Dim Casa As Byte
Casa = RandomNumber(1, 150)
'pluto:desequipar sala casa

If Casa = 50 And UserList(i).flags.Morph = 0 And UserList(i).flags.Angel = 0 And UserList(i).flags.Demonio = 0 Then
Call SendData(ToIndex, i, 0, "|| Los Espiritus de la Casa te hacen perder el inventario" & "´" & FontTypeNames.FONTTYPE_TALK)
Call SendData(ToMap, 0, Map, "TW" & 115)
Call TirarTodosLosItems(i)
Call SendData2(ToPCArea, i, UserList(i).Pos.Map, 22, UserList(i).Char.CharIndex & "," & 33 & "," & 1)
End If
If Casa = 51 Or Casa = 50 Or Casa = 30 Or Casa = 31 Then
Call SendData(ToIndex, i, 0, "|| Los Espiritus de la Casa te hacen perder Oro." & "´" & FontTypeNames.FONTTYPE_TALK)
Call SendData(ToMap, 0, Map, "TW" & 115)
If UserList(i).Pos.Map = Mapcasa Then Call TirarOro(30000, i) Else Call TirarOro(10000, i)
Call SendUserStatsOro(i)
Call SendData2(ToPCArea, i, UserList(i).Pos.Map, 22, UserList(i).Char.CharIndex & "," & 33 & "," & 1)
End If
If Casa = 52 And UserList(i).flags.Morph = 0 And UserList(i).flags.Navegando = 0 And UserList(i).flags.Invisible = 0 And UserList(i).flags.Angel = 0 And UserList(i).flags.Demonio = 0 Then
UserList(i).flags.Morph = UserList(i).Char.Body
UserList(i).Counters.Morph = IntervaloMorphPJ
Call SendData(ToIndex, i, 0, "|| Los Espiritus de la Casa te transforman en Cerdo." & "´" & FontTypeNames.FONTTYPE_TALK)
'[gau]
Call ChangeUserChar(ToMap, 0, UserList(i).Pos.Map, i, val(6), val(0), UserList(i).Char.Heading, UserList(i).Char.WeaponAnim, UserList(i).Char.ShieldAnim, UserList(i).Char.CascoAnim, UserList(i).Char.Botas)
Call SendData2(ToPCArea, i, UserList(i).Pos.Map, 22, UserList(i).Char.CharIndex & "," & 33 & "," & 1)
Call SendData(ToMap, 0, Map, "TW" & 115)
End If
If Casa = 53 Then
Call SendData(ToIndex, i, 0, "|| Los Espiritus de la Casa te teleportan fuera de ella." & "´" & FontTypeNames.FONTTYPE_TALK)
Call SendData(ToMap, 0, Map, "TW" & 115)
If Map = 171 Then Call WarpUserChar(i, 174, 64, 25, True) Else Call WarpUserChar(i, 38, 17, 49, True)
End If
If Casa > 53 And Casa < 57 And UserList(i).flags.Paralizado = 0 Then
Call SendData(ToIndex, i, 0, "|| Los Espiritus de la Casa te han Paralizado." & "´" & FontTypeNames.FONTTYPE_TALK)
Call SendData(ToMap, 0, Map, "TW" & 115)
Call SendData2(ToPCArea, i, UserList(i).Pos.Map, 22, UserList(i).Char.CharIndex & "," & 33 & "," & 1)
UserList(i).flags.Paralizado = 1
UserList(i).Counters.Paralisis = IntervaloParalisisPJ
Call SendData2(ToIndex, i, 0, 68)
Call SendData2(ToIndex, i, 0, 15, UserList(i).Pos.X & "," & UserList(i).Pos.Y)
End If

yupi:
'pluto:2.17 quitamos gusano mapa 20
If ((UserList(i).Pos.Map > 14 And UserList(i).Pos.Map < 18) Or (UserList(i).Pos.Map > 20 And UserList(i).Pos.Map < 22)) Then
If Casa > 120 Then Call Gusano(i)
End If

yupi2:












'pluto:2.24 ---------------------------------------------
If UserList(i).flags.Privilegios > 0 Then GoTo alli
If UserList(i).Stats.GLD + UserList(i).Stats.Banco > MoroOn Then
MoroOn = UserList(i).Stats.GLD + UserList(i).Stats.Banco
NMoroOn = UserList(i).Name

If MoroOn > Moro Then
Moro = MoroOn
NMoro = NMoroOn
End If

End If

If UserList(i).Remort = 0 Then
nvv = UserList(i).Stats.ELV
Else
nvv = UserList(i).Stats.ELV + 55
End If

If Not Criminal(i) And nvv > NivCiuON Then NivCiuON = nvv: NNivCiuON = UserList(i).Name
If Criminal(i) And nvv > NivCrimiON Then NivCrimiON = nvv: NNivCrimiON = UserList(i).Name
alli:

'fin pluto:2.24------------------------------



'pluto:2.4.5
If UserList(i).Reputacion.Promedio >= 0 Then UserCiu = UserCiu + 1 Else UserCrimi = UserCrimi + 1



'pluto:2.22 añado piquete al tigger 3
            If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = 3 Then
                    UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                    Call SendData(ToIndex, i, 0, "E1")
                    If UserList(i).Counters.PiqueteC > 23 Then
                            UserList(i).Counters.PiqueteC = 0
                            Call Encarcelar(i, 3)
                    End If
            Else
    If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
    'muere en torneo: pluto:2.12 añade torneo2
       

       ' If Not Criminal(i) Then
           'Call WarpUserChar(i, Banderbill.Map, Banderbill.X, Banderbill.Y, True)
        'Else
           ' Call WarpUserChar(i, 170, 34, 34, True)
        'End If
    End If




    
    'pluto:2.12
If UserList(i).Pos.Map = MapaTorneo2 And UserList(i).Torneo2 > Torneo2Record Then
Torneo2Record = UserList(i).Torneo2
Torneo2Name = UserList(i).Name
Call SendData2(ToMap, 0, MapaTorneo2, 96, Torneo2Name & "," & Torneo2Record & "," & TorneoBote)
End If
 
 'PLUTO:6.2
If UserList(i).flags.ComproMacro > 0 Then
 UserList(i).flags.ComproMacro = UserList(i).flags.ComproMacro - 1
 Call SendData(ToIndex, i, 0, "|| COMPROBANDO MACRO ASISTIDO: Debes Escribir /macro " & CodigoMacro & " antes de " & UserList(i).flags.ComproMacro & " segundos." & "´" & FontTypeNames.FONTTYPE_TALK)
    
    If UserList(i).flags.ComproMacro < 1 Then
    'COMPROBANDOMACRO = False
    UserList(i).flags.ComproMacro = 0
    UserList(i).flags.Macreanda = 0
    Call TirarTodo(i)
    Call Encarcelar(i, 20, "AntiMacro")
    Call SendData(ToGM, 0, 0, "||AntiMacro Cárcel para: " & UserList(i).Name & "´" & FontTypeNames.FONTTYPE_TALK)
    Call SendData(ToIndex, i, 0, "O3")
    End If
 
End If
 
 


Dim obj As ObjData

'¿Hay algun obj?
If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).OBJInfo.ObjIndex > 0 And UserList(i).flags.Muerto = 1 Then
UserList(i).Counters.bloqueo = UserList(i).Counters.bloqueo + 1
Call SendData(ToIndex, i, 0, "E2")
If UserList(i).Counters.bloqueo > 10 Then
UserList(i).Counters.bloqueo = 0
 Call Encarcelar(i, 25)
  End If
   
   Else
   UserList(i).Counters.bloqueo = 0
   End If
            
 End If
         'pluto:hoy
If segundos > 6 Then
If UserList(i).Char.FX > 37 And UserList(i).Char.FX < 68 Then
Call SendData2(ToMap, i, UserList(i).Pos.Map, 22, UserList(i).Char.CharIndex & "," & 0 & "," & 0)
UserList(i).Char.FX = 0
End If
End If
 'pluto:2.11
 If segundos > 12 And UserList(i).GranPoder > 0 Then
Call SendData2(ToMap, i, UserList(i).Pos.Map, 22, UserList(i).Char.CharIndex & "," & 68 & "," & 1)
UserList(i).Char.FX = 68
Call SendData(ToPCArea, i, UserList(i).Pos.Map, "TW" & 147)
End If

 
 
 
            'If Segundos >= 18 Then
            If segundos >= 18 Then UserList(i).Counters.Pasos = 0
            'End If
            
    'End If
Next i

If segundos >= 18 Then segundos = 0
   
'Exit Sub

'errhandler:
   ' Call LogError("Error en tPiqueteC_Timer")
End Sub


Private Sub tTraficStat_Timer()

'Dim i As Integer
'
'If frmTrafic.Visible Then frmTrafic.lstTrafico.Clear
'
'For i = 1 To LastUser
'    If UserList(i).Flags.UserLogged Then
'        If frmTrafic.Visible Then
'            frmTrafic.lstTrafico.AddItem UserList(i).Name & " " & UserList(i).BytesTransmitidosUser + UserList(i).BytesTransmitidosSvr & " bytes per second"
'        End If
'        UserList(i).BytesTransmitidosUser = 0
'        UserList(i).BytesTransmitidosSvr = 0
'    End If
'Next i

End Sub

Private Sub Userslst_Click()

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC

    Dim a As Byte
    Dim i As Integer
    
    Dim NewIndex As Integer
    'pluto:2.14---------------------
    
    'Static ulti As String
    'Static conta As Integer
    'Static conta2 As Integer
    'Dim ataqnegro(10) As String
    'Static vg As Long
   'pluto:2.15
If joputa2 = 1 Then
Call SendData(ToGM, 0, 0, "||Ip: " & TCPServ.GetIP(ID) & "´" & FontTypeNames.FONTTYPE_TALK)
End If
    'pluto:2.15
If Joputa = "" Then GoTo kgg
    If TCPServ.GetIP(ID) = Joputa Then Exit Sub
kgg:
    
    
    'vg = vg + 1
    ' If TCPServ.GetIP(ID) = ulti Then
    'conta = conta + 1
   ' Else
   ' conta = 0
    'End If
    
    'ulti = TCPServ.GetIP(ID)
    
    'If conta > 55 Or ulti = "" Then
    'TCPServ.CerrarSocket (ID)
     'Exit Sub
   ' End If
    '-----------------------------------
    
    NewIndex = NextOpenUser

If NewIndex < 1 Then
LogCriticEvent ("NEWINDEX > CERO")
Exit Sub
End If


a = 1
    If NewIndex <= MaxUsers And NewIndex > 0 Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato ID, NewIndex
        
a = 2
        If aDos.MaxConexiones(TCPServ.GetIP(ID)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(ID))
            Call CloseSocket(NewIndex, True)
            Exit Sub
        End If
a = 3
        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).ip = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
a = 4
        'For i = 1 To BanIps.Count
            'If BanIps.Item(i) = TCPServ.GetIP(ID) Then
                'Call ResetUserSlot(NewIndex)
                'Exit Sub
            'End If
        'Next i

    Else
a = 5
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer:NuevaConexion " & Err.Description & " Newindex: " & NewIndex & " NextOpen: " & NextOpenUser & " ID: " & ID & " Loc: " & a)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
Dim t() As String
Dim loopc As Long
Dim RD As String
'pluto:2.14


If MiDato < 1 Then Exit Sub


On Error GoTo errorh
If UserList(MiDato).ConnID <> UserList(MiDato).ConnID Then
    Call LogError("Recibi un read de un usuario con ConnId alterada")
    Exit Sub
End If


RD = StrConv(Datos, vbUnicode)





UserList(MiDato).RDBuffer = UserList(MiDato).RDBuffer & RD

t = Split(UserList(MiDato).RDBuffer, ENDC)
If UBound(t) > 0 Then
    UserList(MiDato).RDBuffer = t(UBound(t))
    
    For loopc = 0 To UBound(t) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If t(loopc) <> "" Then
                If Not UserList(MiDato).CommandsBuffer.Push(t(loopc)) Then
                    Call LogError("Cerramos por no encolar. Userindex:" & MiDato)
                    Call CloseSocket(MiDato)
                End If
            End If
        Else ' no encolamos los comandos (MUY VIEJO)
              If UserList(MiDato).ConnID <> -1 Then

                Call HandleData(MiDato, t(loopc))
              Else
                Exit Sub
              End If
        End If
    Next loopc
End If
Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.Description)

End Sub





#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub ws_server_ConnectionRequest(ByVal requestID As Long)
'cuando llegue una petición para conectarsenos la aceptamos
Me.ws_server.Close
Me.ws_server.Accept requestID
Debug.Print "conexion aceptada"
End Sub

Private Sub ws_server_DataArrival(ByVal bytesTotal As Long)
'aqui rebimos los datos que se envían y hacemos un reconocimiento
'así podemos tomar una decision
Dim str_datos As String
Dim i As Long
'el emvío se almacena en una cadena (str_datos)
Me.ws_server.GetData str_datos

'el algoritmo de a continuación se encarga de desglosar la cadena
'la cual unimos en 3 desde el cliente, se incluye el patron(archivo)
'la ruta y la longitud
If Mid(str_datos, 1, 7) = "archivo" Then
    'si el patron coincide lo borramos de la cadena y seguimos analizando
    str_datos = Mid(str_datos, 9, Len(str_datos) - 7)
    For i = 1 To Len(str_datos)
        If Mid(str_datos, 1, 1) <> "|" Then
            'vamos concatenando la ruta hasta que encontremos el carcater
            ' "|" si éste llega sabremos que la ruta esta completa
            Me.str_ruta = Me.str_ruta + Mid(str_datos, 1, 1)
        ElseIf Mid(str_datos, 1, 1) = "|" Then
            ' aqui el bucle encontró al "|" ,es borrado y luego sale del for
                str_datos = Mid(str_datos, 2, Len(str_datos) - 1)
            Exit For
        End If
        'aqui se va concatenando la cadena y borrando el caracter almacenado
        str_datos = Mid(str_datos, 2, Len(str_datos) - 1)
    Next
    'lo restante de la cadena es la longitud del archivo
    lng_tamaño_archivo = val(str_datos)
'una ves capturados los valores informamos al cliente
Me.ws_server.SendData "msg_peticion_aceptada"
'
'la variable 'str_archivo_temporal' guardará el contenido del archivo
'que mande el cliente, la inicializamos en ""(vacío)
Me.str_archivo_temporal = ""

Else
    'si el tamoño del archivo temporal es diferente a la longitud propuesta; vamos uniendo el contenido del archivo recogido
    If Len(str_archivo_temporal) <> lng_tamaño_archivo Then str_archivo_temporal = str_archivo_temporal + str_datos
    'cuando es el mismo tamaño se entiende que ya se envió todo el archivo
    If Len(str_archivo_temporal) = lng_tamaño_archivo Then
        'abrimos un archivo binario para escribirlo(put) en la ruta asignada(str_ruta)
otronombre:
Dim n As Byte
str_ruta = Now & ".jpg"
For n = 1 To Len(str_ruta)
If Mid(str_ruta, n, 1) = ":" Then Mid(str_ruta, n, 1) = " "
If Mid(str_ruta, n, 1) = "/" Then Mid(str_ruta, n, 1) = " "
Next

If FileExist(Me.str_ruta, vbArchive) Then GoTo otronombre
        
        
     
        Open Me.str_ruta For Binary As #1
        Put #1, 1, str_archivo_temporal
        'cerramos el archivo
        Close #1
        Me.str_ruta = ""
        'alertamos al cliente que el archivo a sido recibido satisfactoriamente
        Me.ws_server.SendData "msg_archivo_recibido"
        Call SendData(ToGM, 0, 0, "|| Archivo Recibido " & "´" & FontTypeNames.FONTTYPE_info)

    End If
End If
End Sub

