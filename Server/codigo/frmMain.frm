VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AODrag Servidor"
   ClientHeight    =   4275
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   11175
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
   ScaleHeight     =   4275
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer TimerEfectos 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10320
      Top             =   720
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Zonas"
      Height          =   360
      Left            =   9960
      TabIndex        =   18
      Top             =   120
      Width           =   990
   End
   Begin VB.Timer NPC_AI 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   3240
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5640
      Top             =   3240
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   5160
      Top             =   3720
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6600
      Top             =   3720
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   6600
      Top             =   3240
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   3720
   End
   Begin VB.Timer packetResend 
      Interval        =   1
      Left            =   7080
      Top             =   3240
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   7080
      Top             =   3720
   End
   Begin VB.Timer TimerMinuto 
      Interval        =   60000
      Left            =   6120
      Top             =   3720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00202020&
      Caption         =   "Mensaje a todos los Jugadores"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4935
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar por Pop-Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Enviar por Consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00202020&
      Caption         =   "Estado del mundo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5160
      TabIndex        =   0
      Top             =   1680
      Width           =   4575
      Begin VB.Label Horario 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorteando Clima"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Clima 
         BackStyle       =   0  'Transparent
         Caption         =   "Soleado en todo el mundo."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4575
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7920
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskListen 
      Left            =   8760
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskClient 
      Index           =   0
      Left            =   9360
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   2850
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   240
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   5027
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":1042
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Poseedor del GP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label QuienGP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nadie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Iniciado en el port:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label txtPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label CantUsuarios 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuarios jugando:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tú IP en Inet:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5280
      TabIndex        =   4
      Top             =   225
      Width           =   1575
   End
   Begin VB.Label txtIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      ToolTipText     =   "Click para copiar al portapapeles"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReiniciar 
         Caption         =   "&Reiniciar Servidor"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApagarYGuardar 
         Caption         =   "Apagar Y Guardar"
      End
      Begin VB.Menu mnuApagarsinGuardar 
         Caption         =   "Apagar Sin Guardar"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "Opciones"
      Begin VB.Menu mnuDump 
         Caption         =   "&Dump - Guardar log de momento critico"
      End
      Begin VB.Menu mnusortGran 
         Caption         =   "Sortear Gran Poder"
      End
      Begin VB.Menu mnuMuertesubita 
         Caption         =   "Iniciar Muerte Subita"
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Activar/Desactivar el chat global"
      End
   End
   Begin VB.Menu mnuClimax 
      Caption         =   "Clima"
      Begin VB.Menu mnuClima 
         Caption         =   "Lluvia"
         Index           =   4
      End
      Begin VB.Menu mnuClima 
         Caption         =   "Nieve"
         Index           =   5
      End
      Begin VB.Menu mnuClima 
         Caption         =   "Niebla"
         Index           =   6
      End
      Begin VB.Menu mnuClima 
         Caption         =   "Niebla + Lluvia"
         Index           =   7
      End
   End
   Begin VB.Menu mnusave 
      Caption         =   "Guardar"
      Begin VB.Menu mnuWorldSave 
         Caption         =   "...WorldSave"
      End
      Begin VB.Menu mnuSavePersonajes 
         Caption         =   "...Personajes"
      End
   End
   Begin VB.Menu mnuEventos 
      Caption         =   "Eventos"
      Begin VB.Menu mnuSaqueador 
         Caption         =   "Saqueador"
         Begin VB.Menu mnuSpawnSaqueador 
            Caption         =   "Spawn"
         End
         Begin VB.Menu mnuMatarSaqueador 
            Caption         =   "Matar"
         End
      End
   End
   Begin VB.Menu mnuActualizar 
      Caption         =   "Actualizar"
      Begin VB.Menu mnuIntervalos 
         Caption         =   "Intervalos"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, id As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = id
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To MaxUsers
       'Conexion activa? y es un usuario loggeado?
       If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
            If UserList(iUserIndex).flags.Makro = 0 Then
                'Actualiza el contador de inactividad
                UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
                If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
                    Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado..")
                    'mato los comercios seguros
                    If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                        If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                            If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                                Call WriteConsoleMsg(UserList(iUserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                                Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                                Call FlushBuffer(UserList(iUserIndex).ComUsu.DestUsu) 'flush the buffer to send the message right away
                            End If
                        End If
                        Call FinComerciarUsu(iUserIndex)
                    End If
                    Call Cerrar_Usuario(iUserIndex)
                End If
            End If
        End If
    Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand
Static centinelSecs As Byte

centinelSecs = centinelSecs + 1

If centinelSecs = 5 Then
    'Every 5 seconds, we try to call the player's attention so it will report the code.
    Call modCentinela.CallUserAttention
    
    centinelSecs = 0
End If

Call PasarSegundo 'sistema de desconexion de 10 segs

Call ActualizaEstadisticasWeb

Exit Sub

errhand:

Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
Resume Next

End Sub

Private Sub AutoSave_Timer()
  On Error GoTo errHandler
   
  'fired every minute
  Static Minutos As Long
  Static MinutosLatsClean As Long
  Static MinsPjesSave As Long
  Static MinutosPoder As Byte
        
  Dim i As Integer
  Dim num As Long
  
  Minutos = Minutos + 1
    
  'Actualizamos el centinela
  Call modCentinela.PasarMinutoCentinela
    
    If Minutos = MinutosWs - 1 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto...", FontTypeNames.FONTTYPE_INFOBOLD))
    End If
    
    If Minutos >= MinutosWs Then
        Call GuardarUsuarios
        Call DoBackUp
        Call aClon.VaciarColeccion
        Minutos = 0
    End If
    
    If MinutosLatsClean = 1 Then
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpieza de Mundo en 1 minuto...", FontTypeNames.FONTTYPE_INFOBOLD))
    End If
    
    If MinutosLatsClean >= 60 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        'Call LimpiarMundo
    Else
        MinutosLatsClean = MinutosLatsClean + 1
    End If
    
    Call PurgarPenas
    Call CheckIdleUser
    
    '<<<<<-------- Log the number of users online ------>>>
    Dim n As Integer
    n = FreeFile()
    Open App.Path & "\logs\numusers.log" For Output Shared As n
    Print #n, NumUsers
    Close #n
    '<<<<<-------- Log the number of users online ------>>>
    
    '***************************************
    'Gran Poder
    
    If GranPoder > 0 Then
        If MapInfo(UserList(GranPoder).Pos.Map).Pk = False Then
           MinutosPoder = MinutosPoder + 1
            If MinutosPoder > 3 Then
                Call WriteConsoleMsg(GranPoder, "Perdiste el Gran Poder.", FontTypeNames.FONTTYPE_GUILD)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(GranPoder).Name & " ha perdido el Gran Poder", FontTypeNames.FONTTYPE_WARNING))
                Call OtorgarGranPoder(0)
                MinutosPoder = 0
            Else
                Call WriteConsoleMsg(GranPoder, "Estás en zona segura, te quedan " & 4 - MinutosPoder & " minutos para perder el Gran Poder.", FontTypeNames.FONTTYPE_GUILD)
            End If
        End If
    End If
    
    '17/02/2016 Lorwik: Comentado hasta no arreglar el error de la SQL al estar la DB vacia.
    Call PuntuarCastillos
    
    Exit Sub
errHandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)
    Resume Next
End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
Call AddtoRichTextBox(frmMain.RecTxt, Hour(Now) & ":" & Minute(Now) & " - Pop-Up> " & BroadMsg.Text)
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
Call AddtoRichTextBox(frmMain.RecTxt, Hour(Now) & ":" & Minute(Now) & " - Servidor> " & BroadMsg.Text)
End Sub

Private Sub Command3_Click()
  frmZonas.Show
End Sub

Private Sub mnuClima_Click(Index As Integer)
    Call Mod_Clima.SortearClima(0, True, Index)
End Sub

Private Sub TimerEfectos_Timer()

  'efectos
  Call Drag_Efectos.Procesar_Efectos_Jugador
  Call Drag_Efectos.Procesar_Efectos_NPC
  Call Drag_Efectos.Procesar_Efectos_Suelo
  
  'zonas
  Call Drag_Zonas.duracion_zonas
  
  

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Save stats!!!
Call Statistics.DumpStatistics

Call QuitarIconoSystray

#If SocketType = 1 Then
    Call LimpiaWsApi
#ElseIf SocketType = 2 Then
    wskListen.Close
#End If

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & time & " server cerrado."
Close #n

End

Set SonidosMapas = Nothing

End Sub

Private Sub GameTimer_Timer()
    Dim iUserIndex As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS As Boolean
    
On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    'Lorwik> Vamos a ver, paremosno aqui ¿Enserio que vamos a meter un For que vaya del 1 al MaxUser? ¿¡Estamos locos o que!? _
    Yo veo mejor que vaya del 1 hasta el ultimo usuario...
    For iUserIndex = 1 To LastUser
        With UserList(iUserIndex)
           'Conexion activa?
           If .ConnID <> -1 Then
                '¿User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    
                    Call DoTileEvents(iUserIndex, .Pos.Map, .Pos.X, .Pos.Y)
                    
                    If TimerEventoRinkel > 0 Then Call modArenaRinkel.RestarTimerEvento
                    If TimerRondaEventoRinkel > 0 Then Call modArenaRinkel.RestarTimerRonda
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    If .flags.Morph > 0 Then Call EfectoMorphUser(iUserIndex)
                    
                    If .flags.Muerto = 0 Then
                        
                        '[Consejeros]
                        If (.flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex)
                        If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)
                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex)
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                        End If
                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                        
                        If .flags.Makro <> 0 Then
                            .Counters.Makro = .Counters.Makro + 1
                            If .Counters.Makro >= IntervaloPuedeMakrear Then
                                .Counters.Makro = 0
                                MakroTrabajo iUserIndex, .flags.Makro
                            End If
                        End If
                        
                        Call DuracionPociones(iUserIndex)
                        
                        Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If Not .flags.Descansar Then
                            'No esta descansando
                                    
                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateHP(iUserIndex)
                                    bEnviarStats = False
                                End If
                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateSta(iUserIndex)
                                    bEnviarStats = False
                                End If
                                    
                                Else
                                'esta descansando
                                    
                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateHP(iUserIndex)
                                    bEnviarStats = False
                                End If
                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateSta(iUserIndex)
                                    bEnviarStats = False
                                End If
                                'termina de descansar automaticamente
                                If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                    Call WriteRestOK(iUserIndex)
                                    Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.Descansar = False
                                End If
                                    
                            End If
                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                    End If 'Muerto
                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)
                    End If
                End If 'UserLogged
                
                'If there is anything to be sent, we send it
                Call FlushBuffer(iUserIndex)
            End If
        End With
    Next iUserIndex
Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
End Sub

Private Sub mnuApagarSinGuardar_Click()
If MsgBox("¡ATENCIÓN!" & vbCrLf & "El servidor se cerrará SIN GUARDAR los cambios." & vbCrLf & "¿Desea hacerlo de todas maneras?", vbExclamation + vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
    
    #If SocketType = 1 Then
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    Call LimpiaWsApi ' GSZAO, cerramos los sockets con seguridad...
    #Else
    wskListen.Close
    Dim i As Long
    For i = 1 To LastUser
        Call wskClient(i).Close
    Next i
    #End If
End If
End Sub


Private Sub mnuApagarYGuardar_Click()
If MsgBox("¿Está seguro que desea hacer un WorldSave, guardar los personajes y apagar el servidor?", vbYesNo, "Apagar Magicamente") = vbYes Then
    Me.MousePointer = 11
    frmCargando.Show
    'WorldSave
    Call ES.DoBackUp
    'commit experiencia
    'Call mdParty.ActualizaExperiencias
    'Guardar Pjs
    Call GuardarUsuarios
    'Chauuu
    Unload frmMain
    #If SocketType = 1 Then
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    Call LimpiaWsApi ' GSZAO, cerramos los sockets con seguridad...
    #Else
    wskListen.Close
    Dim i As Long
    For i = 1 To LastUser
        Call wskClient(i).Close
    Next i
    #End If
End If
End Sub

Private Sub mnuDump_Click()
On Error Resume Next

    Dim i As Integer
    For i = 1 To MaxUsers
        Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & _
            ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & _
            " UserLogged: " & UserList(i).flags.UserLogged)
    Next i
    
    Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub mnuGlobal_Click()
  If HayGlobal = False Then
        HayGlobal = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El chat Global ha sido activado. ¡Se ruega no abusar, podria llevar a sanción!", FontTypeNames.FONTTYPE_INFOBOLD))
    Else
        HayGlobal = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El chat Global ha sido desactivado.", FontTypeNames.FONTTYPE_INFOBOLD))
    End If
End Sub

Private Sub mnuMatarSaqueador_Click()
    If SaqueadorIndex = 0 Then MsgBox "No hay ningún saqueador.": Exit Sub
    Call Saqueador(False)
End Sub

Private Sub mnuReiniciar_Click()
If MsgBox("¡ATENCIÓN!" & vbCrLf & "Si reinicia el servidor puede provocar la pérdida de datos de los usarios." & vbCrLf & "¿Desea reiniciar el servidor de todas maneras?", vbYesNo + vbCritical) = vbYes Then
    frmMain.txStatus.Text = "Reiniciando Servidor!"
    Call General.Restart
End If
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
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
        If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
    End If

End Sub

Private Sub mnuSavePersonajes_Click()
Me.MousePointer = 11
'Call mdParty.ActualizaExperiencias
Call GuardarUsuarios
Me.MousePointer = 0
MsgBox "Grabado de personajes OK!"
End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnusortGran_Click()
Call OtorgarGranPoder(0)
End Sub

Private Sub mnuSpawnSaqueador_Click()
    If SaqueadorIndex > 0 Then MsgBox "No puede haber mas de 1 saqueador.": Exit Sub
    Call modTrampas.Saqueador(True)
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
 For i = 1 To LastUser
        With UserList(i)
        
            Call FlushBuffer(i)
        
        End With
        
        Next
End Sub

Private Sub mnuWorldSave_Click()
On Error GoTo eh
    Me.MousePointer = 11
    FrmStat.Show
    Call DoBackUp
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
Exit Sub
eh:
Call LogError("Error en WORLDSAVE")
End Sub

Private Sub NPC_AI_Timer()
   ' On Error GoTo ErrorHandler
    Dim i As Long
    Dim X As Integer
    Dim Y As Integer
    Dim UseAI As Integer
    Dim Mapa As Integer
    Dim e_p As Integer
    
    ' Si no se está haciendo backup y no está en pausa el servidor
    If Not HaciendoBackup And Not EnPausa Then
      For i = 1 To LastNPC
        If NPCList(i).flags.Active Then
               
          If NPCList(i).flags.Paralizado = 1 Or NPCList(i).flags.Inmovilizado = 1 Then
            Call EfectoParalisisNpc(i)
          End If
          
          Mapa = NPCList(i).Pos.Map
              
          If Mapa > 0 Then
            If MapInfo(Mapa).NumUsers > 0 Then
              If NPCList(i).Movement <> TipoAI.Estatico Then
                Call NPCAI(i)
              ElseIf iniAutoSacerdote Then
                If NPCList(i).NPCType = eNPCType.ResucitadorNewbie Or NPCList(i).NPCType = eNPCType.Revividor Then
                  Call NpcAutoSacerdote(i)
                End If
              End If
            End If
          End If
      End If
      Next i
    End If

'ErrorHandler:
    
    'Debug.Print time(); "Error en NPC_AI_Timer"; i; LastNPC; Err.Description; ""
    
    'Call LogError("Error en NPC_AI_Timer " & NPCList(i).Name & " Mapa:" & NPCList(i).Pos.Map)
    'Call MuereNpc(i, 0)
End Sub

Private Sub txtIP_Click()
    Call Clipboard.SetText(txtIP.Caption)
    frmMain.txStatus.Text = "Dirección IP copiada."
End Sub

Private Sub npcataca_Timer()

On Error Resume Next
Dim npc As Long

For npc = 1 To LastNPC
    NPCList(npc).CanAttack = 1
Next npc

End Sub

Private Sub packetResend_Timer()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Attempts to resend to the user all data that may be enqueued.
'***************************************************
On Error GoTo errHandler:
    Dim i As Long
    
    For i = 1 To MaxUsers
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
            End If
        End If
    Next i

Exit Sub

errHandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.Description)
    Resume Next
End Sub


Private Sub TimerMinuto_Timer()
    Static minuto As Integer
    Static minutoGP As Byte
    Static MinutoSub As Byte
    Dim i As Long
    Dim Posi As worldPos
    
    If minuto = 60 Then minuto = 0 'Si el minuto llega a 1h se resetea
    minuto = minuto + 1 'Es un contado para ciertos eventos que se produciran en este Timer
    
    'Gran Poder
    minutoGP = minutoGP + 1
    If minutoGP = 2 Then
        If GranPoder = 0 Then
            Call OtorgarGranPoder(0)
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(GranPoder).Name & " es el poseedor del Gran Poder en " & MapInfo(UserList(GranPoder).Pos.Map).MapName & "(Mapa " & UserList(GranPoder).Pos.Map & ")", FontTypeNames.FONTTYPE_DIOS))
            Call SendData(SendTarget.ToPCArea, GranPoder, PrepareMessageCreateFX(UserList(GranPoder).Char.CharIndex, FxGranPoder, 0))
            frmMain.QuienGP.Caption = UserList(GranPoder).Name
        End If
        minutoGP = 0
    End If
    
    'Inet1.OpenURL ("http://www.aodrag.es/heartbeat.php")
    
'*********************************************************************
'Aqui se va a comprobar la hora actual del host y cambiaria el clima.
    Call SortearHorario
    Call SortearClima(minuto)
'**********************************************************************

'Mensajes automaticos
    For i = 1 To LastUser
        Select Case minuto
            Case 15
                Call WriteConsoleMsg(i, "¡Si tienes alguna duda o problema, no olvides consultar el manual antes de contactar con el soporte: www.aodrag.es/wiki/!", FontTypeNames.FONTTYPE_global)
            Case 25
                Call WriteConsoleMsg(i, "¿Te gustaría tener la primera montura de la versión 8.0? Sube a nivel 40 o más antes del 1 de Abril (incluido) y recibirás una montura de Dragón Dorado cuando empiece la Temporada 2, que dará comienzo a finales de Abril. IMPORTANTE: solamente 1 montura por cuenta, las monturas no serán comerciables ni se podrán traspasar entre personajes.", FontTypeNames.FONTTYPE_global)
            Case 30
                Call WriteConsoleMsg(i, "¡Consigue DragCreditos para obtener fabulosos premios! Muy pronto estarán disponibles.", FontTypeNames.FONTTYPE_global)
            Case 45
                Call WriteConsoleMsg(i, "¡Si encuentras algun bug, no olvides reportarlo, el aprovechamiento del mismo puede llevar a la sanción!", FontTypeNames.FONTTYPE_global)
            Case 60
                Call WriteConsoleMsg(i, "¡Puedes contactar con nosotros o con los demas jugadores a traves de nuestro foro: https://www.aodrag.es/foro/ o desde Facebook: https://www.facebook.com/aodrag!", FontTypeNames.FONTTYPE_global)
        End Select
    Next i
    
    'Controla el retardo de Spawn
    For i = 500 To TotalNPCDat
        If i = RetardoSpawn(i).NPCNUM Then
            If RetardoSpawn(i).Tiempo > 0 Then
                RetardoSpawn(i).Tiempo = RetardoSpawn(i).Tiempo - 1
            ElseIf RetardoSpawn(i).Tiempo = 0 Then
                Posi.Map = RetardoSpawn(i).Mapa
                Posi.X = RetardoSpawn(i).X
                Posi.Y = RetardoSpawn(i).Y
                
                Call SpawnNpc(i, Posi, False, False, True)
                
                'Reseteamos:
                RetardoSpawn(i).Tiempo = 0
                RetardoSpawn(i).Mapa = 0
                RetardoSpawn(i).X = 0
                RetardoSpawn(i).Y = 0
                RetardoSpawn(i).NPCNUM = 0
            End If
        End If
    Next i

End Sub

Private Sub tPiqueteC_Timer()
    Dim NuevaA As Boolean
    Dim NuevoL As Boolean
    Dim GI As Integer
    
    Dim i As Long
    
On Error GoTo errHandler
    '18/02/2016 Lorwik> Aprovecho este Timer para meter el FX del GP cada 6seg
    If GranPoder > 0 Then _
    Call SendData(SendTarget.ToPCArea, GranPoder, PrepareMessageCreateFX(UserList(GranPoder).Char.CharIndex, FxGranPoder, 0))

    For i = 1 To LastUser
        With UserList(i)
            If .flags.UserLogged Then
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                    .Counters.PiqueteC = .Counters.PiqueteC + 1
                    Call WriteMultiMessage(i, eMessages.Piquete)
                    
                    If .Counters.PiqueteC > 23 Then
                        .Counters.PiqueteC = 0
                        Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                    End If
                Else
                    .Counters.PiqueteC = 0
                End If
                
                Call FlushBuffer(i)
            End If
        End With
    Next i
Exit Sub

errHandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)
End Sub

#If SocketType = 2 Then
Private Sub wskClient_Close(Index As Integer)
    Call CloseSocketSL(Index)
    Call Cerrar_Usuario(Index)
End Sub

Private Sub wskClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim data() As Byte
    
    ReDim data(bytesTotal) As Byte
    
    wskClient(Index).GetData data, , bytesTotal
    EventoSockRead Index, data
End Sub

Private Sub wskClient_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call LogCriticEvent("Winsock Error: " & Index & " Desc: " & Description)
    Call CloseSocketSL(Index)
    Call Cerrar_Usuario(Index)
End Sub
Private Sub wskListen_Close()
    Restart_ListenSocket
End Sub
Private Sub wskListen_ConnectionRequest(ByVal requestID As Long)
On Error GoTo Err:
    Dim NewIndex As Integer
    NewIndex = NextOpenUser

    Call wskClient(NewIndex).accept(requestID)
    
    Call TCP.Socket_NewConnection(NewIndex, wskClient(NewIndex).RemoteHostIP, wskClient(NewIndex).SocketHandle)
Exit Sub
Err:
    Restart_ListenSocket
    If NewIndex <> 0 Then Call wskClient(NewIndex).Close
End Sub
Private Sub Restart_ListenSocket()
    wskListen.Close
    wskListen.LocalPort = Puerto
    wskListen.listen
End Sub
#End If
