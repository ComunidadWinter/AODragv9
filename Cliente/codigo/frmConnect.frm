VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdActuallizar 
      DisabledPicture =   "frmConnect.frx":000C
      DownPicture     =   "frmConnect.frx":070E
      Height          =   375
      Left            =   3420
      Picture         =   "frmConnect.frx":0E10
      TabIndex        =   5
      Top             =   2550
      Width           =   375
   End
   Begin VB.CheckBox ChkRecordar 
      Height          =   195
      Left            =   495
      TabIndex        =   4
      Top             =   3090
      Width           =   195
   End
   Begin VB.ComboBox SVList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   420
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2580
      Width           =   2895
   End
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   420
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1620
      Width           =   3450
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   420
      TabIndex        =   0
      Top             =   930
      Width           =   3450
   End
   Begin VB.Image cmdSalir 
      Height          =   375
      Left            =   480
      Top             =   4110
      Width           =   1695
   End
   Begin VB.Image CmdOpciones 
      Height          =   375
      Left            =   2280
      Top             =   4110
      Width           =   1695
   End
   Begin VB.Image CMDConnect 
      Height          =   255
      Left            =   2280
      Top             =   3510
      Width           =   1695
   End
   Begin VB.Image CmdCrearPJ 
      Height          =   255
      Left            =   480
      Top             =   3510
      Width           =   1695
   End
   Begin VB.Label CmdRecPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¿Olvidaste tu contraseña?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   480
      TabIndex        =   2
      Top             =   1950
      Width           =   1950
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkRecordar_Click()
    If ChkRecordar.value = 0 Then
        Call WriteVar(App.Path & "\INIT\Recordar.dat", "Account", "Nombre", "")
        Call WriteVar(App.Path & "\INIT\Recordar.dat", "Account", "Check", "0")
        NameTxt.Text = ""
        PasswordTxt.Text = ""
    End If
End Sub

Private Sub Form_Load()
    Dim ServerPredefinido As Byte

    EngineRun = True
    Me.Picture = General_Load_Picture_From_Resource("29.gif")
    frmCuenta.ListPJ.Clear
    
    If GetVar(App.Path & "\INIT\AODConfig.bnd", "Account", "Check") = 1 Then
        ChkRecordar.value = 1
        NameTxt.Text = GetVar(App.Path & "\INIT\AODConfig.bnd", "Account", "Nombre")
    Else
        ChkRecordar.value = 0
    End If
    
    '***************************************************
    'LISTA DE SERVIDORES
    Call ListarServidores
    
    'ServerPredefinido = GetVar(App.Path & "\INIT\AODConfig.bnd", "EXTRAS", "SERVER")
    'If ServerPredefinido = 0 Then ServerPredefinido = 1
   ' SVList.ListIndex = ServerPredefinido
    '***************************************************
End Sub

Private Sub cmdActuallizar_Click()
    Call ListarServidores
End Sub

Private Sub CMDConnect_Click()
    Call Sound.Sound_Play(SND_CLICK)
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
  
    'update user info
    Cuenta.name = NameTxt.Text
    Cuenta.Pass = PasswordTxt.Text
    
    If ChkRecordar.value = 1 Then
        Call WriteVar(App.Path & "\INIT\AODConfig.bnd", "Account", "Nombre", NameTxt.Text)
        Call WriteVar(App.Path & "\INIT\AODConfig.bnd", "Account", "Check", "1")
    End If
    
    'Guardamos el ultimo servidor elegido
    Call WriteVar(App.Path & "\INIT\AODConfig.bnd", "EXTRAS", "SERVER", SVList.ListIndex)
    
    If CheckUserData(False, True) = True Then
        EstadoLogin = LoginCuenta
        Call ChangeCursorMain(cur_Wait, frmConnect)
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
    End If
End Sub

Private Sub cmdcrearpj_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    Call ShellExecute(0, "Open", "http://www.aodrag.es/cuenta/crear", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub CmdOpciones_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call frmOpciones.Init
End Sub

Private Sub CmdRecPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CmdRecPass.Font.Underline = True
End Sub

Private Sub cmdSalir_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call CloseClient
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CmdRecPass.Font.Underline = True Then CmdRecPass.Font.Underline = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
        frmCargando.Status.Caption = "Cerrando AODrag."
        frmCargando.Refresh
        
        prgRun = False
        
        frmCargando.Status.Caption = "¡¡Gracias por jugar AODrag!!"
        frmCargando.Refresh
        Call UnloadAllForms
End If
End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then CMDConnect_Click
End Sub

Private Sub SVList_Click()
    CurServerIp = Servidor(SVList.ListIndex).Ip
    CurServerPort = Servidor(SVList.ListIndex).Puerto
    ServIndSel = SVList.ListIndex + 1
End Sub
