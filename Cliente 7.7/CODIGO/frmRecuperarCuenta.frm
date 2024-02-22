VERSION 5.00
Begin VB.Form frmRecuperarCuenta 
   BorderStyle     =   0  'None
   Caption         =   "Recuperar Cuenta"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000040&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRecuperarCuenta.frx":0000
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   600
      TabIndex        =   2
      Top             =   3720
      Width           =   5175
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3480
      MouseIcon       =   "frmRecuperarCuenta.frx":010F
      MousePointer    =   99  'Custom
      Picture         =   "frmRecuperarCuenta.frx":0DD9
      Top             =   5880
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   840
      MouseIcon       =   "frmRecuperarCuenta.frx":4F95
      MousePointer    =   99  'Custom
      Picture         =   "frmRecuperarCuenta.frx":5C5F
      Top             =   5880
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Email De La Cuenta:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   4335
   End
End
Attribute VB_Name = "frmRecuperarCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'pluto:2.8.0
    If PideClave = True Then
    MsgBox "Ya has solicitado la recuperación de la cuenta, no es necesario que lo hagas de nuevo."
    Exit Sub
    End If
    
    If Not CheckMailString(Text1) Then
            MsgBox "Direccion de mail invalida."
            Exit Sub
    End If
    
    PideClave = True
        'pluto:6.0A-------------------
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    '---------------------------
    frmMain.Socket1.Connect
End Sub

Private Sub Form_Load()
frmRecuperarCuenta.Picture = LoadPicture(App.Path & "\Graficos\ventanas.jpg")

End Sub
Private Sub Command2_Click()
    frmRecuperarCuenta.Visible = False
End Sub

Private Sub Image1_Click()
    'pluto:2.8.0
    If PideClave = True Then
    MsgBox "Ya has solicitado la recuperación de la cuenta, no es necesario que lo hagas de nuevo."
    Exit Sub
    End If
    
    If Not CheckMailString(Text1) Then
            MsgBox "Direccion de mail invalida."
            Exit Sub
    End If
    
    PideClave = True
        'pluto:6.0A-------------------
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    '---------------------------
    frmMain.Socket1.Connect
End Sub

Private Sub Image2_Click()
frmRecuperarCuenta.Visible = False
End Sub
