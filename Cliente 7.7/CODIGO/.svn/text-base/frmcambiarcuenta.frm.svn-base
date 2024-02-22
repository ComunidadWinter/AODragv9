VERSION 5.00
Begin VB.Form frmcambiarcuenta 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   ControlBox      =   0   'False
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3480
      MouseIcon       =   "frmcambiarcuenta.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmcambiarcuenta.frx":0CCA
      Top             =   5400
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1200
      MouseIcon       =   "frmcambiarcuenta.frx":4F83
      MousePointer    =   99  'Custom
      Picture         =   "frmcambiarcuenta.frx":5C4D
      Top             =   5400
      Width           =   1680
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmcambiarcuenta.frx":9FAD
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Introduce Cuenta de Correo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   3495
   End
End
Attribute VB_Name = "frmcambiarcuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
frmcambiarcuenta.Picture = LoadPicture(App.Path & "\Graficos\ventanas.jpg")
End Sub

Private Sub Image1_Click()

Dim a As String
a = frmCuentas.Cuentas.List(frmCuentas.Cuentas.ListIndex)
Call SendData("RPERSO" & a & "," & frmcambiarcuenta.Text1.Text)
frmcambiarcuenta.Visible = False
frmCuentas.Visible = False
frmConnect.Visible = True
'pluto:2.11
Sleep 1000
Call frmMain.Socket1.Disconnect
MsgBox "Compruebe si el personaje ha cambiado de cuenta. Si el personaje no cambia de cuenta puede deberse a las siguientes causas: La cuenta que indicó no existe - La cuenta que indicó está en uso en estos momentos - El personaje no ha sido creado en esta cuenta que estas usando y en ese caso sólo puedes enviarlo de vuelta a su email de creación, no puedes cambiarlo a ninguna otra cuenta. "

End Sub

Private Sub Image2_Click()
frmcambiarcuenta.Visible = False
frmCuentas.Visible = True
End Sub
