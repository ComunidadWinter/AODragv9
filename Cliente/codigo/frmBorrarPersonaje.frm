VERSION 5.00
Begin VB.Form frmBorrarPersonaje 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5895
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   4200
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtBORRAR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblEstasSeguro 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBorrarPersonaje.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   5445
   End
   Begin VB.Label lblAtenciónVas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡Atención! vas a borrar el personaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lblPersonaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "frmBorrarPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If Not txtBORRAR.Text = "BORRAR" Then
        MsgBox "¡Tienes que escribir ""BORRAR"" para eliminar el personaje!"
        Exit Sub
    End If

    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
            
    EstadoLogin = E_MODO.BorrandoPJ
            
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
    
    Unload Me
End Sub

Private Sub cmdVolver_Click()
    lblPersonaje.Caption = "-"
    Unload Me
End Sub

Private Sub Form_Load()
    lblPersonaje.Caption = frmCuenta.ListPJ.List(frmCuenta.ListPJ.ListIndex)
End Sub
