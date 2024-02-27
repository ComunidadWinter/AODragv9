VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   0  'None
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   6165
   ClipControls    =   0   'False
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
   ScaleHeight     =   7155
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstArmaduras 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   5400
   End
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4650
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.TextBox TxTCant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Text            =   "1"
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Con tu energia actual puedes construir hasta 0 unidades"
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
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   6480
      Width           =   5535
   End
   Begin VB.Image cmdSalir 
      Height          =   300
      Left            =   3480
      MouseIcon       =   "frmHerrero.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":0CCA
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Image cmdConstruir 
      Height          =   300
      Left            =   720
      MouseIcon       =   "frmHerrero.frx":4E86
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":5B50
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Construcción de Objetos"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.Image Command1 
      Height          =   300
      Left            =   360
      MouseIcon       =   "frmHerrero.frx":6178
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":6E42
      Top             =   840
      Width           =   1650
   End
   Begin VB.Image Command2 
      Height          =   300
      Left            =   2040
      MouseIcon       =   "frmHerrero.frx":AF12
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":BBDC
      Top             =   840
      Width           =   1650
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConstruir_Click()
On Error Resume Next
Dim cantidad As String

    cantidad = txtCant.Text
    
    If Not IsNumeric(cantidad) Or Int(Val(cantidad)) < 1 Or Int(Val(cantidad)) > 1000 Then
        MsgBox "La cantidad es invalida.", vbCritical
        Exit Sub
    End If

    If lstArmas.Visible Then
        Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1), cantidad)
    Else
        Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmaduras.ListIndex + 1), cantidad)
    End If

    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    lstArmaduras.Visible = False
    lstArmas.Visible = True
End Sub

Private Sub Command2_Click()
    lstArmaduras.Visible = True
    lstArmas.Visible = False
End Sub

Private Sub Form_Load()
    lblEnergia.Caption = "Con tu energia maxima actual puedes construir hasta " & Round(UserMaxSTA / 2) & " unidades."
    Me.Picture = General_Load_Picture_From_Resource("52.gif")
End Sub

