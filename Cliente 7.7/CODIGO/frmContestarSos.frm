VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmContestarSos 
   BorderStyle     =   0  'None
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox Text2 
      Height          =   1215
      Left            =   720
      TabIndex        =   6
      Top             =   5160
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2143
      _Version        =   393217
      BackColor       =   64
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmContestarSos.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1455
      Left            =   720
      TabIndex        =   5
      Top             =   3000
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   64
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmContestarSos.frx":0083
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000040&
      ForeColor       =   &H80000005&
      Height          =   1815
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Command2 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   4200
      TabIndex        =   8
      Top             =   6720
      Width           =   1600
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   1500
      Width           =   480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "fecha"
      Height          =   375
      Left            =   1905
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "fecha"
      Height          =   255
      Left            =   1900
      TabIndex        =   0
      Top             =   4830
      Width           =   1200
   End
End
Attribute VB_Name = "frmContestarSos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Const GWL_EXSTYLE = (-20)
'Const WS_EX_TRANSPARENT = &H20&

'Private Sub Form_Load()

'Dim result As Long

'frmContestarSos.Picture = LoadPicture(DirGraficos & "contestarsos.jpg")

'result = SetWindowLong(Text1.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
'result = SetWindowLong(Text2.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

'End Sub

'Private Sub Command1_Click()
'SendData ("SOS;" & UserName & ";" & Label6.Caption & ";" & Label3.Caption & ";" & Text1.Text & ";" & Text2.Text & ";")
'Unload Me
'End Sub

'Private Sub Command2_Click()
'Unload Me
'End Sub


'Private Sub List1_Click()

'Call BuscaMensaje(List1.List(List1.ListIndex))

'End Sub
