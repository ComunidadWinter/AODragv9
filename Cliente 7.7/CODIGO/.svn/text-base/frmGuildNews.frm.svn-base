VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox aliados 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
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
      Height          =   870
      ItemData        =   "frmGuildNews.frx":0000
      Left            =   720
      List            =   "frmGuildNews.frx":0002
      TabIndex        =   2
      Top             =   5280
      Width           =   4695
   End
   Begin VB.ListBox guerra 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
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
      Height          =   870
      ItemData        =   "frmGuildNews.frx":0004
      Left            =   720
      List            =   "frmGuildNews.frx":0006
      TabIndex        =   1
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox news 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
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
      Height          =   2175
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2160
      MouseIcon       =   "frmGuildNews.frx":0008
      MousePointer    =   99  'Custom
      Picture         =   "frmGuildNews.frx":0CD2
      Top             =   6480
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Aliados"
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
      Index           =   3
      Left            =   720
      TabIndex        =   6
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Enemigos"
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
      Index           =   2
      Left            =   720
      TabIndex        =   5
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Noticias Del Clan"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Noticias Del Clan"
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
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmGuildNews.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
End Sub
Private Sub Command1_Click()

End Sub

Public Sub ParseGuildNews(ByVal s As String)

news = Replace(ReadField(1, s, Asc("¬")), "º", vbCrLf)

Dim h%, j%

h% = Val(ReadField(2, s, Asc("¬")))

For j% = 1 To h%
    
    guerra.AddItem ReadField(j% + 2, s, Asc("¬"))
    
Next j%

j% = j% + 2

h% = Val(ReadField(j%, s, Asc("¬")))

For j% = j% + 1 To j% + h%
    
    aliados.AddItem ReadField(j%, s, Asc("¬"))
    
Next j%

Me.Show vbModeless, frmMain

End Sub

Private Sub Image1_Click()
On Error Resume Next
Unload Me
'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
End Sub
