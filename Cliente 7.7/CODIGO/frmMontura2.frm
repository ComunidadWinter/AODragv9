VERSION 5.00
Begin VB.Form frmMontura2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   7185
   ClientLeft      =   2730
   ClientTop       =   285
   ClientWidth     =   6135
   LinkTopic       =   "Form4"
   Picture         =   "frmMontura2.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4080
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   480
      MouseIcon       =   "frmMontura2.frx":5FAB
      MousePointer    =   99  'Custom
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   510
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   3720
      MouseIcon       =   "frmMontura2.frx":6C75
      MousePointer    =   99  'Custom
      Picture         =   "frmMontura2.frx":793F
      Top             =   2880
      Width           =   2220
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   4200
      MouseIcon       =   "frmMontura2.frx":C692
      MousePointer    =   99  'Custom
      Picture         =   "frmMontura2.frx":D35C
      Top             =   6705
      Width           =   1650
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   4080
      MouseIcon       =   "frmMontura2.frx":1157C
      MousePointer    =   99  'Custom
      Picture         =   "frmMontura2.frx":12246
      Top             =   4200
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   3780
      MouseIcon       =   "frmMontura2.frx":165A6
      MousePointer    =   99  'Custom
      Picture         =   "frmMontura2.frx":17270
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   38
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   37
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   36
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   35
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   34
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   33
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   32
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asignar"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   360
      MouseIcon       =   "frmMontura2.frx":1BF40
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asignar"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   360
      MouseIcon       =   "frmMontura2.frx":1CC0A
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asignar"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   360
      MouseIcon       =   "frmMontura2.frx":1D8D4
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asignar"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   360
      MouseIcon       =   "frmMontura2.frx":1E59E
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asignar"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   360
      MouseIcon       =   "frmMontura2.frx":1F268
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asignar"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   360
      MouseIcon       =   "frmMontura2.frx":1FF32
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asignar"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      MouseIcon       =   "frmMontura2.frx":20BFC
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label libres 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Puntos Asignar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Nombre1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Exp1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label golpe1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label vida1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label exp2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label nivel1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label descrip 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   1320
      TabIndex        =   15
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label habi1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Def:Cuerpo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label habi2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "At.Cuerpo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label habi3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "At.Flechas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label habi4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Def.Flechas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label habi5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "At.Magias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label habi6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Def.Magias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label habi7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Evasión:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Nivel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nivel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label elu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exp.Prox.Nivel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label vida 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label golpe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Golpe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Experiencia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "frmMontura2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cambiar_Click()

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub eva_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Image1_Click()
If Val(nivel1.Caption) < 1 Then Exit Sub
Text1.Visible = True
Image2.Visible = True
End Sub

Private Sub Image2_Click()
'pluto:2.5.0
If Not AsciiValidos(Text1.Text) Then
MsgBox "Nombre con caracteres invalidos."
Exit Sub
End If
If Len(Text1.Text) > 8 Or Len(Text1.Text) < 3 Then
MsgBox "El Nombre debe tener un Mínimo 3 letras y un Máximo 8 letras."
Exit Sub
End If

Nombre1.Caption = Text1.Text
Image1.Visible = True
Text1.Visible = False
Image2.Visible = False
SendData ("NMAS" & Text1.Text & "," & SELECI)
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Image4_Click()
variable = "http://www.juegosdrag.es/aomanual/?sec=mascotas#ir"

Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.Navigate variable
Call AddtoRichTextBox(frmMain.RecTxt, "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, 0, 0, True, False, False)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label4_Click(Index As Integer)
Call audio.PlayWave(SND_CLICK)
Dim PLibre As Byte
Dim Pana As Byte
Dim n As Byte
Dim HAYTOPE As Boolean

PLibre = Val(libres.Caption)
Pana = Val(Label5(Index).Caption)

If PLibre <= 0 Then Exit Sub

For n = 0 To 6
If Pana > Val(Label5(n).Caption) + 3 And Label4(n).Visible = True Then
Select Case n
Case 0
If Val(Label5(n).Caption) < PMascotas(SELECI).TopeAtCuerpo Then Exit Sub
Case 1
If Val(Label5(n).Caption) < PMascotas(SELECI).TopeDefCuerpo Then Exit Sub
Case 2
If Val(Label5(n).Caption) < PMascotas(SELECI).TopeAtFlechas Then Exit Sub
Case 3
If Val(Label5(n).Caption) < PMascotas(SELECI).TopeDefFlechas Then Exit Sub
Case 4
If Val(Label5(n).Caption) < PMascotas(SELECI).TopeAtMagico Then Exit Sub
Case 5
If Val(Label5(n).Caption) < PMascotas(SELECI).TopeDefMagico Then Exit Sub
Case 6
If Val(Label5(n).Caption) < PMascotas(SELECI).TopeEvasion Then Exit Sub
End Select
End If
Next




Select Case Index
Case 0
If Pana >= PMascotas(SELECI).TopeAtCuerpo Then Exit Sub
Case 1
If Pana >= PMascotas(SELECI).TopeDefCuerpo Then Exit Sub
Case 2
If Pana >= PMascotas(SELECI).TopeAtFlechas Then Exit Sub
Case 3
If Pana >= PMascotas(SELECI).TopeDefFlechas Then Exit Sub
Case 4
If Pana >= PMascotas(SELECI).TopeAtMagico Then Exit Sub
Case 5
If Pana >= PMascotas(SELECI).TopeDefMagico Then Exit Sub
Case 6
If Pana >= PMascotas(SELECI).TopeEvasion Then Exit Sub
End Select

PLibre = PLibre - 1
Pana = Pana + 1
libres.Caption = PLibre
Label5(Index).Caption = "+" & Pana

If PLibre <= 0 Then
    For n = 0 To 6
    Label4(n).Visible = False
    Next n
End If

'ENVIAMOS EL ASIGNAR
SendData ("LIX" & Index & "," & UserIndex & "," & SELECI)


End Sub

Private Sub Text1_Change()
Text1.Text = LTrim(Text1.Text)
End Sub
