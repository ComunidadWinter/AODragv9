VERSION 5.00
Begin VB.Form frmPartyReparto 
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   7200
   ClientLeft      =   2460
   ClientTop       =   570
   ClientWidth     =   6165
   ForeColor       =   &H00000000&
   LinkTopic       =   "Personalizar reparto"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmPartyReparto.frx":0CCA
      Top             =   6360
      Width           =   1680
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Libres:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      TabIndex        =   41
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   9
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":502A
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   8
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":5CF4
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   7
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":69BE
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   6
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":7688
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   5
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":8352
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   4
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":901C
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   3
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":9CE6
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   2
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":A9B0
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   1
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":B67A
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   9
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":C344
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   8
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":D00E
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   7
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":DCD8
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   6
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":E9A2
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   5
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":F66C
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   4
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":10336
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   3
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":11000
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   2
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":11CCA
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   1
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":12994
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   1560
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   1560
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   1560
      TabIndex        =   19
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   1560
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   1560
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   0
      Left            =   2160
      MouseIcon       =   "frmPartyReparto.frx":1365E
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   0
      Left            =   1200
      MouseIcon       =   "frmPartyReparto.frx":14328
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   2760
      TabIndex        =   9
      Top             =   5680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   2760
      TabIndex        =   8
      Top             =   5200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   2760
      TabIndex        =   7
      Top             =   4720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   2760
      TabIndex        =   6
      Top             =   4240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   3760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   4
      Top             =   3300
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   3
      Top             =   2800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   2
      Top             =   2320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1360
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmPartyReparto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private total As Integer
Private Sub Form_Load()
frmPartyReparto.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
    Dim pt1 As Byte
    total = 0
    For pt1 = 1 To MAXMIEMBROS
        If Party.Miembros(pt1).nombre = "0" Then
            frmPartyReparto.Label1(pt1 - 1).Caption = ""
        Else
            frmPartyReparto.Label1(pt1 - 1).Caption = Party.Miembros(pt1).nombre
            frmPartyReparto.Label1(pt1 - 1).Visible = True
            frmPartyReparto.Label3(pt1 - 1).Visible = True
            frmPartyReparto.Label4(pt1 - 1).Visible = True
            frmPartyReparto.Label5(pt1 - 1).Visible = True
        End If
        frmPartyReparto.Label5(pt1 - 1).Caption = Val(Party.Miembros(pt1).privi)
        total = total + Val(Party.Miembros(pt1).privi)
    Next
    frmPartyReparto.Label6.Caption = 100 - total
End Sub

Private Sub Image1_Click()
Dim i As Byte
Dim cad$
total = 0

    For i = 0 To 9
        cad$ = cad$ & Label5(i).Caption & ","
        Party.Miembros(i + 1).privi = Val(Label5(i).Caption)
        total = total + Val(Label5(i).Caption)
    'pluto:6.0A
     'If party.Miembros(i + 1).privi < 5 And party.Miembros(i + 1).Nombre <> "" Then
    'frmPartyReparto.Visible = False
    'Unload frmPartyReparto
     'MsgBox ("Raparto no Válido. Cada miembro debe tener un mínimo del 5%")
     'Exit Sub
    ' End If
    
    Next
    If total <= 100 Then
        Call SendData("PT6" & cad$)
    End If
    frmPartyReparto.Visible = False
    Unload frmPartyReparto
End Sub



Private Sub Label3_Click(Index As Integer)
If Val(frmPartyReparto.Label5(Index).Caption) > 5 Then
    frmPartyReparto.Label5(Index).Caption = (Val(frmPartyReparto.Label5(Index).Caption) - 1)
    frmPartyReparto.Label6.Caption = Val(frmPartyReparto.Label6.Caption) + 1
End If
End Sub

Private Sub Label4_Click(Index As Integer)
'pluto:6.3 añado tope 95%
If Val(frmPartyReparto.Label6.Caption) > 0 And Val(frmPartyReparto.Label5(Index).Caption) < 95 Then
    frmPartyReparto.Label5(Index).Caption = Val(frmPartyReparto.Label5(Index).Caption) + 1
    frmPartyReparto.Label6.Caption = (Val(frmPartyReparto.Label6.Caption) - 1)
End If
End Sub

