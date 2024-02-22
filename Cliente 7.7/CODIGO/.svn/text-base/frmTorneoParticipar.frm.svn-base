VERSION 5.00
Begin VB.Form frmTorneoParticipar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Participar en Torneo"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      Picture         =   "frmTorneoParticipar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Participar"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      MaskColor       =   &H00C0E0FF&
      Picture         =   "frmTorneoParticipar.frx":0929
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      DataField       =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   2460
      Left            =   120
      TabIndex        =   0
      Tag             =   "Inscritos al Torneo"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Torneo AoDraG"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   480
      TabIndex        =   14
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Restricción Nivel:"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2160
      TabIndex        =   10
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   4320
      TabIndex        =   9
      Top             =   4455
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Torneo:"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   3870
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Premio Ganador:"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   3285
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuota Inscripción:"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   3960
      TabIndex        =   3
      Top             =   2655
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Organizador:"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Inscritos al Torneo"
      BeginProperty Font 
         Name            =   "Liberate"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmTorneoParticipar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub load()

End Sub
Private Sub Command1_Click()
SendData ("TO3" & "," & UserName)

End Sub

Private Sub Command2_Click()
frmTorneoParticipar.Visible = False
SendData "/DTT"
End Sub

