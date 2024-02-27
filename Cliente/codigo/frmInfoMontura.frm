VERSION 5.00
Begin VB.Form frmInfoMontura 
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   1080
      TabIndex        =   27
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Atributos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   4335
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar"
         Height          =   255
         Index           =   6
         Left            =   3240
         TabIndex        =   26
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar"
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   25
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Puntos Libres:"
         Height          =   195
         Index           =   9
         Left            =   480
         TabIndex        =   21
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Evasión"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Def.Magias:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "At.Magias:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Defensa:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ataque:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información básica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox TxTInfo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nivel"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Experiencia:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmInfoMontura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAsignar_Click(index As Integer)
    If TxTInfo(9).Text = 0 Then Exit Sub
    
    If MsgBox("¿Seguro que quieres asignar el skill en " & cmdAsignar(index).Caption & "?", vbYesNo, "Atencion!") = vbNo Then Exit Sub
    Call WriteSolInfMonturaClient(2, index)
End Sub

Private Sub cmdCerrar_Click()
Dim i As Byte
    For i = 0 To 9
        TxTInfo(i).Text = ""
    Next i
    MonturaSeleccionada = 0
    Unload Me
End Sub
