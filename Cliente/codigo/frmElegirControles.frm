VERSION 5.00
Begin VB.Form frmElegirControles 
   Caption         =   "Pre-configuración de controles"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6555
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
   ScaleHeight     =   4575
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMDClasico 
      Caption         =   "Seleccionar"
      Height          =   360
      Left            =   3720
      TabIndex        =   14
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton CmdWASD 
      Caption         =   "Seleccionar"
      Height          =   360
      Left            =   360
      TabIndex        =   13
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "D - Domar"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   12
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "A - Agarrar"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   11
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "U - Usar"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   10
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Control - Atacar"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   9
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Flechas - Teclas de dirección"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   8
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "P- Domar"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "G - Agarrar"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Barra Espacio - Usar"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "F - Atacar"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "W, A, S, D - Teclas de dirección"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   3360
      Y1              =   1200
      Y2              =   4320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Clasico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "W, A, S, D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "¡Ya casi estas listo para la aventura! ¿Con que configuración de controles quieres jugar? ¡Elige la que mejor la que mas te guste!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6240
   End
End
Attribute VB_Name = "frmElegirControles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDClasico_Click()
    Call CustomKeys.LoadDefaults(True)
    MsgBox "Has elegido la configuración ""Clasico"", recuerda que podras cambiar los controles desde el panel Opciones."
    Opciones.PrimeraVez = 0
    Unload Me
End Sub

Private Sub CmdWASD_Click()
    Call CustomKeys.LoadDefaults(False)
    MsgBox "Has elegido la configuración ""W,A,S,D"", recuerda que podras cambiar los controles desde el panel Opciones."
    Opciones.PrimeraVez = 0
    Unload Me
End Sub

