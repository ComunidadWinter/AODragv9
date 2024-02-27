VERSION 5.00
Begin VB.Form frmMSGM 
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5940
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
   ScaleHeight     =   7485
   ScaleWidth      =   5940
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   360
      Left            =   480
      TabIndex        =   13
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   360
      Left            =   3240
      TabIndex        =   12
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox mensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3480
      Width           =   5415
   End
   Begin VB.ComboBox categoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00111720&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      ItemData        =   "frmMSGM.frx":0000
      Left            =   2520
      List            =   "frmMSGM.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label lblCodes 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   5270
      Width           =   855
   End
   Begin VB.Label lblAviso 
      Caption         =   "Escribe el siguiente codigo de seguridad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Label Label7 
      Caption         =   "¡¡¡ RECUERDA QUE SI ENVIAS VARIAS VECES EL MISMO MENSAJE PODRIAS SER SANCIONADO !!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   6240
      Width           =   5295
   End
   Begin VB.Label Label6 
      Caption         =   "Antes de enviar tu mensaje, comprueba que todo este correctamente."
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5760
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "Mensaje:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   $"frmMSGM.frx":0080
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "Categoria del mensaje:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   5400
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "El Staff no atendera consultas que esten respondida en el manual del juego."
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5490
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMSGM.frx":014D
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmMSGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVolver_Click()
Unload Me
End Sub

Private Sub cmdEnviar_Click()
Dim GMs As String

    If categoria.ListIndex = -1 Then
        MsgBox "El motivo del mensaje no es válido"
        Exit Sub
    End If
    
    If Len(mensaje.Text) > 250 Then
        MsgBox "La longitud del mensaje debe tener menos de 250 carácteres."
        Exit Sub
    End If
    
    If Len(mensaje.Text) = 0 Then
        MsgBox "Debes ingresar un mensaje."
        Exit Sub
    End If
    
    If lblCodes.Caption <> txtCode.Text Then
        MsgBox "El Codigo ingresado es Invalido.", vbCritical
        lblCodes.Caption = GenerateKey
        Exit Sub
    End If

    
    Call WriteGMRequest(categoria.List(categoria.ListIndex), mensaje.Text)

    mensaje.Text = ""
    categoria.List(categoria.ListIndex) = ""
    AddtoRichTextBox frmMain.RecTxt, "El mensaje fue enviado. Rogamos tengas paciencia y no escribas más de un mensaje sobre el mismo tema, comprende que hay mas usuarios a la espera.", 252, 151, 53, 1, 0
    Unload Me

End Sub

Private Sub Form_Load()
lblCodes.Caption = GenerateKey
End Sub

Private Sub mensaje_Change()
mensaje.Text = LTrim(mensaje.Text)
End Sub


Private Sub mensaje_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 209) And (KeyAscii <> 241) And (KeyAscii <> 8) And (KeyAscii <> 32) And (KeyAscii <> 164) And (KeyAscii <> 165) Then
    If (Index <> 6) And ((KeyAscii < 40 Or KeyAscii > 122) Or (KeyAscii > 90 And KeyAscii < 96)) Then
        KeyAscii = 0
    End If
End If

 KeyAscii = Asc((Chr(KeyAscii)))
End Sub

