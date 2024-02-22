VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBandejaNormas 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3855
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6800
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmBandejaNormas.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1920
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   0
      Picture         =   "frmBandejaNormas.frx":0080
      Top             =   0
      Width           =   5250
   End
End
Attribute VB_Name = "frmBandejaNormas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim result As Long
    result = SetWindowLong(RichTextBox1.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
Me.RichTextBox1.Text = "Para Leer un mensaje solo hay que hacer clic en el numero y darle al botón 'LEER' ." & vbNewLine & _
"Para Borrar un mensaje solo hay que hacer clic en el numero y darle al botón 'BORRAR'." & vbNewLine & _
"Los mensajes descargados se guardan en la carpeta de AODrag '\Mensajes'." & vbNewLine & _
"                       REGLAMENTO" & vbNewLine & _
"1. El uso inadecuado del sistema de mensajeria tendra las siguientes consecuencias:" & vbNewLine & _
"1.1. Restricción del Sistema de Mensajeria temporalmente." & vbNewLine & _
"1.2. Expulsión temporal del juego." & vbNewLine & _
"1.3. Expulsión permanente del juego." & vbNewLine & _
"2. Se considera inadecuado: " & vbNewLine & _
"2.1. SPAM (se permiten enlaces a juegosdrag e Youtube dependiendo del tema.)" & vbNewLine & _
"2.2. Insultos" & vbNewLine & _
"2.3. Floodeo de Mensajes" & vbNewLine & _
"2.4. Denunciar a Usuarios sin ser un mensaje/asunto inadecuado." & vbNewLine & _
"3. El Staff podra:" & vbNewLine & _
"3.1. Ver los mensajes del usuario si es NECESARIO." & vbNewLine & _
"                                                 .- Staff AODrag"
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

