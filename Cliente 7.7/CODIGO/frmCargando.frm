VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7530
   ClientLeft      =   1005
   ClientTop       =   810
   ClientWidth     =   10035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCargando.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   669
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Status 
      Height          =   2220
      Left            =   3480
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   2760
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   3916
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmCargando.frx":0CCA
      MouseIcon       =   "frmCargando.frx":0D4E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image LOGO 
      Height          =   7500
      Left            =   0
      Picture         =   "frmCargando.frx":0D6A
      Top             =   0
      Width           =   9990
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim result As Long
    result = SetWindowLong(status.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    LOGO.Picture = LoadPicture(DirGraficos & "mar.jpg")
End Sub


