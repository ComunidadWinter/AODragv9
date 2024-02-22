VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBandejaLectora 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBandejaLectora.frx":0000
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   351
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2895
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5106
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   0   'False
      TextRTF         =   $"frmBandejaLectora.frx":59E8A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Asunto:"
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
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1035
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
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
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   510
      Width           =   4215
   End
End
Attribute VB_Name = "frmBandejaLectora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim result As Long
    result = SetWindowLong(RichTextBox1.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label8_Click()
If Label2.Caption = "AOdragbot" Then
Exit Sub
End If
seña = "#"
X = Label2.Caption
xx = Label1.Caption
Label2.Caption = Right(Label2, Len(Label2) - 4)
Label1.Caption = Right(Label1, Len(Label1) - 8)
nombresms = Label2.Caption
asuntosms = Label1.Caption

     SendData "/smspam " & nombresms & seña & asuntosms & seña & RichTextBox1
     Label2.Caption = X
     Label1.Caption = xx
End Sub
