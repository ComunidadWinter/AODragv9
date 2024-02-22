VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBandejaRedactar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Bandeja de Redactar"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox3 
      Height          =   1935
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3413
      _Version        =   393217
      TextRTF         =   $"frmBandejaRedactar.frx":0000
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   390
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393217
      BorderStyle     =   0
      MultiLine       =   0   'False
      TextRTF         =   $"frmBandejaRedactar.frx":0082
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   810
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393217
      BorderStyle     =   0
      MultiLine       =   0   'False
      TextRTF         =   $"frmBandejaRedactar.frx":0104
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   4680
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   3360
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   0
      Picture         =   "frmBandejaRedactar.frx":0186
      Top             =   0
      Width           =   6900
   End
End
Attribute VB_Name = "frmBandejaRedactar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim result As Long
    result = SetWindowLong(RichTextBox1.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    result = SetWindowLong(RichTextBox2.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    result = SetWindowLong(RichTextBox3.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
RichTextBox1.SelColor = vbWhite
RichTextBox2.SelColor = vbWhite
End Sub

Private Sub Image2_Click()
seña = "#"
Dim MiString As String
MiString = RichTextBox3.Text
MiString = Replace(MiString, vbCrLf, " ")
     SendData "/smsuser " & RichTextBox1.Text & seña & RichTextBox2.Text & seña & MiString
End Sub

Private Sub Label1_Click()
MsgBox "¿Deseas salir?", vbOKCancel & vbOKOnly
Unload Me
End Sub

Private Sub Label4_Click()
seña = "#"
Dim MiString As String
MiString = RichTextBox3.Text
MiString = Replace(MiString, vbCrLf, " ")
     SendData "/smsuser " & RichTextBox1.Text & seña & RichTextBox2.Text & seña & MiString
End Sub
