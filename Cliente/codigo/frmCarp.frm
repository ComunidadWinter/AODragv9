VERSION 5.00
Begin VB.Form frmCarp 
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   -60
   ClientTop       =   -120
   ClientWidth     =   6765
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
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
      Height          =   2760
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox TxTCant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2850
      TabIndex        =   0
      Text            =   "1"
      Top             =   3750
      Width           =   1095
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Con tu energia actual puedes construir hasta 0 unidades"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   4320
      Top             =   3720
      Width           =   2295
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image1_Click()
    On Error Resume Next
    
    Dim cantidad As String
    
    cantidad = TxTCant.Text
    
    If Not IsNumeric(cantidad) Or Int(Val(cantidad)) < 1 Or Int(Val(cantidad)) > 1000 Then
        MsgBox "La cantidad es invalida.", vbCritical
        Exit Sub
    End If
    Call WriteCraftCarpenter(ObjCarpintero(lstArmas.ListIndex + 1), cantidad)
    
    Unload Me
End Sub

Private Sub Image4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblEnergia.Caption = "Con tu energia maxima actual puedes construir hasta " & Round(UserMaxSTA / 2) & " unidades."
    Me.Picture = General_Load_Picture_From_Resource("51.gif")
End Sub
