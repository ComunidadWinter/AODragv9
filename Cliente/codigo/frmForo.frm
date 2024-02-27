VERSION 5.00
Begin VB.Form frmForo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   5490
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   5055
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   5505
      Index           =   0
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmForo.frx":0000
      Top             =   1080
      Visible         =   0   'False
      Width           =   5025
   End
   Begin VB.TextBox MiMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   345
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox MiMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   4575
      Index           =   1
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image cmdForo 
      Height          =   300
      Left            =   600
      Top             =   6600
      Width           =   1650
   End
   Begin VB.Image cmdEnviar 
      Height          =   300
      Left            =   2280
      Top             =   6600
      Width           =   1650
   End
   Begin VB.Image cmdVolver 
      Height          =   300
      Left            =   3960
      Top             =   6600
      Width           =   1680
   End
End
Attribute VB_Name = "frmForo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ForoIndex As Integer
Private Sub cmdEnviar_Click()
Dim i
For Each i In Text
    i.Visible = False
Next

If Not MiMensaje(0).Visible Then
    List.Visible = False
    MiMensaje(0).Visible = True
    MiMensaje(1).Visible = True
    MiMensaje(0).SetFocus
    cmdEnviar.Enabled = False
    Label1.Visible = True
    Label2.Visible = True
Else
    Call WriteForumPost(MiMensaje(0).Text, Left$(MiMensaje(1).Text, 450))
    List.AddItem MiMensaje(0).Text
    Load Text(List.ListCount)
    Text(List.ListCount - 1).Text = MiMensaje(1).Text
    List.Visible = True
    
    MiMensaje(0).Visible = False
    MiMensaje(1).Visible = False
    'Limpio los textboxs (NicoNZ) 04/24/08
    MiMensaje(0).Text = vbNullString
    MiMensaje(1).Text = vbNullString
    
    cmdEnviar.Enabled = True
    Label1.Visible = False
    Label2.Visible = False
End If
End Sub

Private Sub cmdVolver_Click()
Unload Me
End Sub

Private Sub cmdForo_Click()

    MiMensaje(0).Visible = False
    MiMensaje(1).Visible = False
    cmdEnviar.Enabled = True
    Label1.Visible = False
    Label2.Visible = False
    Dim i
    For Each i In Text
        i.Visible = False
    Next
    List.Visible = True
End Sub

Private Sub Form_Load()
    Me.Picture = General_Load_Picture_From_Resource("32.gif")
    cmdForo.Picture = General_Load_Picture_From_Resource("33.gif")
    cmdEnviar.Picture = General_Load_Picture_From_Resource("34.gif")
    cmdvolver.Picture = General_Load_Picture_From_Resource("35.gif")
    
End Sub


Private Sub List_Click()
    If (List.ListIndex + 1) > List.ListCount Then Exit Sub
    List.Visible = False
    Text(List.ListIndex).Visible = True

End Sub

Private Sub MiMensaje_Change(index As Integer)
    If Len(MiMensaje(0).Text) <> 0 And Len(MiMensaje(1).Text) <> 0 Then
        cmdEnviar.Enabled = True
    End If

End Sub

