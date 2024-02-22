VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6135
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
   ScaleHeight     =   7185
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000040&
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
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000040&
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
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000040&
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
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000040&
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
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000040&
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
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000040&
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
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000040&
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
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   5655
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000040&
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
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   1
      Top             =   5880
      Width           =   5655
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00000040&
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
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   3480
      MouseIcon       =   "frmGuildDetails.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmGuildDetails.frx":0CCA
      Top             =   6480
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   1
      Left            =   1080
      MouseIcon       =   "frmGuildDetails.frx":519A
      MousePointer    =   99  'Custom
      Picture         =   "frmGuildDetails.frx":5E64
      Top             =   6480
      Width           =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Codexs"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   10
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGuildDetails.frx":A1C4
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   5895
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
frmGuildDetails.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
End Sub
Private Sub Command1_Click(index As Integer)




End Sub

Private Sub Form_Deactivate()

If Not frmGuildLeader.Visible Then
    Me.SetFocus
Else
    'Unload Me
End If

End Sub

Private Sub Image1_Click(index As Integer)
Select Case index

Case 0
    Unload Me
Case 1
    Dim fdesc$
    fdesc$ = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)
    
'    If Not AsciiValidos(fdesc$) Then
'        MsgBox "La descripcion contiene caracteres invalidos"
'        Exit Sub
'    End If
    
    Dim k As Integer
    Dim Cont As Integer
    Cont = 0
    For k = 0 To txtCodex1.UBound
'        If Not AsciiValidos(txtCodex1(k)) Then
'            MsgBox "El codex tiene invalidos"
'            Exit Sub
'        End If
        If Len(txtCodex1(k).Text) > 0 Then Cont = Cont + 1
    Next k
    If Cont < 4 Then
            MsgBox "Debes definir al menos cuatro mandamientos."
            Exit Sub
    End If
    
    Dim chunk$
    
    If CreandoClan Then
        chunk$ = "CIG" & fdesc$
        chunk$ = chunk$ & "¬" & ClanName & "¬" & Site & "¬" & Cont
    Else
        chunk$ = "DESCOD" & fdesc$ & "¬" & Cont
    End If
    
    
    
    For k = 0 To txtCodex1.UBound
        chunk$ = chunk$ & "¬" & txtCodex1(k)
    Next k
    
    
    Call SendData(chunk$)
    
    CreandoClan = False
    
    Unload Me
    
End Select
End Sub
