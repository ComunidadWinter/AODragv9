VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   0  'None
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   15
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   5580
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2160
      Picture         =   "frmMSG.frx":0000
      Top             =   6480
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensajes de los Usuarios"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   405
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   3525
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_GM_MSG = 300

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG) As Integer

Public Sub CrearGMmSg(nick As String, msg As String)
If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1
End If
End Sub



Private Sub Form_Deactivate()
Me.Visible = False
List1.Clear
End Sub

Private Sub Form_Load()
frmMSG.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
List1.Clear

End Sub

Private Sub Image1_Click()
Me.Visible = False
List1.Clear
End Sub

Private Sub List1_Click()
Dim ind As String
ind = ReadField(2, List1.List(List1.ListIndex), Asc(";"))
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu menU_usuario
End If

End Sub

Private Sub mnuBorrar_Click()
If List1.ListIndex < 0 Then Exit Sub
SendData ("SOSDONE" & List1.List(List1.ListIndex))

List1.RemoveItem List1.ListIndex

End Sub

Private Sub mnuIR_Click()
SendData ("/IRA " & ReadField(1, List1.List(List1.ListIndex), Asc(";")))
End Sub

Private Sub mnutraer_Click()
SendData ("/SUM " & ReadField(1, List1.List(List1.ListIndex), Asc(";")))
End Sub
