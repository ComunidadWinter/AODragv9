VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensajes de GMs"
   ClientHeight    =   5715
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8715
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar Todo"
      Height          =   480
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   480
      Left            =   2880
      TabIndex        =   6
      Top             =   5040
      Width           =   5535
   End
   Begin VB.TextBox txtMensaje 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   5655
   End
   Begin VB.TextBox TxTCategoria 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   5655
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2460
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Categoria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   825
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
Option Explicit

Private Const MAX_GM_MSG = 300

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG) As Integer

Public Sub CrearGMmSg(Nick As String, msg As String)
If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1
End If
End Sub

Private Sub cmdBorrar_Click()
    Dim i As Byte
    
    If MsgBox("¿Seguro que desea eliminar todas las consultas?", vbYesNo, "Atencion!") = vbNo Then Exit Sub
    
    For i = 0 To List1.ListCount
        Dim aux As String
        aux = mid$(ReadField(1, List1.List(i), Asc("-")), 10, Len(ReadField(1, List1.List(i), Asc("-"))))
        Call WriteSOSRemove(aux)
     Next i
     
     List1.Clear
End Sub

Private Sub Command1_Click()
Me.Visible = False
List1.Clear
txtMensaje.Text = ""
TxTCategoria.Text = ""
End Sub

Private Sub Form_Deactivate()
Me.Visible = False
List1.Clear
End Sub

Private Sub Form_Load()
List1.Clear

End Sub

Private Sub List1_Click()
Dim ind As Integer
ind = Val(ReadField(2, List1.List(List1.ListIndex), Asc("-")))
TxTCategoria.Text = ReadField(3, List1.List(List1.ListIndex), Asc("|"))
txtMensaje.Text = ReadField(4, List1.List(List1.ListIndex), Asc("|"))
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    PopupMenu menU_usuario
End If

End Sub


Private Sub mnuBorrar_Click()
    If List1.ListIndex < 0 Then Exit Sub
    'Pablo (ToxicWaste)
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.ListIndex), Asc("-"))))
    Call WriteSOSRemove(aux)
    '/Pablo (ToxicWaste)
    'Call WriteSOSRemove(List1.List(List1.listIndex))
    
    List1.RemoveItem List1.ListIndex
End Sub

Private Sub mnuIR_Click()
    'Pablo (ToxicWaste)
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.ListIndex), Asc("-"))))
    aux = ReadField(1, aux, Asc("|"))
    Call WriteGoToChar(aux)
    '/Pablo (ToxicWaste)
    'Call WriteGoToChar(ReadField(1, List1.List(List1.listIndex), Asc("-")))
    
End Sub

Private Sub mnutraer_Click()
    'Pablo (ToxicWaste)
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.ListIndex), Asc("-"))))
    Call WriteSummonChar(aux)
    'Pablo (ToxicWaste)
    'Call WriteSummonChar(ReadField(1, List1.List(List1.listIndex), Asc("-")))
End Sub
