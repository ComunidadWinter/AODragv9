VERSION 5.00
Begin VB.Form frmBandejaEntrada 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Bandeja de Entrada"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   5070
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   570
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   5070
      Left            =   5955
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   5070
      Left            =   4635
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   5070
      Left            =   795
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2880
      Top             =   6720
      Width           =   1680
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bandeja de entrada de: Parmentier - 0 Mensajes nuevos"
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
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   0
      Top             =   0
      Width           =   7350
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2880
      Picture         =   "frmBandejaEntrada.frx":0000
      Top             =   6720
      Width           =   1680
   End
   Begin VB.Image Image3 
      Height          =   7200
      Left            =   0
      Picture         =   "frmBandejaEntrada.frx":1A82
      Top             =   0
      Width           =   7350
   End
End
Attribute VB_Name = "frmBandejaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dia = Format(Now, "dd/mm/yy")
Label5.Caption = Dia
SendData "/SMSREFRESH"
frmBandejaEntrada.Label2.Caption = "Bandeja de entrada de: " & frmMain.Label8.Caption & " - " & 0 & " Mensajes nuevos"
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Label1_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
SendData "/SMSREFRESH"
'MsgBox "Tienes la bandeja llena, elimina mensajes para poder recibir mas mensajes"
End Sub

Private Sub Label3_Click()
frmBandejaRedactar.Show
End Sub

Private Sub Label4_Click()
If List1.ListIndex = -1 Then
Exit Sub
Else
If List1.ListIndex = 0 Then
Indice1 = List1.ListIndex + 1
SendData "/SMSKILL " & Indice1
Indice = List1.ListIndex
List1.RemoveItem Indice
List2.RemoveItem Indice
List3.RemoveItem Indice
List4.RemoveItem Indice
Else
Indice1 = List1.ListIndex + 1
SendData "/SMSKILL " & Indice1
Indice = List1.ListIndex
List1.RemoveItem Indice
List2.RemoveItem Indice
List3.RemoveItem Indice
List4.RemoveItem Indice
End If
End If
End Sub

Private Sub Label7_Click()
If List1.ListIndex = 0 Then
Indice1 = List1.ListIndex + 1
SendData "/SMSREAD " & Indice1
frmBandejaLectora.Show
Else
Indice1 = List1.ListIndex + 1
SendData "/SMSREAD " & Indice1
frmBandejaLectora.Show
End If
End Sub

Private Sub Label8_Click()
frmBandejaNormas.Show
End Sub
