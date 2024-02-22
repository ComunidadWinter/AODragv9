VERSION 5.00
Begin VB.Form frmEquipo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   3720
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   4800
      Width           =   555
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   4080
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   480
      Width           =   555
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   5160
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   2760
      Width           =   555
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   3840
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1920
      Width           =   555
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2400
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   2280
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   4440
      Picture         =   "frmEquipo.frx":0000
      Top             =   5880
      Width           =   1680
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   2760
      TabIndex        =   9
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   4200
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.Image FONDO 
      Height          =   6525
      Left            =   0
      Picture         =   "frmEquipo.frx":42B9
      Top             =   0
      Width           =   6960
   End
End
Attribute VB_Name = "frmEquipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
'Cargamos la interfase
Me.Picture = LoadPicture(App.Path & "\Graficos\guerre.jpg")
Dim n As Integer
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

For n = 1 To MAX_INVENTORY_SLOTS
If UserInventory(n).Equipped > 0 Then
If UserInventory(n).OBJType = 2 Then
Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hdc, UserInventory(n).GrhIndex, SR, DR)
Label1.Caption = "Daño: " & UserInventory(n).MinHIT & "-" & UserInventory(n).MaxHIT & vbCrLf & "Peso: " & UserInventory(n).peso
End If
If UserInventory(n).OBJType = 3 And UserInventory(n).SubTipo = 0 Then
Call DrawGrhtoHdc(Picture2.hWnd, Picture2.hdc, UserInventory(n).GrhIndex, SR, DR)
Label2.Caption = "Defensa: " & UserInventory(n).DefMax & vbCrLf & "Peso: " & UserInventory(n).peso
End If
If UserInventory(n).OBJType = 3 And UserInventory(n).SubTipo = 2 Then
Call DrawGrhtoHdc(Picture3.hWnd, Picture3.hdc, UserInventory(n).GrhIndex, SR, DR)
Label3.Caption = "Defensa: " & UserInventory(n).DefMax & vbCrLf & "Peso: " & UserInventory(n).peso
End If
If UserInventory(n).OBJType = 3 And UserInventory(n).SubTipo = 1 Then
Call DrawGrhtoHdc(Picture4.hWnd, Picture4.hdc, UserInventory(n).GrhIndex, SR, DR)
Label4.Caption = "Defensa: " & UserInventory(n).DefMax & vbCrLf & "Peso: " & UserInventory(n).peso
End If

If UserInventory(n).OBJType = 3 And UserInventory(n).SubTipo = 3 Then
Call DrawGrhtoHdc(Picture5.hWnd, Picture5.hdc, UserInventory(n).GrhIndex, SR, DR)
Label5.Caption = "Defensa: " & UserInventory(n).DefMax & vbCrLf & "Peso: " & UserInventory(n).peso
End If


End If
Next n
End Sub

Private Sub Image1_Click()
Unload Me
End Sub
