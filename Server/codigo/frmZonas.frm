VERSION 5.00
Begin VB.Form frmZonas 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7530
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
   ScaleHeight     =   6630
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   5160
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6960
      Top             =   6120
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jugadores en la zona"
      Height          =   195
      Left            =   5160
      TabIndex        =   3
      Top             =   840
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zonas en memoria"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1305
   End
End
Attribute VB_Name = "frmZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()


  frmZonas.List1.Clear
  
  Dim ZonaIndex As Long
  For ZonaIndex = 0 To UBound(ZonaList) - 1
    frmZonas.List1.AddItem ZonaList(ZonaIndex).Mapa & ZonaList(ZonaIndex).nombre & " " & " " & ZonaList(ZonaIndex).x1 & " " & ZonaList(ZonaIndex).y1
  Next
End Sub
