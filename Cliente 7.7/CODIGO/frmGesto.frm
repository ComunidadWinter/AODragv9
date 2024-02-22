VERSION 5.00
Begin VB.Form frmGesto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elige un Icono"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmGesto.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   1305
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   20
      Left            =   0
      Top             =   960
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   21
      Left            =   435
      Top             =   960
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   22
      Left            =   870
      Top             =   960
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   23
      Left            =   1305
      Top             =   960
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   24
      Left            =   1740
      Top             =   960
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   25
      Left            =   2175
      Top             =   960
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   26
      Left            =   2640
      Top             =   960
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   27
      Left            =   3120
      Top             =   960
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   28
      Left            =   3600
      Top             =   960
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   29
      Left            =   4080
      Top             =   960
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   10
      Left            =   0
      Top             =   480
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   11
      Left            =   435
      Top             =   480
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   12
      Left            =   870
      Top             =   480
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   13
      Left            =   1305
      Top             =   480
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   14
      Left            =   1740
      Top             =   480
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   15
      Left            =   2175
      Top             =   480
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   16
      Left            =   2640
      Top             =   480
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   17
      Left            =   3120
      Top             =   480
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   18
      Left            =   3600
      Top             =   480
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   19
      Left            =   4080
      Top             =   480
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   9
      Left            =   4080
      Top             =   0
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   8
      Left            =   3600
      Top             =   0
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   7
      Left            =   3120
      Top             =   0
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   6
      Left            =   2640
      Top             =   0
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   5
      Left            =   2170
      Top             =   0
      Width           =   480
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   4
      Left            =   1740
      Top             =   0
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   3
      Left            =   1300
      Top             =   0
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   2
      Left            =   870
      Top             =   0
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   1
      Left            =   440
      Top             =   0
      Width           =   420
   End
   Begin VB.Image gestosImage1 
      Height          =   495
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "frmGesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'pluto:hoy
Private Sub Form_Load()
frmGesto.Picture = LoadPicture(DirGraficos & "9720.bmp")
End Sub
Private Sub Commandsalir_Click()
frmgestos.Visible = False
End Sub

Private Sub gestosImage1_Click(index As Integer)
Call SendData("IC" & index)
Call audio.PlayWave(SND_CLICK)
Unload frmGesto
End Sub


