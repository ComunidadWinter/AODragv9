VERSION 5.00
Begin VB.Form frmResucitar 
   BorderStyle     =   0  'None
   Caption         =   "   "
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cementerio"
      Height          =   375
      Left            =   1440
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Haz click para transportado al cementerio más cercano."
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Estás muerto."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmResucitar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Call WriteIrAlCementerio
    frmResucitar.Hide
End Sub
