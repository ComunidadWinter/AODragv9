VERSION 5.00
Begin VB.Form frmBorrar 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3810
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3705
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
   ScaleHeight     =   3810
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdEliminarPersonaje 
      Caption         =   "Eliminar Personaje"
      Height          =   360
      Left            =   1920
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtBorrar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblSeguro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¿Seguro que quieres borrar a XX?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   3120
   End
   Begin VB.Label lblCUIDADO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡CUIDADO!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label lblInform 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡Vas a borrar un personaje! ¿Estas seguro que quieres borrar este personaje? Para borrar el personaje escribe BORRAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmBorrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEliminarPersonaje_Click()
    If MsgBox("¿Seguro que deseas eliminar este personaje?", vbYesNo, "Atencion!") = vbNo Then Exit Sub
    
    Call WriteBorrarPersonaje(frmCuenta.ListPJ.List(frmCuenta.ListPJ.ListIndex))
End Sub

Private Sub cmdVolver_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblSeguro.Caption = "¿Seguro que quieres borrar a " & frmCuenta.ListPJ.List(frmCuenta.ListPJ.ListIndex) & "?"
End Sub
