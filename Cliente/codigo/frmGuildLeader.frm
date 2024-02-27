VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración del Clan"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
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
   ScaleHeight     =   7005
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   3840
      MouseIcon       =   "frmGuildLeader.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lista de clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3375
      Begin VB.ListBox guildslist 
         Height          =   2595
         ItemData        =   "frmGuildLeader.frx":0152
         Left            =   120
         List            =   "frmGuildLeader.frx":0154
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Información"
         Height          =   375
         Left            =   360
         MouseIcon       =   "frmGuildLeader.frx":0156
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   3000
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "Mensaje del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   10095
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "frmGuildLeader.frx":02A8
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   9735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command2 
         Caption         =   "Expulsar"
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   3120
         Width           =   1575
      End
      Begin VB.ListBox members 
         Height          =   2595
         ItemData        =   "frmGuildLeader.frx":03FA
         Left            =   120
         List            =   "frmGuildLeader.frx":03FC
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command6 
         Caption         =   "Rechazar"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   3120
         Width           =   855
      End
      Begin VB.ListBox solicitudes 
         Height          =   2595
         ItemData        =   "frmGuildLeader.frx":03FE
         Left            =   120
         List            =   "frmGuildLeader.frx":0400
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Const MAX_NEWS_LENGTH As Integer = 512

Private Sub cmdElecciones_Click()
    Call WriteGuildOpenElections
    Unload Me
End Sub

Private Sub Command1_Click()
    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))

    'Unload Me
End Sub

Private Sub Command2_Click()
    If members.ListIndex = -1 Then Exit Sub
    
    Call WriteGuildKickMember(members.List(members.ListIndex))
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    

    'Unload Me
End Sub

Private Sub Command3_Click()
    Dim k As String

    k = Replace(txtguildnews, vbCrLf, "º")
    
    Call WriteGuildUpdateNews(k)
End Sub

Private Sub Command4_Click()
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

    'Unload Me
End Sub

Private Sub Command5_Click()
   Call WriteGuildAcceptNewMember(solicitudes.List(solicitudes.ListIndex))
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub Command6_Click()
Call WriteGuildRejectNewMember(solicitudes.List(solicitudes.ListIndex))
'Unload Me
End Sub

Private Sub Command7_Click()
    Call WriteGuildPeacePropList
End Sub
Private Sub Command9_Click()
    Call WriteGuildAlliancePropList
End Sub

Private Sub Command8_Click()
    Unload Me
    frmMain.SetFocus
End Sub



Private Sub txtguildnews_Change()
    If Len(txtguildnews.Text) > MAX_NEWS_LENGTH Then _
        txtguildnews.Text = Left$(txtguildnews.Text, MAX_NEWS_LENGTH)
End Sub
