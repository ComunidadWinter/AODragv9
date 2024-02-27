VERSION 5.00
Begin VB.Form frmMensaje 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   -60
   ClientTop       =   -465
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Label msg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuMensaje 
      Caption         =   "Mensaje"
      Visible         =   0   'False
      Begin VB.Menu mnuNormal 
         Caption         =   "Normal"
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Global"
      End
      Begin VB.Menu mnuPrivado 
         Caption         =   "Privado"
      End
      Begin VB.Menu mnuGritar 
         Caption         =   "Gritar"
      End
      Begin VB.Menu mnuGrupo 
         Caption         =   "Grupo"
      End
      Begin VB.Menu mnuClan 
         Caption         =   "Clan"
      End
   End
End
Attribute VB_Name = "frmMensaje"
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

Private Sub Command1_Click()
msg.Caption = ""
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Public Sub PopupMenuMensaje()

Select Case SendingType
    Case 1
        mnuNormal.Checked = True
        mnuGritar.Checked = False
        mnuPrivado.Checked = False
        mnuClan.Checked = False
        mnuGrupo.Checked = False
        mnuGlobal.Checked = False
        frmMain.LbLChat.Caption = "1. Normal"
    Case 2
        mnuNormal.Checked = False
        mnuGritar.Checked = True
        mnuPrivado.Checked = False
        mnuClan.Checked = False
        mnuGrupo.Checked = False
        mnuGlobal.Checked = False
        frmMain.LbLChat.Caption = "2. Gritar"
    Case 3
        mnuNormal.Checked = False
        mnuGritar.Checked = False
        mnuPrivado.Checked = True
        mnuClan.Checked = False
        mnuGrupo.Checked = False
        mnuGlobal.Checked = False
        frmMain.LbLChat.Caption = "3. Privado"
    Case 4
        mnuNormal.Checked = False
        mnuGritar.Checked = False
        mnuPrivado.Checked = False
        mnuClan.Checked = True
        mnuGrupo.Checked = False
        mnuGlobal.Checked = False
        frmMain.LbLChat.Caption = "4. Clan"
    Case 5
        mnuNormal.Checked = False
        mnuGritar.Checked = False
        mnuPrivado.Checked = False
        mnuClan.Checked = False
        mnuGrupo.Checked = True
        mnuGlobal.Checked = False
        frmMain.LbLChat.Caption = "5. Party"
    Case 6
        mnuNormal.Checked = False
        mnuGritar.Checked = False
        mnuPrivado.Checked = False
        mnuClan.Checked = False
        mnuGrupo.Checked = False
        mnuGlobal.Checked = True
        frmMain.LbLChat.Caption = "6. Global"
End Select

PopupMenu mnuMensaje
frmMain.LbLChat.Refresh
End Sub

'[Lorwik]
'Moví este menú acá para que se pueda ver el caption del
'frmMain sin que se tenga que ver el ControlBox

Private Sub mnuNormal_Click()

SendingType = 1
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGritar_click()

SendingType = 2
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuPrivado_click()

sndPrivateTo = InputBox("Nombre del destinatario:", vbNullString)

If sndPrivateTo <> vbNullString Then
    SendingType = 3
    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
Else
    MsgBox "¡Escribe un nombre."
End If

End Sub

Private Sub mnuClan_click()

SendingType = 4
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGrupo_click()

SendingType = 5
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGlobal_Click()

SendingType = 6
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

