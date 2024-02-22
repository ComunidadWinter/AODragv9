VERSION 5.00
Begin VB.Form frmUserRequest 
   BorderStyle     =   0  'None
   Caption         =   "Peticion"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6120
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
   ScaleHeight     =   7185
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000040&
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2280
      MouseIcon       =   "frmUserRequest.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmUserRequest.frx":0CCA
      Top             =   5280
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Petición"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "frmUserRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
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

Private Sub Form_Load()
frmUserRequest.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
End Sub



Public Sub recievePeticion(ByVal p As String)

Text1 = Replace(p, "º", vbCrLf)
Me.Show vbModeless, frmMain

End Sub

Private Sub Image1_Click()
Unload Me
End Sub
