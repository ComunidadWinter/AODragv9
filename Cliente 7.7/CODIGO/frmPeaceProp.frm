VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   0  'None
   Caption         =   "Ofertas de paz"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6150
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
   ScaleHeight     =   7215
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
      BackColor       =   &H00000040&
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
      Height          =   2790
      ItemData        =   "frmPeaceProp.frx":0000
      Left            =   600
      List            =   "frmPeaceProp.frx":0002
      TabIndex        =   0
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   4080
      MouseIcon       =   "frmPeaceProp.frx":0004
      MousePointer    =   99  'Custom
      Picture         =   "frmPeaceProp.frx":0CCE
      Top             =   5280
      Width           =   1650
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   4200
      MouseIcon       =   "frmPeaceProp.frx":5052
      Picture         =   "frmPeaceProp.frx":5D1C
      Top             =   6600
      Width           =   1650
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2280
      MouseIcon       =   "frmPeaceProp.frx":9F3C
      MousePointer    =   99  'Custom
      Picture         =   "frmPeaceProp.frx":AC06
      Top             =   5280
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   480
      MouseIcon       =   "frmPeaceProp.frx":EDDB
      MousePointer    =   99  'Custom
      Picture         =   "frmPeaceProp.frx":FAA5
      Top             =   5280
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ofertas de Paz"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
End
Attribute VB_Name = "frmPeaceProp"
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

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
frmPeaceProp.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
End Sub


Private Sub Command1_Click()

End Sub

Public Sub ParsePeaceOffers(ByVal s As String)

Dim t%, r%

t% = Val(ReadField(1, s, 44))

For r% = 1 To t%
    Call lista.AddItem(ReadField(r% + 1, s, 44))
Next r%

Me.Show vbModeless, frmMain

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Image1_Click()
'Me.Visible = False
Call SendData("ACEPPEAT" & lista.List(lista.ListIndex))
Unload Me
End Sub

Private Sub Image2_Click()
'Me.Visible = False
Call SendData("PEACEDET" & lista.List(lista.ListIndex))
End Sub

Private Sub Image3_Click()
Unload Me
End Sub
