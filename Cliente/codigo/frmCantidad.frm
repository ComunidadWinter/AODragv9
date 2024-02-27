VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2025
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   4185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   135
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   585
      TabIndex        =   0
      Text            =   "1"
      Top             =   600
      Width           =   3045
   End
   Begin VB.Image cmdTirarTodo 
      Height          =   255
      Index           =   0
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Image cmdCerrar 
      Height          =   255
      Left            =   1320
      MousePointer    =   99  'Custom
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Image cmdAceptar 
      Height          =   255
      Left            =   2160
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmCantidad"
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

Private Sub cmdAceptar_Click()
    If LenB(frmCantidad.Text1.Text) > 0 Then
        If Not IsNumeric(frmCantidad.Text1.Text) Then Exit Sub  'Should never happen
        Call WriteDrop(Inventario.SelectedItem, frmCantidad.Text1.Text)
        frmCantidad.Text1.Text = ""
        Unload Me
    End If
    
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdTirarTodo_Click(Index As Integer)
    Call WriteDrop(Inventario.SelectedItem, 999)
    
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Picture = General_Load_Picture_From_Resource("50.gif")
End Sub

Private Sub text1_Change()
On Error GoTo errhandler
    If Val(Text1.Text) < 0 Then
        Text1.Text = "1"
    End If
    
    If Val(Text1.Text) > MAX_INVENTORY_OBJS Then
        Text1.Text = "10000"
    End If
    
    Exit Sub
    
errhandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Text1.Text = "1"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
