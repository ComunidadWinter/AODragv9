VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Opciones AodraG"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   9015
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4490
      TabIndex        =   3
      Text            =   "FPS 17"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4490
      TabIndex        =   2
      Text            =   "FPS 68"
      Top             =   2280
      Width           =   615
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Index           =   0
      Left            =   6795
      TabIndex        =   0
      Top             =   2235
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Index           =   1
      Left            =   6795
      TabIndex        =   1
      Top             =   3100
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   10
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.Image Image70 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      Left            =   5160
      MouseIcon       =   "frmOpciones.frx":15FA94
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   240
   End
   Begin VB.Image Image70 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   5160
      MouseIcon       =   "frmOpciones.frx":16075E
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   240
   End
   Begin VB.Image Image20 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   1
      Left            =   5040
      MouseIcon       =   "frmOpciones.frx":161428
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   255
   End
   Begin VB.Image Image20 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   0
      Left            =   5040
      MouseIcon       =   "frmOpciones.frx":1620F2
      MousePointer    =   99  'Custom
      Top             =   6650
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   4320
      MouseIcon       =   "frmOpciones.frx":162DBC
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   3375
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":163A86
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   3375
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   7800
      MouseIcon       =   "frmOpciones.frx":164750
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   3495
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   2
      Left            =   9360
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   1
      Left            =   8160
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   255
      Index           =   0
      Left            =   6960
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   3
      Left            =   9720
      MouseIcon       =   "frmOpciones.frx":16541A
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   255
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   2
      Left            =   9720
      MouseIcon       =   "frmOpciones.frx":1660E4
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   255
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   1
      Left            =   5280
      MouseIcon       =   "frmOpciones.frx":166DAE
      MousePointer    =   99  'Custom
      Top             =   4350
      Width           =   255
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   0
      Left            =   5280
      MouseIcon       =   "frmOpciones.frx":167A78
      MousePointer    =   99  'Custom
      Top             =   3950
      Width           =   255
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   1
      Left            =   5400
      MouseIcon       =   "frmOpciones.frx":168742
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   0
      Left            =   5400
      MouseIcon       =   "frmOpciones.frx":16940C
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   255
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   2
      Left            =   4200
      MouseIcon       =   "frmOpciones.frx":16A0D6
      MousePointer    =   99  'Custom
      Top             =   3020
      Width           =   240
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   210
      Index           =   1
      Left            =   4200
      MouseIcon       =   "frmOpciones.frx":16ADA0
      MousePointer    =   99  'Custom
      Top             =   2670
      Width           =   210
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   4200
      MouseIcon       =   "frmOpciones.frx":16BA6A
      MousePointer    =   99  'Custom
      Top             =   2300
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   7440
      MouseIcon       =   "frmOpciones.frx":16C734
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   1815
   End
End
Attribute VB_Name = "frmOpciones"
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
Option Explicit
Private loading As Boolean


Private Sub Command1_Click()

If SinTecho = 0 And navida = 0 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "normal", 1)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sintechos", 0)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "navidad", 0)
ElseIf SinTecho = 1 And navida = 0 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "normal", 0)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sintechos", 1)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "navidad", 0)
ElseIf SinTecho = 0 And navida = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "normal", 0)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sintechos", 0)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "navidad", 1)
End If

If VelFPS = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "VelFPS", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "VelFPS", 0)
End If
If ConFlash = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "flash", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "flash", 0)
End If
If Resolu = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "resolucion", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "resolucion", 0)
End If
'pluto:6.3
If LugarServer = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Server", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Server", 2)
End If


Unload frmOpciones
 frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 0, 0, 1, 0, 1
        
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Argentum Online!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms

        
End Sub

Private Sub Command2_Click()

End Sub



Private Sub Command3_Click()

End Sub

Private Sub Form_Deactivate()
Me.Visible = False
End Sub

Private Sub Form_Load()
frmOpciones.Picture = LoadPicture(DirGraficos & "Opciones.jpg")
  ConFlash = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "flash"))
  navida = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "navidad"))
  SinTecho = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sintechos"))
  Resolu = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "resolucion"))
LugarServer = Val(GetSetting("AODRAG", "SERVIDOR", "ACTUAL", 1)) 'Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Server"))
Chats = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Chat"))
VelFPS = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "VelFPS"))
If navida = 1 Then
    Image1(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image1(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image1(2).Picture = LoadPicture(DirGraficos & "validar.jpg")
ElseIf SinTecho = 1 Then
     Image1(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image1(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image1(2).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
Else
Image1(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
Image1(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
Image1(2).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
End If

If ConFlash = 1 Then
   
    Image2(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image2(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Else
  Image2(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image2(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
 
End If

If Chats = 1 Then
   
    Image20(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image20(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Else
  Image20(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image20(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
 
End If
If VelFPS = 1 Then
   
    Image70(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image70(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Else
    Image70(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image70(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
 
End If


If Resolu = 1 Then
    Image3(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image3(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Else
    Image3(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image3(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
End If
'pluto:6.3
If LugarServer = 1 Then
Image3(3).Picture = LoadPicture(DirGraficos & "validar.jpg")
Image3(2).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
Else
Image3(2).Picture = LoadPicture(DirGraficos & "validar.jpg")
Image3(3).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
End If




loading = True      'Prevent sounds when setting check's values
    
    If audio.MusicActivated Then
        Image4(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
        Slider1(0).Enabled = True
        Slider1(0).Value = audio.MusicVolume
    Else
        Image4(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")

        Slider1(0).Enabled = False
    End If
    
    If audio.SoundActivated Then
        Image4(1).Picture = LoadPicture(DirGraficos & "validar.jpg")

        Slider1(1).Enabled = True
        Slider1(1).Value = audio.SoundVolume
    'pluto:6.0A
    
If Fasis > 0 Then
Image4(2).Picture = LoadPicture(DirGraficos & "validar.jpg")
Else
Image4(2).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
End If
    
    Else
        Image4(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")

        Slider1(1).Enabled = False
    End If

    
 loading = False





End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

End Function

Private Sub Image1_Click(Index As Integer)
Call audio.PlayWave(SND_CLICK)
Select Case Index
Case 0
navida = 0
SinTecho = 0
Case 1
SinTecho = 1
navida = 0
Case 2
SinTecho = 0
navida = 1
End Select

If navida = 1 Then
    Image1(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image1(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image1(2).Picture = LoadPicture(DirGraficos & "validar.jpg")
ElseIf SinTecho = 1 Then
     Image1(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image1(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image1(2).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
Else
Image1(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
Image1(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
Image1(2).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
End If

End Sub

Private Sub Image2_Click(Index As Integer)
Call audio.PlayWave(SND_CLICK)
Select Case Index
Case 0
ConFlash = 1
Case 1
ConFlash = 0
End Select

If ConFlash = 1 Then
   
    Image2(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image2(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Else
  Image2(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image2(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
 
End If
End Sub

Private Sub Image20_Click(Index As Integer)
Call audio.PlayWave(SND_CLICK)
Select Case Index
Case 0
Chats = 1
Case 1
Chats = 0
End Select

If Chats = 1 Then
   
    Image20(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image20(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Else
  Image20(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image20(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
 
End If
End Sub

Private Sub Image3_Click(Index As Integer)
Call audio.PlayWave(SND_CLICK)
Select Case Index
Case 0
Resolu = 1
Case 1
Resolu = 0
Case 2
LugarServer = 2
Case 3
LugarServer = 1

End Select
If Resolu = 1 Then
    Image3(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image3(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Else
    Image3(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image3(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
End If



If LugarServer = 2 Then
Image3(2).Picture = LoadPicture(DirGraficos & "validar.jpg")
Image3(3).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
Else
Image3(3).Picture = LoadPicture(DirGraficos & "validar.jpg")
Image3(2).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
End If
End Sub

Private Sub Image4_Click(Index As Integer)
    Call audio.PlayWave(SND_CLICK)
    Select Case Index
    Case 0
    
If Musi = 1 Then
    Musi = 0
    Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "musica", 0)
    Image4(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    audio.MusicActivated = False
    Slider1(0).Enabled = False
ElseIf Not audio.MusicActivated Then  'Prevent the music from reloading
    Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "musica", 1)
    Image4(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
    audio.MusicActivated = True
    Slider1(0).Enabled = True
    Slider1(0).Value = audio.MusicVolume
    Musi = 1
End If

If Chats = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Chat", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Chat", 0)
End If

    
    Case 1
    
    If Son = 0 Then
Son = 1
audio.SoundActivated = True
Slider1(1).Enabled = True
Slider1(1).Value = audio.SoundVolume
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sonido", 1)
Image4(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
  
    Else
Son = 0
audio.SoundActivated = False
frmMain.IsPlaying = PlayLoop.plNone
Slider1(1).Enabled = False
Image4(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sonido", 0)
    
    End If
    
    Case 2
   If Fasis = 0 Then
Fasis = 1
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "asistente", 1)
Image4(2).Picture = LoadPicture(DirGraficos & "validar.jpg")
 
    Else
Fasis = 0
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "asistente", 0)
  Image4(2).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
  
    End If
    
    End Select
    
    
End Sub

Private Sub Image5_Click()

If SinTecho = 0 And navida = 0 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "normal", 1)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sintechos", 0)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "navidad", 0)
ElseIf SinTecho = 1 And navida = 0 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "normal", 0)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sintechos", 1)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "navidad", 0)
ElseIf SinTecho = 0 And navida = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "normal", 0)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sintechos", 0)
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "navidad", 1)
End If

If ConFlash = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "flash", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "flash", 0)
End If
If Chats = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Chat", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Chat", 0)
End If
If Resolu = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "resolucion", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "resolucion", 0)
End If
'pluto:6.3
If LugarServer = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Server", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Server", 2)
End If
If VelFPS = 1 Then
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "VelFPS", 1)
Else
Call WriteVar(App.Path & "\Init\opciones.dat", "OPCIONES", "VelFPS", 0)
End If


Unload frmOpciones
 frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 0, 0, 1, 0, 1
        
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Argentum Online!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
End Sub

Private Sub Image6_Click()
Me.Visible = False
End Sub

Private Sub Image7_Click()
Me.Visible = False
End Sub

Private Sub Image70_Click(Index As Integer)
Call audio.PlayWave(SND_CLICK)
Select Case Index
Case 0
VelFPS = 1
Case 1
VelFPS = 0
End Select

If VelFPS = 1 Then
   
    Image70(0).Picture = LoadPicture(DirGraficos & "validar.jpg")
    Image70(1).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Else
    Image70(0).Picture = LoadPicture(DirGraficos & "novalidar.jpg")
    Image70(1).Picture = LoadPicture(DirGraficos & "validar.jpg")
 
End If
End Sub

Private Sub Image8_Click()
frmCustomKeys.Show
End Sub

Private Sub Slider1_Change(Index As Integer)
    Select Case Index
        Case 0
            audio.MusicVolume = Slider1(0).Value
        Case 1
            audio.SoundVolume = Slider1(1).Value
    End Select
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    Select Case Index
        Case 0
            audio.MusicVolume = Slider1(0).Value
        Case 1
            audio.SoundVolume = Slider1(1).Value
    End Select
End Sub

