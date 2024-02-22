VERSION 5.00
Begin VB.Form frmOldPersonaje 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Introduce los datos de tu cuenta"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   6150
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOldPersonaje.frx":0000
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      MouseIcon       =   "frmOldPersonaje.frx":2CF82
      MousePointer    =   99  'Custom
      Picture         =   "frmOldPersonaje.frx":2DC4C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      MouseIcon       =   "frmOldPersonaje.frx":2E901
      MousePointer    =   99  'Custom
      Picture         =   "frmOldPersonaje.frx":2F5CB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox NameTxt 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   3000
      Width           =   3330
   End
   Begin VB.TextBox PasswordTxt 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3840
      Width           =   3330
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor fundado en el año 2001"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "www.juegosdrag.es"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "Introduce los"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de tu cuenta"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Correo Electrónico"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmOldPersonaje"
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



Private Sub Command1_Click()
     Call audio.PlayWave(SND_CLICK)
       'pluto:2.5.0
        If frmMain.Socket1.Connected Then
        Call frmMain.Socket1.Disconnect
        End If
        
        If frmConnect.MousePointer = 11 Then
            Exit Sub
        End If
        
'pluto:2.5.0
KeyCodi = ""
Keycodi2 = ""
        'update user info
        UserName = NameTxt.Text
        'PLUTO:2.11
        frmCuentas.Label1(1).Caption = LCase$(UserName)
        'frmCuentas.Conectar.Caption = "Entrar con: "

        Dim aux As String
        aux = PasswordTxt.Text
        UserPassword = MD5String(aux)
        If CheckUserData(False) = True Then
        
'pluto:6.3----------------------------------------------
'antes de conectarme el winsock debe estar cerrado
'frmMain.ws_cliente.Close
'me conecto a la pc remota utilizando su ip(la que ponemos en el textbox), x el puerto 8888
'If ServActual = 2 Then
'frmMain.ws_cliente.Connect "redpluto.no-ip.org", "7665"
'frmMain.ws_cliente.Connect "200.85.152.50", "10290"
'Else
'frmMain.ws_cliente.Connect "92.43.20.27", "7667"
'End If
'Me.ws_cliente.Connect "redpluto.no-ip.org", "7667"
'Debug.Print frmMain.ws_cliente.State
'---------------------------------------------------------
        
        
        
        
        
'            Load frmMain.Socket1

            frmMain.Socket1.HostName = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            SendNewChar = False
            Me.MousePointer = 11
            Debug.Print (CurServerIp)
            Debug.Print (CurServerPort)
'pluto:6.7
'UserRecibe = 0
'UserEnvia = 1
            frmMain.Socket1.Connect
        End If
End Sub

Private Sub Command2_Click()
Call audio.PlayWave(SND_CLICK)
Me.Visible = False
End Sub

Private Sub Form_Load()
Dim j



NameTxt.Text = ""
PasswordTxt.Text = ""
frmOldPersonaje.Picture = LoadPicture(App.Path & "\Graficos\ventanas.jpg")
'Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotonVolver.jpg")
'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotonSiguiente.jpg")

End Sub






Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Command1_Click
    End If
End Sub
