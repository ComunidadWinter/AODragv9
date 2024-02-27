VERSION 5.00
Begin VB.Form frmCustomKeys 
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   713
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cerrar sin guardar"
      Height          =   360
      Left            =   8760
      TabIndex        =   76
      Top             =   5160
      Width           =   1770
   End
   Begin VB.CommandButton cmdGuardaryCerrar 
      Caption         =   "Guardar y Cerrar"
      Height          =   360
      Left            =   6480
      TabIndex        =   75
      Top             =   5160
      Width           =   2130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cargar configuración clásica"
      Height          =   360
      Left            =   3360
      TabIndex        =   74
      Top             =   5160
      Width           =   2970
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar configuración W, A, S, D"
      Height          =   360
      Left            =   240
      TabIndex        =   73
      Top             =   5160
      Width           =   2970
   End
   Begin VB.Frame Frame6 
      Caption         =   "Macros de Comandos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   8160
      TabIndex        =   44
      Top             =   120
      Width           =   2415
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   12
         Left            =   600
         TabIndex        =   67
         Text            =   "Text2"
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   11
         Left            =   600
         TabIndex        =   65
         Text            =   "Text2"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   10
         Left            =   600
         TabIndex        =   63
         Text            =   "Text2"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   9
         Left            =   600
         TabIndex        =   61
         Text            =   "Text2"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   8
         Left            =   600
         TabIndex        =   59
         Text            =   "Text2"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   7
         Left            =   600
         TabIndex        =   57
         Text            =   "Text2"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   6
         Left            =   600
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   5
         Left            =   600
         TabIndex        =   53
         Text            =   "Text2"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   600
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   600
         TabIndex        =   49
         Text            =   "Text2"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   45
         Text            =   "Text2"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "F12/"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   68
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F11/"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   66
         Top             =   3990
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F10/"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   64
         Top             =   3630
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F9 /"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   62
         Top             =   3270
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F8 /"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   60
         Top             =   2910
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F7 /"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   58
         Top             =   2550
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F6 /"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   56
         Top             =   2190
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F5 /"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   54
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F4 /"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   52
         Top             =   1470
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F3 /"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   50
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F2 /"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   48
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "F1 /"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   390
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   2280
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   2280
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   2280
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   2280
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   2280
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar Mapa "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   72
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir Menú "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   70
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Modo Combate"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   42
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Capturar Pantalla"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar/Ocultar FPS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Hablar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   3960
      TabIndex        =   3
      Top             =   3840
      Width           =   4095
      Begin VB.CheckBox ChkMovement 
         Caption         =   "Desactivar movimiento en escritura"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   2280
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hablar a Todos"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   1920
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1920
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1920
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   1920
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Atacar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Usar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tirar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ocultar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Robar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Domar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Equipar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Agarrar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones Personales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar/Ocultar Nombres"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Corregir Posicion"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Activar/Desactivar Musica"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movimiento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Derecha"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Izquierda"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Abajo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Arriba"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCustomKeys"
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

''
'frmCustomKeys - Allows the user to customize keys.
'Implements class clsCustomKeys
'
'@author Rapsodius
'@date 20070805
'@version 1.0.0
'@see clsCustomKeys

Option Explicit

Private Sub ChkMovement_Click()
    If ChkMovement.value = vbUnchecked Then
        Opciones.MovEscritura = 0
    ElseIf Not Opciones.MovEscritura Then  'Prevent the music from reloading
        Opciones.MovEscritura = 1
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGuardaryCerrar_Click()
    Dim i As Long
    
    For i = 1 To CustomKeys.count
        If LenB(Text1(i).Text) = 0 Then
            Call MsgBox("Hay una o mas teclas no validas, por favor verifique.", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Argentum Online")
            Exit Sub
        End If
    Next i
    
    Call CustomKeys.SaveCustomKeys

    For i = 1 To 12
        Call WriteVar(App.Path & "\Init\AODConfig.bnd", "CommandMacros", "F" & Text2(i).Index, Text2(i).Text)
    Next i

    Unload Me
End Sub

Private Sub Command1_Click()
    Call CustomKeys.LoadDefaults(False)
    Dim i As Long
    
    For i = 1 To CustomKeys.count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub

Private Sub Command3_Click()
    Call CustomKeys.LoadDefaults(True)
    Dim i As Long
    
    For i = 1 To CustomKeys.count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = 1 To CustomKeys.count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
    
    For i = 1 To 12
        Text2(i).Text = GetVar(App.Path & "\Init\AODConfig.bnd", "CommandMacros", "F" & (i))
    Next i
    
    If Opciones.MovEscritura = 1 Then
        ChkMovement.value = vbChecked
    Else
        ChkMovement.value = vbUnchecked
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    
    Text1(Index).Text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
    For i = 1 To CustomKeys.count
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                Text1(Index).Text = "" 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call Text1_KeyDown(Index, KeyCode, Shift)
End Sub
