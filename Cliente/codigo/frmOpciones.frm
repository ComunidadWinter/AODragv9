VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8625
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
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCustomKeys 
      Caption         =   "Configurar Teclas"
      Height          =   360
      Left            =   5280
      TabIndex        =   30
      Top             =   2040
      Width           =   2730
   End
   Begin VB.CommandButton cmdManual 
      Caption         =   "Manual"
      Height          =   360
      Left            =   120
      TabIndex        =   29
      Top             =   5400
      Width           =   2730
   End
   Begin VB.CommandButton cmdWeb 
      Caption         =   "Web Oficial"
      Height          =   360
      Left            =   2925
      TabIndex        =   28
      Top             =   5400
      Width           =   2730
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   5760
      TabIndex        =   27
      Top             =   5400
      Width           =   2730
   End
   Begin VB.Frame Frame3 
      Caption         =   "Miscelánea"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4440
      TabIndex        =   4
      Top             =   0
      Width           =   4095
      Begin VB.CheckBox Check1 
         Caption         =   "Bloqueo de cruceta"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hechizos Clasicos (BETA)"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   15
         Top             =   1330
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar noticias de clan al conectar"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1030
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ver nombre de jugadores"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Desactivar deteccion de URL en consola"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Video"
      Height          =   2655
      Left            =   4440
      TabIndex        =   1
      Top             =   2640
      Width           =   4095
      Begin VB.Frame Frame4 
         Caption         =   "Vertex Processing:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   3855
         Begin VB.OptionButton optProcessing 
            Caption         =   "HARDWARE"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optProcessing 
            Caption         =   "SOFTWARE"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optProcessing 
            Caption         =   "MIXED"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Activar / Desactivar Pantalla Completa (BETA)"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Activar / Desactivar Sincronización Vertical"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   3615
      End
      Begin VB.HScrollBar SldTechos 
         Height          =   315
         LargeChange     =   20
         Left            =   240
         Max             =   255
         SmallChange     =   2
         TabIndex        =   3
         Top             =   600
         Value           =   255
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Transparencia de techos:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Audio"
      ForeColor       =   &H00000000&
      Height          =   3780
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.HScrollBar scrAmbient 
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   26
         Top             =   2760
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Invertir los canales (L/R)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   1580
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Sonidos Ambientales habilitado"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1290
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Música habilitada"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Efectos de navegación"
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   2955
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Sonido habilitado"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   660
         Width           =   2985
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   17
         Top             =   2130
         Width           =   2895
      End
      Begin VB.HScrollBar scrMidi 
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   16
         Top             =   3300
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de audio"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   2835
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de sonidos ambientales"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Top             =   2520
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de música"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   18
         Top             =   3090
         Width           =   2865
      End
   End
   Begin VB.Label lblInfo 
      Caption         =   "Información:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   4095
   End
End
Attribute VB_Name = "frmOpciones"
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

Private loading As Boolean
Dim Reiniciar As Boolean

Private Sub Check1_Click(Index As Integer)
    If Not loading Then _
        Call Sound.Sound_Play(SND_CLICK)
    
    Select Case Index
        Case 0
            If Check1(4).value = vbUnchecked Then
                Opciones.BloqCruceta = 0
            Else
                Opciones.BloqCruceta = 1
            End If
            lblInfo.Caption = "Información: La cruceta se bloqueara mientras dure el intervalo de hechizos."
            
        Case 4
            If Check1(4).value = vbUnchecked Then
                Opciones.URLCON = 0
                Call StopURLDetect
            Else
                Opciones.URLCON = 1
                #If Desarrollo = 0 Then
                Call StartURLDetect(frmMain.RecTxt.hwnd, frmMain.hwnd)
                #End If
            End If
            lblInfo.Caption = "Información: Cuando escribes una dirección web en consola, este la mostrara como un vinculo."
            
        Case 5
            If Check1(5).value = vbUnchecked Then
                Opciones.NamePlayers = 0
                
            Else
                Opciones.NamePlayers = 1
            End If
            lblInfo.Caption = "Información: Desactiva o Activa los nombres de los personajes."
            
        Case 6
            If Check1(6).value = vbUnchecked Then
                Opciones.GuildNews = 0
                
            Else
                Opciones.GuildNews = 1
            End If
            lblInfo.Caption = "Información: Desactiva o Activa las noticias de clan al iniciar (solo para lideres de clanes)."
            
        Case 7
            If Check1(7).value = vbUnchecked Then
                Opciones.VSynC = 0
            Else
                Opciones.VSynC = 1
            End If
            Reiniciar = True
            lblInfo.Caption = "Información: Activa o Desactiva la Sincronización Vertical (más información en el manual del juego)."
            
        Case 8
            If Check1(8).value = vbUnchecked Then
                Opciones.NoRes = 0
            Else
                Opciones.NoRes = 1
            End If
            lblInfo.Caption = "Información: Activa o Desactiva el modo pantalla completa (debes reiniciar el cliente)."

            Reiniciar = True
            
        Case 9
            If Check1(9).value = vbUnchecked Then
                Opciones.HechizosClasicos = 0
            Else
                Opciones.HechizosClasicos = 1
            End If
            
            If Opciones.HechizosClasicos Then
                frmMain.hlst.Visible = True
                frmMain.cmdLanzar.Visible = True
                frmMain.cmdInfo.Visible = True
                frmMain.picSpell.Visible = False
                frmMain.cmdMoverHechi(1).Visible = True
                frmMain.cmdMoverHechi(0).Visible = True
            Else
                frmMain.hlst.Visible = False
                frmMain.cmdLanzar.Visible = False
                frmMain.cmdInfo.Visible = False
                frmMain.picSpell.Visible = True
                frmMain.cmdMoverHechi(1).Visible = False
                frmMain.cmdMoverHechi(0).Visible = False
            End If
            Reiniciar = True
            
            lblInfo.Caption = "Información: Activa o Desactiva el modo de ver los hechizos en lista o iconos."
    End Select
End Sub

Private Sub chkop_Click(Index As Integer)
Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
        Case 0
                    
            If chkop(Index).value = vbUnchecked Then
                Sound.Music_Stop
                Opciones.sMusica = CONST_DESHABILITADA
                scrMidi.Enabled = False
            Else
                Opciones.sMusica = CONST_MP3
                scrMidi.Enabled = True
            End If
        
        Case 1
    
            If chkop(Index).value = vbUnchecked Then
                chkop(2).Enabled = False
                'scrAmbient.Enabled = False
                scrVolume.Enabled = False
                Opciones.Audio = 0
            Else
                Opciones.Audio = 1
                chkop(2).Enabled = True
                scrVolume.Enabled = True
            End If
    
        Case 2
    
            If chkop(Index).value = vbUnchecked Then
                Opciones.FxNavega = 0
            Else
                Opciones.FxNavega = 1
            End If
            
        Case 3
            
            If chkop(Index).value = vbUnchecked Then
                Opciones.Ambient = 0
                Call Sound.Sound_Stop_All
            Else
                Opciones.Ambient = 1
                scrAmbient.Enabled = True
                Call Sound.Ambient_Load(Sound.AmbienteActual, Opciones.AmbientVol)
                Call Sound.Ambient_Play
            End If
    End Select
End Sub

Private Sub cmdCustomKeys_Click()
    If Not loading Then _
        Call Sound.Sound_Play(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub cmdCerrar_Click()
    Call GuardarOpciones
    If Reiniciar = True Then
        MsgBox "Algunos de los cambios realizados no surtirán efecto hasta no reiniciar el cliente.", vbInformation
        Reiniciar = False
    End If
    Unload Me
End Sub

Private Sub cmdManual_Click()
    Call ShellExecute(0, "Open", "http://www.aodrag.es/wiki/", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub cmdWeb_Click()
    If Not loading Then _
        Call Sound.Sound_Play(SND_CLICK)
        Call ShellExecute(0, "Open", "http://www.aodrag.es", "", App.Path, SW_SHOWNORMAL)
End Sub

Public Sub Init()
    loading = True      'Prevent sounds when setting check's values
    
    If Opciones.sMusica = CONST_DESHABILITADA Then
        chkop(0).value = 0
        scrMidi.value = Opciones.MusicVolume
    Else
        chkop(0).value = 1
        scrMidi.value = Opciones.MusicVolume
    End If
    
    If Opciones.Audio = 1 Then
        chkop(1).value = vbChecked
        chkop(2).value = Opciones.FxNavega
        chkop(4).value = IIf(Opciones.InvertirSonido = True, 1, 0)
        scrVolume.value = Opciones.FXVolume
    Else
        chkop(1).value = vbUnchecked
        chkop(2).value = Opciones.FxNavega
        chkop(2).Enabled = False
        chkop(4).value = IIf(Opciones.InvertirSonido = True, 1, 0)
        chkop(4).Enabled = False
        scrVolume.value = Opciones.FXVolume
        scrVolume.Enabled = False
    End If
    
    If Opciones.Ambient = 1 Then
        chkop(3).value = vbChecked
        scrAmbient.value = Opciones.AmbientVol
    Else
        chkop(3).value = vbUnchecked
        scrAmbient.value = Opciones.AmbientVol
    End If
    
    If Opciones.URLCON Then
        Check1(4).value = vbChecked
    Else
        Check1(4).value = vbUnchecked
    End If
    
    If Opciones.NamePlayers Then
        Check1(5).value = vbChecked
    Else
        Check1(5).value = vbUnchecked
    End If
    
    If Opciones.GuildNews Then
        Check1(6).value = vbChecked
    Else
        Check1(6).value = vbUnchecked
    End If
    
    If Opciones.VSynC Then
        Check1(7).value = vbChecked
    Else
        Check1(7).value = vbUnchecked
    End If
    
    If Opciones.NoRes Then
        Check1(8).value = vbChecked
    Else
        Check1(8).value = vbUnchecked
    End If
    
    If Opciones.VProcessing = 0 Then
        optProcessing(0).value = vbChecked
    ElseIf Opciones.VProcessing = 1 Then
        optProcessing(1).value = vbChecked
    Else
        optProcessing(2).value = vbChecked
    End If
    
    SldTechos.value = Opciones.BaseTecho
    
    If Opciones.HechizosClasicos Then
        Check1(9).value = vbChecked
    Else
        Check1(9).value = vbUnchecked
    End If
    
    If Opciones.BloqCruceta Then
        Check1(0).value = vbChecked
    Else
        Check1(0).value = vbUnchecked
    End If
    
    If Not frmMain.Visible Then
        Me.Show vbModeless, frmConnect
    Else
        Me.Show vbModeless, frmMain
    End If
    
    Reiniciar = False
    
    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub optProcessing_Click(Index As Integer)
    Select Case Index
        Case 0
            Opciones.VProcessing = 0
        Case 1
            Opciones.VProcessing = 1
        Case 2
            Opciones.VProcessing = 2
    End Select
    
    'Si se hicieron cambios que requiera reiniciar lo notificaremos al guardar
    Reiniciar = True
End Sub

Private Sub scrMidi_Change()

    If Opciones.sMusica <> CONST_DESHABILITADA Then
        Sound.Music_Volume_Set scrMidi.value
        Sound.VolumenActualMusicMax = scrMidi.value
        Opciones.MusicVolume = Sound.VolumenActualMusicMax
    End If

End Sub

Private Sub scrAmbient_Change()
    If Opciones.Ambient = 1 Then
        Sound.VolumenActualAmbient_set scrAmbient.value
        Opciones.AmbientVol = Sound.VolumenActualAmbient
    End If
End Sub

Private Sub scrVolume_Change()

If Opciones.Audio = 1 Then
    Sound.VolumenActual = scrVolume.value
    Opciones.FXVolume = Sound.VolumenActual
End If

End Sub

Private Sub SldTechos_Change()
    Opciones.BaseTecho = SldTechos.value
End Sub

Private Sub SldTechos_click()
    lblInfo.Caption = "Información: Ajusta la opacidad de techos (!debes de estar bajo un techo para que los cambios surjan efecto!)"
End Sub
