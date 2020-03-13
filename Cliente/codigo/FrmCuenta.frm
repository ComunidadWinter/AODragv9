VERSION 5.00
Begin VB.Form frmCuenta 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox RenderPJPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   2865
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   1
      Top             =   1170
      Width           =   885
   End
   Begin VB.ListBox ListPJ 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2970
      ItemData        =   "FrmCuenta.frx":0000
      Left            =   330
      List            =   "FrmCuenta.frx":0002
      TabIndex        =   0
      Top             =   750
      Width           =   2175
   End
   Begin VB.Image cmdEntrar 
      Height          =   375
      Left            =   2400
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Image cmdcrearpj 
      Height          =   375
      Left            =   360
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Image cmdBorrarPj 
      Height          =   255
      Left            =   480
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Image cmdCambiarPass 
      Height          =   255
      Left            =   2400
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Image cmdSalir 
      Height          =   255
      Left            =   1425
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblMap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa Desconocido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clase - Raza - Nivel: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Image Marco 
      Height          =   1470
      Left            =   11640
      Picture         =   "FrmCuenta.frx":0004
      Top             =   9000
      Width           =   1065
   End
End
Attribute VB_Name = "frmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdborrarpj_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    frmBorrarPersonaje.Show vbModeless, frmCuenta
End Sub

Private Sub cmdCambiarPass_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    Call ShellExecute(0, "Open", "http://www.aodrag.es/", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub cmdcrearpj_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
        
    If Cuenta.CantPJ = 8 Then
        MsgBox "No tienes más espacio para continuar creando personajes."
        Exit Sub
    End If
        
    EstadoLogin = E_MODO.Dados
        
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
        
    If Opciones.sMusica <> CONST_DESHABILITADA Then
        If Opciones.sMusica <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_CrearPersonaje
            Sound.Fading = 500
        End If
    End If
            
    frmCuenta.Visible = False
End Sub

Private Sub cmdEntrar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    If ListPJ.ListIndex < 0 Then
        MsgBox "Selecciona antes un personaje."
        Exit Sub
    End If
    
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
            
    EstadoLogin = E_MODO.Normal
            
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
End Sub

Private Sub PJ_dblClick(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    If ListPJ.ListIndex < 0 Then
        MsgBox "Selecciona antes un personaje."
        Exit Sub
    End If
            
    EstadoLogin = E_MODO.Normal
            
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
    Exit Sub
End Sub

Private Sub cmdSalir_Click()
    Call Sound.Sound_Play(SND_CLICK)
    frmMain.Winsock1.Close
    EstadoLogin = Normal
    frmRenderConnect.btnConsejo.Visible = False
    Unload Me
    frmConnect.Show
End Sub

Private Sub Form_Load()
    Me.Picture = General_Load_Picture_From_Resource("30.gif")
End Sub

Private Sub ListPJ_Click()
    If (frmCuenta.ListPJ.ListIndex + 1) > frmCuenta.ListPJ.ListCount Then Exit Sub
    If ListPJ.ListIndex < 0 Then Exit Sub
    
    lblInfo.Caption = ListaClases(Cuenta.pjs(frmCuenta.ListPJ.ListIndex + 1).ClasePJ) & " " & ListaRazas(Cuenta.pjs(frmCuenta.ListPJ.ListIndex + 1).rcvRaza) & " Nivel: " & Cuenta.pjs(frmCuenta.ListPJ.ListIndex + 1).LvlPJ
    lblMap.Caption = "<" & Cuenta.pjs(ListPJ.ListIndex + 1).Mapa & ">"
    Engine.RenderPJCuenta
End Sub

Private Sub listpj_dblclick()
    If (frmCuenta.ListPJ.ListIndex + 1) > frmCuenta.ListPJ.ListCount Then Exit Sub
    If ListPJ.ListIndex < 0 Then Exit Sub
    
    Call cmdEntrar_Click
End Sub
