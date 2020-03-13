VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton MasHead 
      Caption         =   ">"
      Height          =   480
      Left            =   4080
      TabIndex        =   12
      Top             =   3240
      Width           =   210
   End
   Begin VB.CommandButton MenosHead 
      Caption         =   "<"
      Height          =   480
      Left            =   2880
      TabIndex        =   11
      Top             =   3240
      Width           =   210
   End
   Begin VB.PictureBox PicHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3300
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   4
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   315
      TabIndex        =   3
      Top             =   615
      Width           =   7575
   End
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   360
      List            =   "frmCrearPersonaje.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2520
      Width           =   1620
   End
   Begin VB.ComboBox lstProfesion 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      ItemData        =   "frmCrearPersonaje.frx":001D
      Left            =   360
      List            =   "frmCrearPersonaje.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   1620
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      ItemData        =   "frmCrearPersonaje.frx":0021
      Left            =   360
      List            =   "frmCrearPersonaje.frx":0023
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3360
      Width           =   1620
   End
   Begin VB.Image cmdCrear 
      Height          =   255
      Left            =   5400
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Image cmdVolver 
      Height          =   255
      Left            =   1200
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblConstitucion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Constitución: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   10
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Label lblEnergia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Energia: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblInteligencia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inteligencia: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   8
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Label lblAgilidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agilidad: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   750
   End
   Begin VB.Label lblFuerza 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fuerza: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      Top             =   1680
      Width           =   690
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCrearPersonaje.frx":0025
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
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Image ImgProfesion 
      Height          =   3150
      Left            =   5085
      Top             =   1335
      Width           =   2505
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ModFuerza As Integer
Private ModAgilidad As Integer
Private ModEnergia As Integer
Private ModInteligencia As Integer
Private ModConstitucion As Integer
Private baseFuerza As Integer
Private baseAgilidad As Integer
Private baseEnergia As Integer
Private baseInteligencia As Integer
Private baseConstitucion As Integer
Private Const AtributoBase As Byte = 16

Private Sub CmdCrear_Click()
    Call Sound.Sound_Play(SND_CLICK)
    UserName = NameTxt.Text
            
    If Right$(UserName, 1) = " " Then
        UserName = RTrim$(UserName)
        MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
    End If
            
    UserRaza = lstRaza.ListIndex + 1
    UserSexo = lstGenero.ListIndex + 1
    UserClase = lstProfesion.ListIndex + 1
                
    EstadoLogin = E_MODO.CrearNuevoPj
    
    If frmMain.Winsock1.State <> sckConnected Then
        MsgBox "Error: Se ha perdido la conexion con el server."
        DoEvents
        
    Else
        PJName = NameTxt.Text
        Call Login
    End If
End Sub

Private Sub cmdVolver_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Unload Me
    frmCuenta.Show vbModeless, frmRenderConnect
    If Opciones.sMusica <> CONST_DESHABILITADA Then
        If Opciones.sMusica <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_VolverInicio
            Sound.Fading = 200
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    lstProfesion.Clear
    For i = LBound(ListaClases) To UBound(ListaClases)
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    lstRaza.Clear
    
    For i = LBound(ListaRazas()) To UBound(ListaRazas())
        lstRaza.AddItem ListaRazas(i)
    Next i
    
    lstProfesion.Clear
    
    For i = LBound(ListaClases()) To UBound(ListaClases())
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    lstProfesion.ListIndex = 1
    
    ImgProfesion.Picture = General_Load_Picture_From_Resource(lstProfesion.ListIndex + 1 & ".gif")
    
    Me.Picture = General_Load_Picture_From_Resource("31.gif")

End Sub

Function CheckData() As Boolean
If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If NameTxt = "" Then
    MsgBox "Debes de poner un nombre a tu personaje."
    Exit Function
End If

If UserSexo = 0 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = 0 Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

If Len(UserName) > 10 Then
    MsgBox ("El nombre debe tener menos de 10 letras.")
    Exit Function
End If

If Actual = 0 Then
    MsgBox "Selecciona una cabeza."
    Exit Function
End If

CheckData = True

End Function

Private Sub lstGenero_Click()
Call DameOpciones
Call DibujarHead
End Sub
 
Private Sub lstProfesion_click()
    ImgProfesion.Picture = General_Load_Picture_From_Resource(lstProfesion.ListIndex + 1 & ".gif")
    
    If lstProfesion.List(lstProfesion.ListIndex) = "Asesino" Then
        MsgBox "Esta clase esta deshabilitada temporalmente. Disculpa las molestias."
        lstProfesion.ListIndex = 1
        Exit Sub
    End If
    
    If lstProfesion.List(lstProfesion.ListIndex) = "Bardo" Then
        MsgBox "Esta clase esta deshabilitada temporalmente. Disculpa las molestias."
        lstProfesion.ListIndex = 1
        Exit Sub
    End If
    
    Select Case lstProfesion.ListIndex + 1
        Case 1
            ModFuerza = 0
            ModAgilidad = 1
            ModEnergia = 0
            ModInteligencia = 2
            ModConstitucion = 2
            
        Case 2
            ModFuerza = 1
            ModAgilidad = 0
            ModEnergia = 0
            ModInteligencia = 2
            ModConstitucion = 2
            
        Case 3
            ModFuerza = 1
            ModAgilidad = 0
            ModEnergia = 2
            ModInteligencia = 0
            ModConstitucion = 2
            
        Case 4
            ModFuerza = 0
            ModAgilidad = 1
            ModEnergia = 0
            ModInteligencia = 2
            ModConstitucion = 2
            
        Case 5
            ModFuerza = 1
            ModAgilidad = 0
            ModEnergia = 0
            ModInteligencia = 2
            ModConstitucion = 2
            
        Case 6
            ModFuerza = 0
            ModAgilidad = 1
            ModEnergia = 0
            ModInteligencia = 2
            ModConstitucion = 2
            
        Case 7
            ModFuerza = 1
            ModAgilidad = 0
            ModEnergia = 0
            ModInteligencia = 2
            ModConstitucion = 2
            
        Case 8
            ModFuerza = 0
            ModAgilidad = 1
            ModEnergia = 2
            ModInteligencia = 0
            ModConstitucion = 2
     End Select
     
     lblFuerza.Caption = "Fuerza: " & AtributoBase + baseFuerza + ModFuerza
     lblAgilidad.Caption = "Agilidad: " & AtributoBase + baseAgilidad + ModAgilidad
     lblenergia.Caption = "Energia: " & AtributoBase + baseEnergia + ModEnergia
     lblInteligencia.Caption = "Inteligencia: " & AtributoBase + baseInteligencia + ModInteligencia
     lblConstitucion.Caption = "Constitución: " & AtributoBase + baseConstitucion + ModConstitucion
End Sub

Private Sub lstRaza_Click()
Call DameOpciones
Call DibujarHead

  Select Case lstRaza.ListIndex + 1
     Case 1
            baseFuerza = 1
            baseAgilidad = 1
            baseEnergia = 2
            baseInteligencia = 1
            baseConstitucion = 2
            
        Case 2
            baseFuerza = 0
            baseAgilidad = 2
            baseEnergia = 1
            baseInteligencia = 2
            baseConstitucion = 1
            
        Case 3
            baseFuerza = 2
            baseAgilidad = 0
            baseEnergia = 1
            baseInteligencia = 2
            baseConstitucion = 1
            
        Case 4
            baseFuerza = -3
            baseAgilidad = 0
            baseEnergia = 1
            baseInteligencia = 3
            baseConstitucion = 1
            
        Case 5
            baseFuerza = 4
            baseAgilidad = -2
            baseEnergia = 0
            baseInteligencia = -5
            baseConstitucion = 3

        Case 6
            baseFuerza = 3
            baseAgilidad = -1
            baseEnergia = 1
            baseInteligencia = -3
            baseConstitucion = 3

        Case 7
            baseFuerza = 3
            baseAgilidad = 3
            baseEnergia = 0
            baseInteligencia = 2
            baseConstitucion = 0
   End Select
   
     lblFuerza.Caption = "Fuerza: " & AtributoBase + baseFuerza + ModFuerza
     lblAgilidad.Caption = "Agilidad: " & AtributoBase + baseAgilidad + ModAgilidad
     lblenergia.Caption = "Energia: " & AtributoBase + baseEnergia + ModEnergia
     lblInteligencia.Caption = "Inteligencia: " & AtributoBase + baseInteligencia + ModInteligencia
     lblConstitucion.Caption = "Constitución: " & AtributoBase + baseConstitucion + ModConstitucion
End Sub

Private Sub MenosHead_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Actual = Actual - 1
    If Actual > MaxEleccion Then
       Actual = MaxEleccion
    ElseIf Actual < MinEleccion Then
       Actual = MinEleccion
    End If
    Call DibujarHead
End Sub

Private Sub MasHead_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Actual = Actual + 1
    If Actual > MaxEleccion Then
       Actual = MaxEleccion
    ElseIf Actual < MinEleccion Then
       Actual = MinEleccion
    End If
    Call DibujarHead
End Sub

Private Sub NameTxt_Change()
    lblInfo.Caption = "En AoDrag premiamos a quienes utilizan nombres únicos y roleros. Elige un nombre rolero y podrás participar en el concurso de Rol y ganar fantásticos premios."
End Sub

Private Sub DibujarHead()
Dim Grh As Long

    Grh = HeadData(Actual).Head(3).GrhIndex
    
    PicHead.BackColor = PicHead.BackColor
    
    Call DrawGrhtoHdc(PicHead.hDC, Grh, PicHead, 10, 10)
End Sub
