VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Begin VB.Form FrmMainLauncher 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "AoDraG"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   Icon            =   "FrmMainLauncher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMainLauncher.frx":000C
   ScaleHeight     =   3585
   ScaleWidth      =   6795
   StartUpPosition =   1  'CenterOwner
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Bienvenido al mundo de AoDraG"
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Image ImgSalir 
      Height          =   615
      Left            =   3480
      Top             =   2880
      Width           =   3210
   End
   Begin VB.Image ImgJugar 
      Height          =   570
      Left            =   120
      Top             =   2040
      Width           =   6570
   End
   Begin VB.Image imgForo 
      Height          =   615
      Left            =   120
      Top             =   2880
      Width           =   3240
   End
End
Attribute VB_Name = "FrmMainLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Directorio As String = "\Graficos\Launcher\"

Dim f As Integer

Private Sub Form_Load()

   ' WebBrowser1.Navigate "http://181.47.237.170/Noticias.txt"
    ImgSalir.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Salir_N.jpg")
    ImgJugar.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Jugar_N.jpg")
    imgForo.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Foro_N.jpg")
    Me.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main.jpg")

    Call Analizar
    IPdelServidor = "181.47.237.170"    'Host
End Sub

Private Sub Analizar()
    On Error Resume Next
    Dim ix As Integer, tX As Integer, DifX As Integer

    'lEstado.Caption = "Obteniendo datos..."
    lblEstado.Caption = "Buscando Actualizaciones..."
    lblEstado.ForeColor = vbGreen

    'ix = Inet1.OpenURL("http://181.47.237.170/VEREXE.txt")    'Host
    tX = LeerInt(App.Path & "\INIT\AU.ini")


    DifX = ix - tX

    If Not (DifX = 0) Then
        lblEstado.Caption = "Hay " & DifX & " actualizaciones disponibles."
        lblEstado.ForeColor = vbRed
    Else
        lblEstado.Caption = "AoDraG está actualizado. Pulsa el botón Jugar."
        lblEstado.ForeColor = vbGreen
    End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ImgSalir.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Salir_N.jpg")
    ImgJugar.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Jugar_N.jpg")
    imgForo.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Foro_N.jpg")

End Sub

Private Sub imgForo_Click()
    Call ShellExecute(Me.hWnd, "Open", "https://www.facebook.com/groups/2599258103512061/", &O0, &O0, 1)
End Sub

Private Sub imgForo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgForo.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Foro_A.jpg")
End Sub

Private Sub imgForo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgForo.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Foro_I.jpg")
End Sub

Private Sub ImgJugar_Click()

    If lblEstado.ForeColor = vbRed Then
        Call MsgBox("Se abrirá el AutoUpdate para actualizar AoDraG a la versión mas actual.", vbInformation, "Atención")
        Unload FrmMainLauncher
        Call ShellExecute(Me.hWnd, "open", App.Path & "/AutoUpdate.exe", "", "", 1)
        End
        Exit Sub
    Else
        Unload FrmMainLauncher
        Call Main
        Exit Sub
    End If

End Sub

Private Sub ImgJugar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgJugar.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Jugar_A.jpg")
End Sub

Private Sub ImgJugar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgJugar.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Jugar_I.jpg")
End Sub

Private Sub ImgSalir_Click()

    End

End Sub

Private Sub ImgSalir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgSalir.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Salir_A.jpg")
End Sub

Private Sub ImgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgSalir.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Salir_I.jpg")
End Sub

Private Function LeerInt(ByVal Ruta As String) As Integer
    f = FreeFile
    Open Ruta For Input As f
    LeerInt = Input$(LOF(f), #f)
    Close #f
End Function

