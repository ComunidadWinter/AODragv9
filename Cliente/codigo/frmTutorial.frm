VERSION 5.00
Begin VB.Form frmTutorial 
   Caption         =   "¡Primeros Pasos!"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12675
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   4290
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "Siguiente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   3
      Top             =   4560
      Width           =   7890
   End
   Begin VB.Timer TimerCierre 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   3750
      Left            =   120
      ScaleHeight     =   3690
      ScaleWidth      =   4245
      TabIndex        =   0
      Top             =   600
      Width           =   4305
   End
   Begin VB.Label lblMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3750
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   7965
   End
   Begin VB.Label lblTitulo 
      Caption         =   "¡Bienvenido a las tierras del Dragón!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12255
   End
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tTutorial
    sTitle As String
    sPage As String
End Type

Private Tutorial() As tTutorial
Private NumPages As Long
Private CurrentPage As Long
Public DesdElMain As Boolean

Private Sub LoadTutorial()

    Dim Leer As New clsIniReader
    Dim TutorialPath As String
    Dim lPage As Long
    Dim NumLines As Long
    Dim lLine As Long
    Dim sLine As String
    
    TutorialPath = Get_Extract(Scripts, "Tutorial.dat")
    Call Leer.Initialize(TutorialPath)
    
    NumPages = Val(Leer.GetValue("INIT", "NumPags"))
    
    If NumPages > 0 Then
        ReDim Tutorial(1 To NumPages)
        
        ' Cargo paginas
        For lPage = 1 To NumPages
            NumLines = Val(Leer.GetValue("PAG" & lPage, "NumLines"))
            
            With Tutorial(lPage)
                
                .sTitle = Leer.GetValue("PAG" & lPage, "Title")
                
                ' Cargo cada linea de la pagina
                For lLine = 1 To NumLines
                    sLine = Leer.GetValue("PAG" & lPage, "Line" & lLine)
                    .sPage = .sPage & sLine & vbCrLf
                Next lLine
            End With
            
        Next lPage
    End If
End Sub

Private Sub cmdSiguiente_Click()
    
    If CurrentPage = NumPages And DesdElMain = True Then Unload Me: Exit Sub
    
    CurrentPage = CurrentPage + 1
    
    ' DEshabilita el boton siguiente si esta en la ultima pagina
    
    If CurrentPage = NumPages Then
        If DesdElMain = False Then
            TimerCierre.Enabled = True
            cmdSiguiente.Enabled = False
        Else
            cmdSiguiente.Caption = "Salir"
        End If
    End If
    
    ' Habilita el boton anterior
    If Not cmdVolver.Enabled Then cmdVolver.Enabled = True
    
    Call SelectPage(CurrentPage)
    
End Sub

Private Sub cmdVolver_Click()
    
    CurrentPage = CurrentPage - 1
    
    cmdSiguiente.Caption = "Siguiente"
    
    If CurrentPage = 1 Then cmdVolver.Enabled = False
    
    If Not cmdSiguiente.Enabled Then cmdSiguiente.Enabled = True
    
    Call SelectPage(CurrentPage)
End Sub

Private Sub Form_Load()
    Call LoadTutorial
    
    CurrentPage = 1
    Call SelectPage(CurrentPage)
End Sub

Private Sub SelectPage(ByVal lPage As Long)
    lblTitulo.Caption = lPage & "/" & NumPages & " - " & Tutorial(lPage).sTitle
    lblMensaje.Caption = Tutorial(lPage).sPage
    Picture1.Picture = General_Load_Picture_From_Resource("4" & lPage + 1 & ".gif")
End Sub

Private Sub TimerCierre_Timer()
    Static Segundo As Byte
    
    Segundo = Segundo + 1
    cmdSiguiente.Caption = "Cerrando en: " & 10 - Segundo
    
    If Segundo = 10 Then
        Segundo = 0
        Unload Me
        If DesdElMain = False Then frmElegirControles.Show , frmMain
        DesdElMain = False
    End If
    
    If Not CurrentPage = NumPages Then
        Segundo = 0
        cmdSiguiente.Caption = "Siguiente"
        TimerCierre.Enabled = False
    End If
    
End Sub
