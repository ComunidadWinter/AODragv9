VERSION 5.00
Begin VB.Form frmCrearCaracter 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1815
      TabIndex        =   3
      Top             =   2085
      Width           =   3735
   End
   Begin VB.CheckBox SavePassChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1920
      TabIndex        =   2
      Top             =   1305
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   1845
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   660
      Width           =   3735
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1845
      TabIndex        =   0
      Top             =   270
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   3660
      Tag             =   "0"
      Top             =   2475
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   1320
      Tag             =   "0"
      Top             =   2475
      Width           =   960
   End
End
Attribute VB_Name = "frmCrearCaracter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)
Select Case Index
    Case 0
        'Actualizamos la Info
            UserName = NameTxt.Text
            UserPassword = PasswordTxt.Text
            UserEmail = text1.Text
            Dim aux As String
            aux = UserPassword
            UserPassword = EncryptINI$(aux, Seed)
        
        If CheckUserData(True) = True Then
            Me.Visible = False
            frmRaza.Show vbModal
        End If
    Case 1
        Me.Visible = False
End Select
    
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Graficos\CrearCaracter.jpg")

Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")

UserName = ""
UserPassword = ""
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1(0).Tag = 1 Then
    Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
    Image1(0).Tag = 0
End If
If Image1(1).Tag = 1 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")
    Image1(1).Tag = 0
End If
  
End Sub

Private Sub Image1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0
        If Musica = 0 Then
            CurMidi = DirMidi & "2.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If
        
        frmConnect.FONDO.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        Me.Visible = False
        
    Case 1
        'Me.Visible = False
        'Actualizamos la Info
            UserName = NameTxt.Text
            UserPassword = PasswordTxt.Text
            UserEmail = text1.Text
            Dim aux As String
            aux = UserPassword
            UserPassword = EncryptINI$(aux, Seed)
        
        If CheckUserData(True) = True Then
            Me.Visible = False
            frmRaza.Show vbModal
        End If
End Select

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
    If Image1(0).Tag = 0 Then
            Call PlayWaveDS(SND_OVER)
            Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónVolverapretado.jpg")
            Image1(0).Tag = 1
    End If
Case 1
    If Image1(1).Tag = 0 Then
            Call PlayWaveDS(SND_OVER)
            Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguienteapretado.jpg")
            Image1(1).Tag = 1
    End If
    
End Select
    
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Image1(0).Tag = 1 Then
            Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónVolver.jpg")
            Image1(0).Tag = 0
  End If
  If Image1(1).Tag = 1 Then
    Image1(1).Picture = LoadPicture(App.Path & "\Graficos\BotónSiguiente.jpg")
    Image1(1).Tag = 0
  End If

        
End Sub
