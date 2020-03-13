VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Mapa del Mundo AodraG"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   DrawStyle       =   1  'Dash
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "frmMapa.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "14"
      Height          =   255
      Index           =   4
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   2
      Top             =   7440
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Top             =   6720
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7440
      TabIndex        =   8
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7440
      TabIndex        =   7
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   5
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Criaturas del Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Info del Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mapa8 As String
Dim Dir7 As String
Dim Vezc As Byte

Private Sub Command1_Click()
busca7 = Text3.Text
Dir7 = App.Path
Mapa8 = "Mapa" & busca7
Label6.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "bichos")
Label5.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "Info")
Label4.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "nombre")
Option1(busca7).value = True
End Sub

Private Sub Form_Load()

Dim n As Integer
Dim colortipo As String
Vezc = Vezc + 1
Mapa8 = "Mapa" & n
Dir7 = App.Path


For n = 1 To 303
    Mapa8 = "Mapa" & n
    Option1(n).Caption = n
    colortipo = GetVar(Dir7 & "\Data.txt", Mapa8, "tipo")

    Select Case colortipo
        Case "Agua"
            Option1(n).BackColor = RGB(0, 128, 192)
        Case "Bosque"
            Option1(n).BackColor = RGB(0, 128, 128)
        Case "Poblado"
            Option1(n).BackColor = RGB(204, 102, 204)
            'Option1(n).Picture = LoadPicture(App.Path & "\Graficos\azul.bmp")
        Case "Costa"
            Option1(n).BackColor = RGB(150, 200, 150)
        Case "Dungeon"
            Option1(n).BackColor = RGB(200, 140, 130)
        Case "Ciudad"
            Option1(n).BackColor = RGB(255, 170, 85)
       If n = 1 And DueñoUlla = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       If n = 1 And DueñoUlla = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       If n = 20 And DueñoDesierto = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       If n = 20 And DueñoDesierto = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")

       If n = 59 And DueñoBander = 1 Then
       Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       Option1(58).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       Option1(60).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       Option1(61).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       Option1(66).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")

       End If

       If n = 59 And DueñoBander = 2 Then
       Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       Option1(58).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       Option1(60).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       Option1(61).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       Option1(66).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       End If

       If n = 62 And DueñoLindos = 1 Then
       Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       Option1(63).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       Option1(64).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       End If

       If n = 62 And DueñoLindos = 2 Then
       Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       Option1(63).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       Option1(64).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       End If

       If n = 34 And DueñoNix = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       If n = 34 And DueñoNix = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       If n = 81 And DueñoDescanso = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       If n = 81 And DueñoDescanso = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")

       If n = 84 And DueñoAtlantis = 1 Then
       Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       Option1(83).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       Option1(85).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       End If

       If n = 84 And DueñoAtlantis = 2 Then
       Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       Option1(83).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       Option1(85).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       End If

       If (n = 111 Or n = 112) And DueñoEsperanza = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       If (n = 111 Or n = 112) And DueñoEsperanza = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")
       If (n = 150 Or n = 151) And DueñoArghal = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       If (n = 150 Or n = 151) And DueñoArghal = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")

       If n = 157 And DueñoQuest = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       If n = 157 And DueñoQuest = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")


       If n = 170 And DueñoCaos = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
       If n = 170 And DueñoCaos = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")

If n = 20 And DueñoDesierto = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
If n = 20 And DueñoDesierto = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")


         If (n = 183 Or n = 184) And DueñoLaurana = 1 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Real.bmp")
        If (n = 183 Or n = 184) And DueñoLaurana = 2 Then Option1(n).Picture = LoadPicture(App.Path & "\Graficos\Caos.bmp")


                Option1(n).ForeColor = vbWhite
       
        Case "Desierto"
            Option1(n).BackColor = &H80FFFF
        Case "Castillo"
            Option1(n).BackColor = RGB(205, 128, 0)
        Case "Quest"
            Option1(n).BackColor = &HFF00FF
        Case "Mina"
            Option1(n).BackColor = RGB(192, 192, 192)
        Case "Piramide"
            Option1(n).BackColor = RGB(200, 140, 130)
        Case "Nieve"
            Option1(n).BackColor = &HFFFFFF
        Case "Isla"
            Option1(n).BackColor = RGB(220, 185, 185)
        Case "Encantada"
            Option1(n).BackColor = RGB(200, 140, 130)
        Case "Alquimista"
            Option1(n).BackColor = RGB(200, 140, 130)
        Case Else
            Option1(n).Enabled = False
            Option1(n).Visible = False
    End Select
Next

End Sub

Private Sub Option1_Click(Index As Integer)
    If Vezc > 0 Then
Vezc = 0
Index = UserMap
If Option1(Index).Enabled = True And Option1(Index).Visible = True Then
Option1(Index).SetFocus
End If
    End If
busca7 = Index
Dir7 = App.Path
Mapa8 = "Mapa" & busca7
Label6.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "bichos")
Label5.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "Info")
Label4.Caption = GetVar(Dir7 & "\Data.txt", Mapa8, "nombre")
Text3.Text = busca7
End Sub


