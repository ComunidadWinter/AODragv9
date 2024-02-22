VERSION 5.00
Begin VB.Form FrmPuntosClanes 
   BorderStyle     =   0  'None
   Caption         =   "Listado de los clanes más Poderosos de Aodrag"
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListLevel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      ItemData        =   "FrmPuntosClanes.frx":0000
      Left            =   2640
      List            =   "FrmPuntosClanes.frx":0007
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ListBox ListNombres 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      ItemData        =   "FrmPuntosClanes.frx":0016
      Left            =   360
      List            =   "FrmPuntosClanes.frx":001D
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ListBox Listpuntos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      ItemData        =   "FrmPuntosClanes.frx":002E
      Left            =   3960
      List            =   "FrmPuntosClanes.frx":0030
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2160
      MouseIcon       =   "FrmPuntosClanes.frx":0032
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos del Clan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel del Clan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Clan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "FrmPuntosClanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
FrmPuntosClanes.Picture = LoadPicture(DirGraficos & "interfazpoderos.jpg")
End Sub
Public Sub PuntosGuildList(ByVal Rdata As String)

Dim j As Integer, k As Integer
Dim Guildpuntos(1 To 250)
Dim GuildName(1 To 250)
'pluto:6.0A
Dim GuildLevel(1 To 250)

k = CInt(ReadField(1, Rdata, 44))


For j = 1 To k
Guildpuntos(j) = Val(ReadField(j + 1 + k, Rdata, 44))
GuildName(j) = ReadField(j + 1, Rdata, 44)
'pluto:6.0A
GuildLevel(j) = Val(ReadField(j + 1 + k + k, Rdata, 44))
Next

Dim i As Integer, e As Integer
For e = 1 To k
      For i = 1 To k
    
           If Guildpuntos(i) < Guildpuntos(e) Then
              NomAux = GuildName(i)
              GuildName(i) = GuildName(e)
              GuildName(e) = NomAux
              'pluto:6.0A
              LevelAux = GuildLevel(i)
              GuildLevel(i) = GuildLevel(e)
              GuildLevel(e) = LevelAux
           
              DNIAux = Guildpuntos(i)
              Guildpuntos(i) = Guildpuntos(e)
              Guildpuntos(e) = DNIAux
           End If
      
    Next i
Next e

' Vacío los ListBox
frmGuildAdm.guildslist.Clear
FrmPuntosClanes.Listpuntos.Clear
FrmPuntosClanes.ListLevel.Clear
FrmPuntosClanes.ListNombres.Clear
' Cargo los ListBox con los que contienen datos
Dim a1 As Byte
Dim a2 As Byte
For i = 1 To 14

  If GuildName(i) <> "" Then
  'a1 = 20 - Len(GuildName(i))
  'If Len(GuildName(i)) < 6 Then GuildName(i) = GuildName(i) + "     "
     FrmPuntosClanes.ListNombres.AddItem i & "- " & GuildName(i)
     FrmPuntosClanes.Listpuntos.AddItem Guildpuntos(i)
     FrmPuntosClanes.ListLevel.AddItem GuildLevel(i)
          'frmordenado.lstdni.AddItem DNI(I)
     End If
Next
a:
Me.Show

End Sub


Private Sub Image1_Click()
Unload Me
End Sub
