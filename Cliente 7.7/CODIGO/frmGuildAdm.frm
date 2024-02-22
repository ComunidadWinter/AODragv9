VERSION 5.00
Begin VB.Form frmGuildAdm 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6030
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildAdm.frx":0000
   ScaleHeight     =   4545
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2130
      ItemData        =   "frmGuildAdm.frx":1B6DC
      Left            =   240
      List            =   "frmGuildAdm.frx":1B6DE
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmGuildAdm.frx":1B6E0
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2175
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2040
      MouseIcon       =   "frmGuildAdm.frx":1B7C2
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   240
      MouseIcon       =   "frmGuildAdm.frx":1C48C
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4200
      MouseIcon       =   "frmGuildAdm.frx":1D156
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   1695
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
frmGuildAdm.Picture = LoadPicture(DirGraficos & "interfazclanes.jpg")
End Sub




Public Sub ParseGuildList(ByVal Rdata As String)
'pluto:2.4
Dim j As Integer, k As Integer
Dim Guildpuntos(1 To 250)
Dim GuildName(1 To 250)
k = CInt(ReadField(1, Rdata, 44))

For j = 1 To k
    guildslist.AddItem ReadField(1 + j, Rdata, 44)
Next j
GoTo a:

'MsgBox (Rdata)
For j = 1 To k
Guildpuntos(j) = Val(ReadField(j + 2 + k, Rdata, 44))
GuildName(j) = ReadField(j + 1, Rdata, 44)
Next

Dim i As Integer, e As Integer
For e = 1 To k
      For i = 1 To k
    
           If Guildpuntos(i) < Guildpuntos(e) Then
              NomAux = GuildName(i)
              GuildName(i) = GuildName(e)
              GuildName(e) = NomAux
           
              DNIAux = Guildpuntos(i)
              Guildpuntos(i) = Guildpuntos(e)
              Guildpuntos(e) = DNIAux
           End If
      
    Next i
Next e

' Vacío los ListBox
frmGuildAdm.guildslist.Clear

' Cargo los ListBox con los que contienen datos
For i = 1 To k
  If GuildName(i) <> "" Then
     frmGuildAdm.guildslist.AddItem GuildName(i) & "-->" & Guildpuntos(i)
     'frmordenado.lstdni.AddItem DNI(I)
     
  End If
Next
a:
Me.Show

End Sub

Private Sub Image1_Click()
Call SendData("CLANDETAILS" & guildslist.List(guildslist.ListIndex))

End Sub

Private Sub Image2_Click()
Unload Me
'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
End Sub

Private Sub Image3_Click()
Unload Me
Call SendData("CL8")
End Sub
