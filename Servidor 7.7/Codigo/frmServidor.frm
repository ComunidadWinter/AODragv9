VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Servidor"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command16 
      Caption         =   "Reload Server.ini"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   585
      TabIndex        =   15
      Top             =   3825
      Width           =   3435
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Update MOTD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   585
      TabIndex        =   14
      Top             =   3465
      Width           =   3435
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Unban All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   570
      TabIndex        =   13
      Top             =   3135
      Width           =   3435
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Debug listening socket"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   12
      Top             =   2850
      Width           =   3435
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Debug Npcs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   11
      Top             =   2580
      Width           =   3435
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reiniciar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   10
      Top             =   855
      Width           =   3435
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ReSpawn Guardias en posiciones originales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   9
      Top             =   570
      Width           =   3435
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Stats de los slots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   555
      TabIndex        =   8
      Top             =   2325
      Width           =   3435
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Trafico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   7
      Top             =   2025
      Width           =   3435
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Reload Lista Nombres Prohibidos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   6
      Top             =   1740
      Width           =   3435
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Actualizar hechizos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   5
      Top             =   1440
      Width           =   3435
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Configurar intervalos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   4
      Top             =   1155
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar objetos.dat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   3
      Top             =   270
      Width           =   3435
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hacer un Backup del mundo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   510
      TabIndex        =   2
      Top             =   4725
      Width           =   3405
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cargar BackUp del mundo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   510
      TabIndex        =   1
      Top             =   5085
      Width           =   3405
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   3345
      TabIndex        =   0
      Top             =   5595
      Width           =   945
   End
   Begin VB.Shape Shape3 
      Height          =   4215
      Left            =   90
      Top             =   195
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   75
      Top             =   4515
      Width           =   4335
   End
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call LoadOBJData
End Sub

Private Sub Command10_Click()
frmTrafic.Show
End Sub

Private Sub Command11_Click()
frmConID.Show
End Sub

Private Sub Command12_Click()
frmDebugNpc.Show
End Sub

Private Sub Command13_Click()
frmDebugSocket.Visible = True
End Sub

Private Sub Command14_Click()
Call LoadMotd
End Sub

Private Sub Command15_Click()
On Error Resume Next

Dim Fn As String
Dim cad$
Dim n As Integer, k As Integer

Fn = App.Path & "\logs\GenteBanned.log"

If PersonajeExiste(Fn) Then
    n = FreeFile
    Open Fn For Input Shared As #n
    Do While Not EOF(n)
        k = k + 1
        Input #n, cad$
        Call UnBan(cad$)
        
    Loop
    Close #n
    MsgBox "Se han habilitado " & k & " personajes."
    Kill Fn
End If




End Sub

Private Sub Command16_Click()
Call LoadSini
End Sub

Private Sub Command2_Click()
frmServidor.Visible = False
End Sub

Private Sub Command3_Click()
If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la perdida de datos de los usarios. ¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbYes Then
    Me.Visible = False
    Call Restart
End If
End Sub

Private Sub Command4_Click()
On Error GoTo eh
    Me.MousePointer = 11
    FrmStat.Show
    Call DoBackUp
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
Exit Sub
eh:
Call LogError("Error en WORLDSAVE")
End Sub

Private Sub Command5_Click()

'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

FrmStat.Show

If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"

 #If UsarQueSocket = 1 Then
Call apiclosesocket(SockListen)
#ElseIf UsarQueSocket = 0 Then
frmMain.Socket1.Cleanup
frmMain.Socket2(0).Cleanup
#ElseIf UsarQueSocket = 2 Then
frmMain.Serv.Detener
#End If
  
Dim loopc As Integer

For loopc = 1 To MaxUsers
    Call CloseSocket(loopc)
Next
  

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call CargarBackUp
Call LoadOBJData


#If UsarQueSocket = 1 Then
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = Puerto
frmMain.Socket1.Listen
#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

End Sub

Private Sub Command6_Click()
Call ReSpawnOrigPosNpcs
End Sub

Private Sub Command7_Click()
FrmInterv.Show
End Sub

Private Sub Command8_Click()
Call CargarHechizos
End Sub

Private Sub Command9_Click()
Call CargarForbidenWords
End Sub

Private Sub Form_Deactivate()
frmServidor.Visible = False
End Sub

