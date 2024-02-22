VERSION 5.00
Begin VB.Form frmAdverten 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Advertencia"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form3"
   ScaleHeight     =   7185
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mi Nombre de Personaje no cumple esta norma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      MouseIcon       =   "frmAdverten.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmAdverten.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mi nombre de Personaje cumple esta norma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      MouseIcon       =   "frmAdverten.frx":197F
      MousePointer    =   99  'Custom
      Picture         =   "frmAdverten.frx":2649
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quedan Prohibidos y serán Borrados"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   5295
   End
End
Attribute VB_Name = "frmAdverten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vete As Byte

Private Sub Command1_Click()
Call audio.PlayWave(SND_CLICK)
vete = vete + 1
Select Case vete
Case 1
Label1.Caption = "Nombres con Adornos de cualquier tipo delante o detrás del nombre del personaje. Ejemplos: XIX, oOo , III , OIO ..."
Case 2
Label1.Caption = "Nombres con la clase del personaje delante o detrás del nombre. Ejemplos: Talador DraG, Lohtus Guerrero, Arquero Powa..."
Case 3
Label1.Caption = "Nombres o alusiones a grupos políticos o religiosos de cualquier índole, bandas callejeras, criminales, personas o grupos de cualquier tipo que estuvieron o están sometidas a juicio, bajo sospecha o imputadas por crímenes."
Case 4
Label1.Caption = "Nombres que contengan el del servidor o el de otros servidores."
Case 5
Label1.Caption = "Nombres que por su parecido con el de algún GM puedan llevar a confusión ( revisar la lista de GMs en el momento de creación del pje para evitarlo)."
Case 6
Label1.Caption = "Nombres de contenido racista , xenófobo o discriminatorio."
Case 7
Label1.Caption = "Nombres o juegos de palabras que se intuyan no apropiados."
Case 8
NameCorrecto = True
vete = 0
Unload Me
End Select

End Sub


Private Sub Command2_Click()
vete = 0
NameCorrecto = False

Unload Me
MsgBox ("Si el nombre elegido no cumple las normas cambialo o será borrado sin previo aviso.")
End Sub

Private Sub Form_Load()
frmAdverten.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
Label1.Caption = "Nombres que contengan; insultos, abreviaturas de éstos, alusiones a drogas o estados inducidos por ellas, referencias soeces , ordinarias, groseras, escatológicas o sexuales , menciones a estados de discapacidad física o mental."
Label3.Caption = "Has elegido " & frmCrearPersonaje.txtNombre & " como nombre para el personaje."
Call audio.StopWave
Call audio.PlayWave("nombre.wav")
End Sub

