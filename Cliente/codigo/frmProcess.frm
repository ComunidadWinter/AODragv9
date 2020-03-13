VERSION 5.00
Begin VB.Form FrmCaption 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Visor de CAPTIONs"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   480
      TabIndex        =   2
      Top             =   5400
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   4620
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Captions de"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "FrmCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible _
    Lib "user32" ( _
        ByVal hwnd As Long) As Long

'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength _
    Lib "user32" _
    Alias "GetWindowTextLengthA" ( _
        ByVal hwnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText _
    Lib "user32" _
    Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow _
    Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&
Public CANTv As Byte

Public Function Listar() As String
Static alter As String
Dim buf As Long, Handle As Long, titulo As String, lenT As Long, ret As Long
    'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
    Handle = GetWindow(hwnd, GW_HWNDFIRST)

    'Este bucle va a recorrer todas las ventanas.
    'cuando GetWindow devielva un 0, es por que no hay mas
    Do While Handle <> 0
        'Tenemos que comprobar que la ventana es una de tipo visible
        If IsWindowVisible(Handle) Then
            'Obtenemos el número de caracteres de la ventana
            lenT = GetWindowTextLength(Handle)
            'si es el número anterior es mayor a 0
            If lenT > 0 Then
                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                titulo = String$(lenT, 0)
                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                'y también debemos pasarle el Hwnd de dicha ventana
                ret = GetWindowText(Handle, titulo, lenT + 1)
                titulo$ = Left$(titulo, ret)
                'La agregamos al ListBox
                Listar = titulo & "#" & Listar
                CANTv = CANTv + 1
            End If
        End If
        'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
        Handle = GetWindow(Handle, GW_HWNDNEXT)
       Loop
End Function

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

Debug.Print Listar
Dim raa As String
raa = Listar
Debug.Print raa

End Sub


