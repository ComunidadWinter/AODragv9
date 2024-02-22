VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Finalizar Aplicación (si tiene menú del sistema)"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMinimizeAll 
      Caption         =   "Minimizar todas las carpetas"
      Height          =   405
      Left            =   3990
      TabIndex        =   13
      Top             =   3510
      Width           =   2655
   End
   Begin VB.CommandButton cmdCerrarCarpetas 
      Caption         =   "Cerrar las carpetas"
      Height          =   405
      Left            =   3990
      TabIndex        =   12
      Top             =   3060
      Width           =   2655
   End
   Begin VB.CommandButton cmdRefrescar 
      Caption         =   "Actualizar lista de ventanas activas"
      Height          =   405
      Left            =   240
      TabIndex        =   9
      Top             =   3510
      Width           =   3645
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   405
      Left            =   7800
      TabIndex        =   3
      Top             =   3540
      Width           =   1245
   End
   Begin VB.CommandButton cmdCerrarVentanas 
      Caption         =   "Cerrar las ventanas seleccionadas"
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   3060
      Width           =   3645
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   570
      Width           =   6405
   End
   Begin VB.Label Label1 
      Caption         =   "ClassName:"
      Height          =   255
      Index           =   4
      Left            =   6750
      TabIndex        =   11
      Top             =   2370
      Width           =   1245
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   6750
      TabIndex        =   10
      Top             =   2610
      Width           =   2235
   End
   Begin VB.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   8
      Top             =   2040
      Width           =   945
   End
   Begin VB.Label Label2 
      Height          =   1065
      Index           =   0
      Left            =   6780
      TabIndex        =   7
      Top             =   870
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Handle (hWnd):"
      Height          =   255
      Index           =   3
      Left            =   6780
      TabIndex        =   6
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Título:"
      Height          =   255
      Index           =   2
      Left            =   6750
      TabIndex        =   5
      Top             =   630
      Width           =   585
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   1590
      TabIndex        =   4
      Top             =   300
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Ventanas activas:"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   1
      Top             =   300
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
'Varias pruebas con ventanas                                            (01/Ene/99)
'Para enumerar las ventanas visibles y poder cerrarlas.
'
'©Guillermo 'guille' Som, 1999
'----------------------------------------------------------------------------------
Option Explicit

Dim m_UtilVentanas As cWindows

Private Sub cmdCerrarCarpetas_Click()
    'Cierra las carpetas abiertas, el ClassName es: CabinetWClass
    'Nota: el IE4 también tiene ese ClassName
    Dim sTitulo As String
    Dim i As Long
    
    With List1
        For i = 0 To .ListCount - 1
            sTitulo = .List(i)
            Call m_UtilVentanas.CloseApp(sTitulo, "CabinetWClass")
            DoEvents
        Next
    End With
    
    'No se refresca bien, así que seguramente tendrás que pulsar en el botón...
    cmdRefrescar_Click
    
End Sub

Private Sub cmdCerrarVentanas_Click()
    'Cerrar las ventanas seleccionadas del ListBox
    '
    Dim sTitulo As String
    Dim i As Long
    
    With List1
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                sTitulo = .List(i)
                'No cerrar esta aplicación
                If (sTitulo <> App.Title) And (sTitulo <> Caption) Then
                    Call m_UtilVentanas.CloseApp(sTitulo)
                    DoEvents
                End If
            End If
        Next
    End With
    
    'No se refresca bien, así que seguramente tendrás que pulsar en el botón...
    cmdRefrescar_Click
End Sub

Private Sub cmdMinimizeAll_Click()
    '¡Cuidado!
    'Si no se especifica el ClassName se minimizan todas las ventanas,
    'si tienes alguna aplicación de VB, se minimiza una ventana que no es el form
    'y después no se puede restaurar...
    '
    'por eso en este ejemplo uso "CabinetWClass" para minimizar las carpetas
    '
    Call m_UtilVentanas.MinimizeAll("CabinetWClass")
End Sub

Private Sub cmdRefrescar_Click()
    
    Call m_UtilVentanas.EnumTopWindows(List1)
    With List1
        Label1(1) = .ListCount
        If .ListCount Then
            .ListIndex = 0
        End If
    End With
End Sub


Private Sub cmdSalir_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Set m_UtilVentanas = New cWindows
    
    
    If App.PrevInstance Then
        Caption = Caption & " (otra más)"
        App.Title = App.Title & " (otra más)"
    End If
    
    Show
        
    cmdRefrescar_Click
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_UtilVentanas = Nothing
    Set Form1 = Nothing
End Sub


Private Sub List1_Click()
    Dim i As Long
    
    With List1
        i = .ListIndex
        If i > -1 Then
            Label2(0) = .List(i)
            Label2(1) = .ItemData(i)
            Label2(2) = m_UtilVentanas.ClassName(Label2(0))
        Else
            Label2(0) = ""
            Label2(1) = ""
            Label2(2) = ""
        End If
    End With
End Sub


