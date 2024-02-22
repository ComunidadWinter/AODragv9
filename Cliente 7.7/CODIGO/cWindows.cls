VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
'Clase para manipular ventanas de Windows                               (01/Ene/99)
'
'Esta clase enumera las ventas visibles, cierra la indicada, etc.
'
'©Guillermo 'guille' Som, 1999
'----------------------------------------------------------------------------------
Option Explicit

Private Declare Function IsWindowVisible Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
    (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

' GetWindow() Constants
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&

Private Declare Function GetWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal wFlag As Long) As Long

'
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const SC_MINIMIZE = &HF020&
Private Const SC_CLOSE = &HF060&
Private Const WM_SYSCOMMAND = &H112
Private Const WM_CLOSE = &H10

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Sub CloseApp(ByVal Titulo As String, Optional ClassName As String)
    'Cerrar la ventana indicada, mediante el menú del sistema (o de windows)
    'Esto funcionará si la aplicación tiene menú de sistema
    '(aunque lo he probado con una utilidad sin controlBox y la cierra bien)
    '
    'Si se especifica ClassName, se cerrarán la ventana si es de ese ClassName
    '
    Dim hWnd As Long
    
    'No cerrar la ventana "Progman"
    If Titulo <> "Progman" Then
        hWnd = FindWindow(ClassName, Titulo)
        
        Call SendMessage(hWnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
    End If
End Sub

Private Function WindowTitle(ByVal hWnd As Long) As String
    'Devuelve el título de una ventana, según el hWnd indicado
    '
    Dim sTitulo As String
    Dim lenTitulo As Long
    Dim ret As Long
    
    'Leer la longitud del título de la ventana
    lenTitulo = GetWindowTextLength(hWnd)
    If lenTitulo > 0 Then
        lenTitulo = lenTitulo + 1
        sTitulo = String$(lenTitulo, 0)
        'Leer el título de la ventana
        ret = GetWindowText(hWnd, sTitulo, lenTitulo)
        WindowTitle = Left$(sTitulo, ret)
    End If
End Function
Public Function EnumTopWindows(Optional unListBox As Control) As Variant
    'Enumera las ventanas que tienen título y son visibles
    'Devuelve un array del tipo Variant con los nombres de las ventanas
    'y su hWnd
    'Por tanto la forma de acceder a este array sería:
    '   Set col = EnumTopWindows
    '   numItems = col.Count
    '   For i = 1 To numItems Step 2
    '       With List2
    '           .AddItem col.Item(i)
    '           .ItemData(.NewIndex) = col.Item(i + 1)
    '       End With
    '   Next
    '
    'Opcionalemente se puede especificar como parámetro un ListBox o ComboBox
    'y los datos se añadirán a ese control
    '
    Dim sTitulo As String
    Dim hWnd As Long
    Dim col As Collection
    
    Set col = New Collection
    
    If Not unListBox Is Nothing Then
        unListBox.Clear
    End If
    
    'Primera ventana
    hWnd = GetWindow(GetDesktopWindow(), GW_CHILD)
    
    'Recorrer el resto de las ventanas
    Do While hWnd <> 0&
        'Si la ventana es visible
        If IsWindowVisible(hWnd) Then
            'Leer el caption de la ventana
            sTitulo = WindowTitle(hWnd)
            If Len(sTitulo) Then
                'Añadimos el título
                col.Add sTitulo
                'y el hWnd por si fuese útil
                col.Add hWnd
                'Si se especifica el ListBox
                If Not unListBox Is Nothing Then
                    With unListBox
                        .AddItem sTitulo
                        .ItemData(.NewIndex) = hWnd
                    End With
                End If
            End If
        End If
        'Siguiente ventana
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop
    Set EnumTopWindows = col
End Function


Public Function ClassName(ByVal Title As String) As String
    'Devuelve el ClassName de una ventana, indicando el título de la misma
    Dim hWnd As Long
    Dim sClassName As String
    Dim nMaxCount As Long
    
    hWnd = FindWindow(sClassName, Title)
    
    nMaxCount = 256
    sClassName = Space$(nMaxCount)
    nMaxCount = GetClassName(hWnd, sClassName, nMaxCount)
    ClassName = Left$(sClassName, nMaxCount)
End Function

Public Sub MinimizeAll(Optional ClassName As String)
    'Minimizar todas las ventanas
    '
    Dim col As Collection
    Dim numItems As Long
    Dim i As Long
    Dim sTitulo As String
    Dim hWnd As Long
    
    Set col = New Collection
    
    Set col = Me.EnumTopWindows
    numItems = col.Count
    For i = 1 To numItems Step 2
        sTitulo = col.Item(i)
        hWnd = FindWindow(ClassName, sTitulo)
        'hWnd = col.Item(i + 1)
        Call SendMessage(hWnd, WM_SYSCOMMAND, SC_MINIMIZE, ByVal 0&)
    Next
    
    Set col = Nothing
End Sub
