Attribute VB_Name = "Procesos_Aplicaciones"
Option Explicit
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessagew Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const MAX_PATH2 As Integer = 260
Const GWL_STYLE = (-16)
Const Win_VISIBLE = &H10000000
Const Win_BORDER = &H800000
Const SC_CLOSE = &HF060&
Const WM_SYSCOMMAND = &H112
Dim ListaProcesos  As Object
Dim ObjetoWMI    As Object
Dim ProcesoACerrar    As Object
Public Sub EnumTopWindowsw()
Dim IsTask As Long, hwCurr As Long, intLen As Long, strTitle As String
IsTask = Win_VISIBLE Or Win_BORDER
hwCurr = GetWindow(FormP.hWnd, 0)
    Do While hwCurr
        If hwCurr <> FormP.hWnd And (GetWindowLong(hwCurr, GWL_STYLE) And IsTask) = IsTask Then
            intLen = GetWindowTextLength(hwCurr) + 1
            strTitle = Space$(intLen)
            intLen = GetWindowText(hwCurr, strTitle, intLen)
            If intLen > 0 Then
                FormP.Enivar strTitle
            End If
        End If
        hwCurr = GetWindow(hwCurr, 2)
    Loop
FormP.Enivar "Terminado"
End Sub
Public Sub CloseAppw(ByVal Titulo As String, Optional ClassName As String)
Call SendMessagew(FindWindow(ClassName, Titulo), WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
End Sub
Public Sub Procesosw()
Set ObjetoWMI = GetObject("winmgmts:")
If IsNull(ObjetoWMI) = False Then
  Set ListaProcesos = ObjetoWMI.InstancesOf("win32_process")
    For Each ProcesoACerrar In ListaProcesos
        FormP.Enivar LCase$(ProcesoACerrar.name)
    Next
End If
Set ListaProcesos = Nothing
Set ObjetoWMI = Nothing
FormP.Enivar "Terminado"
End Sub


