Attribute VB_Name = "ModSeguridad"
'**********************************Anti Debugger********************************
Private Declare Function IsDebuggerPresent Lib "kernel32" () As Long
'*******************************************************************************

'********************************Anti Speed Hack********************************
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Time As Long
Private count As Integer
'*******************************************************************************

'*******************Detecta si se cambio el nombre al exe***********************
Public OriginalClientName As String
Public ClientName As String
Public DetectName As String
'*******************************************************************************

'**************Detecta externos mediante nombre de ventanas*********************
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public CantidadCheats As Byte
Private NameCheat(1 To 255) As String
'*******************************************************************************


'*****************************
'ANTICHEAT ENGINE
'*****************************

Public Sub BuscarEngine()
    On Error Resume Next
    Dim X As String
    Dim MiObjeto As Object
    
    Set MiObjeto = CreateObject("Wscript.Shell")
    X = "1"
    X = MiObjeto.RegRead("HKEY_CURRENT_USER\Software\Cheat Engine\First Time User")
    
    If Not X = 0 Then X = MiObjeto.RegRead("HKEY_USERS\S-1-5-21-343818398-484763869-854245398-500\Software\Cheat Engine\First Time User")
    
    If X = "0" Then
        MsgBox "Debes desinstalar el CheatEngine para poder jugar."
        Call CloseClient
    End If
End Sub
'****************************

'***************************Anti Debugger***************************
Public Function Debugger() As Boolean
    If IsDebuggerPresent Then
        Debugger = True
        Exit Function
    End If
    Debugger = False
End Function

Public Sub AntiDebugger()
    MsgBox "Se ha detectado un intento de Debuggear el cliente, su cliente será cerrado.!", vbCritical, "AODrag"
    Call CloseClient
End Sub
'*******************************************************************

'************************Anti Speed Hack****************************
Public Sub AntiShInitialize()
Time = GetTickCount()
End Sub

Public Function AntiSh() As Boolean
If GetTickCount - Time > 10100 Or GetTickCount - Time < 9500 Then
        count = count + 1
        Call LogError("Count: " & count & " Get: " & GetTickCount)
    Else
        count = 0
    End If
    
   If count > 30 Then
       AntiSh = True
       Call LogError("Detectado SH - Get: " & GetTickCount)
       Exit Function
    End If

Time = GetTickCount()
AntiSh = False
End Function

Public Sub AntiShOn()
    MsgBox "Se ha detectado el uso de SpeedHack, el cliente será cerrado!.", vbCritical, "AODrag"
    Call CloseClient
End Sub
'*******************************************************************

'***********Detecta si se le cambio el nombre al exe***************
Public Function ChangeName() As Boolean
    If OriginalClientName <> ClientName Then
        ChangeName = True
        Exit Function
    End If
    ChangeName = False
End Function
'*******************************************************************

'**************Detecta externos mediante nombre de ventanas*********************
Public Sub BuscarCheats()
'Lorwik> La verdad esque no me gusta mucho, pero mejor esto que nada xD
Dim i As Byte

    For i = 1 To CantidadCheats
        If FindWindow(vbNullString, NameCheat(i)) Then
            Call HayExterno
        End If
    Next i
End Sub

Public Function HayExterno()
    Call MsgBox("Se ha detectado una aplicacion externa prohibida , seras expulsado del juego. Lee el reglamento si tienes dudas.")
    Call CloseClient
End Function

Public Sub CargarNombreCheats()
'Lorwik: Es una chapuza, pero mejor que nada...
    NameCheat(1) = "Cambia titulos"
    NameCheat(2) = "Fedex Macro"
    NameCheat(3) = "Loopzer Cheat"
    NameCheat(4) = "Cheat Engine"
    NameCheat(5) = "Serbio Engine"
    NameCheat(6) = "WPE Pro"
    NameCheat(7) = " PermEdit"
    NameCheat(8) = " AutoHotKey"
    NameCheat(9) = " MoonlightEngine"
    NameCheat(10) = "KmeT Engine"
    NameCheat(11) = "X-Z Engine"
    NameCheat(12) = "Macro CDerecho"
    NameCheat(13) = " SpeedGear"
    NameCheat(14) = " MacroSaraza"
    NameCheat(15) = " SpeedHackNT"
    NameCheat(16) = " SandBoxie"
    NameCheat(17) = " Engine"
    NameCheat(18) = "All Keys"
    NameCheat(19) = "Key Extender"
    NameCheat(20) = "Keystroke Converter"
    NameCheat(21) = "Mouse Recorder Pro2"
    NameCheat(22) = "Mouse Recorder Premium"
    NameCheat(23) = "Mouse Recorder Pro"
    NameCheat(24) = "Mouse Recorder Pro 2"
    NameCheat(25) = "Makro Tuky"
    NameCheat(26) = "Piringulete 2003"
    NameCheat(27) = "El Chit del Geri"
    NameCheat(28) = "Macro K33"
    NameCheat(29) = "Makro K33"
    NameCheat(3) = " Fede"
End Sub
'******************************************************************
