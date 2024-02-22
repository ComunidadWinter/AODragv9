Attribute VB_Name = "WSKSOCK"
'Argentum Online 0.11.20
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez





'AMEN ALEJO POR ESTE MODULO QUE HAS HECHO :D

Option Explicit

#If UsarQueSocket = 1 Then

'Si la variable esta en TRUE , al iniciar el WsApi se crea
'una ventana LABEL para recibir los mensajes. Al detenerlo,
'se destruye.
'Si es FALSE, los mensajes se envian al form frmMain (o el
'que sea).
#Const WSAPI_CREAR_LABEL = True

Private Const SD_RECEIVE As Long = &H0
Private Const SD_SEND As Long = &H1
Private Const SD_BOTH As Long = &H2


Private Const MAX_TIEMPOIDLE_COLALLENA = 1 'minutos
Private Const MAX_COLASALIDA_COUNT = 800

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)

'====================================================================================
'====================================================================================
'Esto es para agilizar la busqueda del slot a partir de un socket dado,
'sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.

Public Type tSockCache
    Sock As Long
    Slot As Long
End Type

'Public WSAPISockChache() As tSockCache 'Lista de pares SOCKET -> SLOT
'Public WSAPISockChacheCant As Long 'cantidad de elementos para hacer una busqueda eficiente :P
Public WSAPISock2Usr As New Collection

'====================================================================================
'====================================================================================

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long

'====================================================================================
'====================================================================================

Public SockListen As Long

#End If

'====================================================================================
'====================================================================================


Public Sub IniciaWsApi(ByVal hwndParent As Long)
#If UsarQueSocket = 1 Then

Call LogApiSock("IniciaWsApi")
Debug.Print "IniciaWsApi"

#If WSAPI_CREAR_LABEL Then
hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
#Else
hWndMsg = hwndParent
#End If 'WSAPI_CREAR_LABEL

OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

Dim Desc As String
Call StartWinsock(Desc)

#End If
End Sub

Public Sub LimpiaWsApi(ByVal hwnd As Long)
#If UsarQueSocket = 1 Then

Call LogApiSock("LimpiaWsApi")

If WSAStartedUp Then
    Call EndWinsock
End If

If OldWProc <> 0 Then
    SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
    OldWProc = 0
End If

#If WSAPI_CREAR_LABEL Then
If hWndMsg <> 0 Then
    DestroyWindow hWndMsg
End If
#End If

#End If
End Sub

Public Function BuscaSlotSock(ByVal s As Long, Optional ByVal CacheInd As Boolean = False) As Long
Debug.Print "BuscaSockSlot"
#If UsarQueSocket = 1 Then

On Error GoTo hayerror

BuscaSlotSock = WSAPISock2Usr.Item(CStr(s))
Exit Function

hayerror:
BuscaSlotSock = -1

'Dim i As Long
'
'For i = 1 To MaxUsers
'    If UserList(i).ConnID = s And UserList(i).ConnIDValida Then
'        BuscaSlotSock = i
'        Exit Function
'    End If
'Next i
'
'BuscaSlotSock = -1

'
'Dim Pri As Long, Ult As Long, Med As Long
'
'If WSAPISockChacheCant > 0 Then
'    'Busqueda Dicotomica :D
'    Pri = 1
'    Ult = WSAPISockChacheCant
'    Med = Int((Pri + Ult) / 2)
'
'    Do While (Pri <= Ult) And (WSAPISockChache(Med).Sock <> s)
'        If s < WSAPISockChache(Med).Sock Then
'            Ult = Med - 1
'        Else
'            Pri = Med + 1
'        End If
'        Med = Int((Pri + Ult) / 2)
'    Loop
'
'    If Pri <= Ult Then
'        If CacheInd Then
'            BuscaSlotSock = Med
'        Else
'            BuscaSlotSock = WSAPISockChache(Med).Slot
'        End If
'    Else
'        BuscaSlotSock = -1
'    End If
'Else
'    BuscaSlotSock = -1
'End If

#End If
End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal Slot As Long)
Debug.Print "AgregaSockSlot"
#If (UsarQueSocket = 1) Then

If frmMain.SUPERLOG.value = 1 Then LogCustom ("AgregaSlotSock:: sock=" & Sock & " slot=" & Slot)

If WSAPISock2Usr.Count > MaxUsers Then
    If frmMain.SUPERLOG.value = 1 Then LogCustom ("Imposible agregarSlotSock (wsapi2usr.count>maxusers)")
    Call CloseSocket(Slot)
    Exit Sub
End If

WSAPISock2Usr.Add CStr(Slot), CStr(Sock)

'Dim Pri As Long, Ult As Long, Med As Long
'Dim LoopC As Long
'
'If WSAPISockChacheCant > 0 Then
'    Pri = 1
'    Ult = WSAPISockChacheCant
'    Med = Int((Pri + Ult) / 2)
'
'    Do While (Pri <= Ult) And (Ult > 1)
'        If Sock < WSAPISockChache(Med).Sock Then
'            Ult = Med - 1
'        Else
'            Pri = Med + 1
'        End If
'        Med = Int((Pri + Ult) / 2)
'    Loop
'
'    Pri = IIf(Sock < WSAPISockChache(Med).Sock, Med, Med + 1)
'    Ult = WSAPISockChacheCant
'    For LoopC = Ult To Pri Step -1
'        WSAPISockChache(LoopC + 1) = WSAPISockChache(LoopC)
'    Next LoopC
'    Med = Pri
'Else
'    Med = 1
'End If
'WSAPISockChache(Med).Slot = Slot
'WSAPISockChache(Med).Sock = Sock
'WSAPISockChacheCant = WSAPISockChacheCant + 1

#End If
End Sub

Public Sub BorraSlotSock(ByVal Sock As Long, Optional ByVal CacheIndice As Long)
#If (UsarQueSocket = 1) Then
Dim Cant As Long

Cant = WSAPISock2Usr.Count
On Error Resume Next
WSAPISock2Usr.Remove CStr(Sock)

Debug.Print "BorraSockSlot " & Cant & " -> " & WSAPISock2Usr.Count

'Dim N As Long, Indice As Long
'
''If IsMissing(CacheIndice) Then
'    Indice = BuscaSlotSock(Sock, True)
'    If Indice < 1 Then Exit Sub
''Else
''    Indice = CacheIndice
''End If
'
'WSAPISockChacheCant = WSAPISockChacheCant - 1
'
'For N = Indice To WSAPISockChacheCant
'    WSAPISockChache(N) = WSAPISockChache(N + 1)
'Next N

#End If
End Sub



Public Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#If UsarQueSocket = 1 Then

On Error Resume Next

Dim Ret As Long
Dim Tmp As String

Dim s As Long, E As Long
Dim n As Integer
    
Dim Dale As Boolean
Dim UltError As Long


WndProc = 0


'If CamaraLenta = 1 Then
   ' Sleep 1
'End If


Select Case msg
Case 1025

    s = wParam
    E = WSAGetSelectEvent(lParam)
    'Debug.Print "Msg: " & msg & " W: " & wParam & " L: " & lParam
    Call LogApiSock("Msg: " & msg & " W: " & wParam & " L: " & lParam)
    
    Select Case E
    Case FD_ACCEPT
            If frmMain.SUPERLOG.value = 1 Then LogCustom ("FD_ACCEPT")
        If s = SockListen Then
            If frmMain.SUPERLOG.value = 1 Then LogCustom ("sockLIsten = " & s & ". Llamo a Eventosocketaccept")
            Call EventoSockAccept(s)
        End If
        
'    Case FD_WRITE
'        N = BuscaSlotSock(s)
'        If N < 0 And s <> SockListen Then
'            'Call apiclosesocket(s)
'            call WSApiCloseSocket(s)
'            Exit Function
'        End If
'
'        UserList(N).SockPuedoEnviar = True

'        Call IntentarEnviarDatosEncolados(N)
'
''        Dale = UserList(N).ColaSalida.Count > 0
''        Do While Dale
''            Ret = WsApiEnviar(N, UserList(N).ColaSalida.Item(1), False)
''            If Ret <> 0 Then
''                If Ret = WSAEWOULDBLOCK Then
''                    Dale = False
''                Else
''                    'y aca que hacemo' ?? help! i need somebody, help!
''                    Dale = False
''                    Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
''                End If
''            Else
''            '    Debug.Print "Dato de la cola enviado"
''                UserList(N).ColaSalida.Remove 1
''                Dale = (UserList(N).ColaSalida.Count > 0)
''            End If
''        Loop

    Case FD_READ
        
        n = BuscaSlotSock(s)
        If n < 0 And s <> SockListen Then
            'Call apiclosesocket(s)
            Call WSApiCloseSocket(s)
            Exit Function
        End If
        
        'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (0))
        
        '4k de buffer
        Tmp = Space(8192)   'si cambias este valor, tambien hacelo mas abajo
                            'donde dice ret = 8192 :)
        
        Ret = recv(s, Tmp, Len(Tmp), 0)
        ' Comparo por = 0 ya que esto es cuando se cierra
        ' "gracefully". (mas abajo)
        If Ret < 0 Then
            UltError = Err.LastDllError
            If UltError = WSAEMSGSIZE Then
                Debug.Print "WSAEMSGSIZE"
                Ret = 8192
            Else
                Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
                Call LogApiSock("Error en Recv: N=" & n & " S=" & s & " Str=" & GetWSAErrorString(UltError))
                
                'no hay q llamar a CloseSocket() directamente,
                'ya q pueden abusar de algun error para
                'desconectarse sin los 10segs. CREEME.
            '    Call C l o s e Socket(N)
            
                Call CloseSocketSL(n)
                Call Cerrar_Usuario(n)
                Exit Function
            End If
        ElseIf Ret = 0 Then
            Call CloseSocketSL(n)
            Call Cerrar_Usuario(n)
        End If
        
        'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT))
        
        Tmp = Left(Tmp, Ret)
        
        'Call LogApiSock("WndProc:FD_READ:N=" & N & ":TMP=" & Tmp)
        
        Call EventoSockRead(n, Tmp)
        
    Case FD_CLOSE
        n = BuscaSlotSock(s)
        If s <> SockListen Then Call apiclosesocket(s)
        
        Call LogApiSock("WndProc:FD_CLOSE:N=" & n & ":Err=" & WSAGetAsyncError(lParam))
        
        If n > 0 Then
            Call BorraSlotSock(UserList(n).ConnID)
            UserList(n).ConnID = -1
            UserList(n).ConnIDValida = False
            Call EventoSockClose(n)
        End If
        
    End Select
Case Else
    WndProc = CallWindowProc(OldWProc, hwnd, msg, wParam, lParam)
End Select

#End If
End Function

'Retorna 0 cuando se envió o se metio en la cola,
'retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
Public Function WsApiEnviar(ByVal Slot As Integer, ByVal str As String, Optional Encolar As Boolean = True) As Long
#If UsarQueSocket = 1 Then

If frmMain.SUPERLOG.value = 1 Then LogCustom ("WsApiEnviar:: slot=" & Slot & " str=" & str & " len(str)=" & Len(str) & " encolar=" & Encolar)

Dim Ret As String
Dim UltError As Long
Dim Retorno As Long

Retorno = 0

'Debug.Print ">>>> " & str

If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
'    If  Then
'        ' SI hay elementos sin enviar, lo mete en la cola
'        ' ya q hay q mantener un orden de paquetes
'        UserList(Slot).ColaSalida.Add str 'Metelo en la cola Vite'
'    If (UserList(Slot).SockPuedoEnviar And (UserList(Slot).ColaSalida.Count = 0)) Or (Not Encolar) Then
    If ((UserList(Slot).ColaSalida.Count = 0)) Or (Not Encolar) Then
        If frmMain.SUPERLOG.value = 1 Then LogCustom ("WsApiEnviar:: Previo a ret = send(" & UserList(Slot).ConnID & "," & str & "," & Len(str) & ",0)")
        Ret = Send(ByVal UserList(Slot).ConnID, ByVal str, ByVal Len(str), ByVal 0)
        If frmMain.SUPERLOG.value = 1 Then LogCustom ("WsApiEnviar:: Post a ret = send(" & UserList(Slot).ConnID & "," & str & "," & Len(str) & ",0) devolvio: " & Ret)
        If Ret < 0 Then
            UltError = Err.LastDllError
            If frmMain.SUPERLOG.value = 1 Then LogCustom ("WsApiEnviar:: if ret<0 then.. ulterror=" & UltError)
        '    Debug.Print "Error en Send " & Ret & " " & UltError & " " & GetWSAErrorString(Err.LastDllError)
            If UltError = WSAEWOULDBLOCK Then
                UserList(Slot).SockPuedoEnviar = False
                If frmMain.SUPERLOG.value = 1 Then LogCustom ("WsApiEnviar:: seteo UL(" & Slot & ").SockPuedOenviar=false")
                If Encolar Then
                    UserList(Slot).ColaSalida.Add str 'Metelo en la cola Vite'
                    If frmMain.SUPERLOG.value = 1 Then LogCustom ("WsApiEnviar:: encolo en UL(" & Slot & ")")
        '            Debug.Print "Dato encolado."
'                Else
'                    Retorno = UltError
                End If
'            Else
'                Retorno = Ret
            End If
            Retorno = UltError
            'LogApiSock ("Error en Send, slot: " & Slot)
            'Call CloseSocket(Slot)
        End If
    Else
        If UserList(Slot).ColaSalida.Count < MAX_COLASALIDA_COUNT Or UserList(Slot).Counters.IdleCount < MAX_TIEMPOIDLE_COLALLENA Then
            UserList(Slot).ColaSalida.Add str 'Metelo en la cola Vite'
        Else
            Retorno = -1
        End If
    End If
ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then
    'If Not UserList(Slot).Counters.Saliendo Then
        Retorno = -1
    'End If
End If

WsApiEnviar = Retorno

#End If
End Function


Public Sub LogCustom(ByVal str As String)
#If (UsarQueSocket = 1) Then

On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\custom.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

Exit Sub

errhandler:

#End If
End Sub


Public Sub LogApiSock(ByVal str As String)
#If (UsarQueSocket = 1) Then

On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

Exit Sub

errhandler:

#End If
End Sub


Public Sub IntentarEnviarDatosEncolados(ByVal n As Integer)
#If UsarQueSocket = 1 Then

Dim Dale As Boolean
Dim Ret As Long

Dale = UserList(n).ColaSalida.Count > 0
Do While Dale
    Ret = WsApiEnviar(n, UserList(n).ColaSalida.Item(1), False)
    If Ret <> 0 Then
        If Ret = WSAEWOULDBLOCK Then
            Dale = False
        Else
            'y aca que hacemo' ?? help! i need somebody, help!
            Dale = False
            Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
            Call LogApiSock("IntentarEnviarDatosEncolados: N=" & n & " " & GetWSAErrorString(Ret))
            Call CloseSocketSL(n)
            Call Cerrar_Usuario(n)
        End If
    Else
    '    Debug.Print "Dato de la cola enviado"
        UserList(n).ColaSalida.Remove 1
        Dale = (UserList(n).ColaSalida.Count > 0)
    End If
Loop

#End If
End Sub


Public Sub EventoSockAccept(ByVal SockID As Long)
#If UsarQueSocket = 1 Then
'==========================================================
'USO DE LA API DE WINSOCK
'========================

'Call LogApiSock("EventoSockAccept")

If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Pedido de conexion SocketID:" & SockID & vbCrLf

'on error Resume Next
    
    Dim NewIndex As Integer
    Dim Ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
    Dim tStr As String
    
    Tam = sockaddr_size
    
    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "NextOpenUser" & vbCrLf
    
    NewIndex = NextOpenUser ' Nuevo indice
    If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventSockAccept:: Newindex=" & NewIndex)
    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "UserIndex asignado " & NewIndex & vbCrLf
    
'    Debug.Print NewIndex

    '=============================================
    'SockID es en este caso es el socket de escucha,
    'a diferencia de socketwrench que es el nuevo
    'socket de la nueva conn
    
    If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventSockAccept:: Llamo a accept(" & SockID & ",sa," & Tam & ")")
    
    Ret = Accept(SockID, sa, Tam)
    
    If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventSockAccept:: accept devuelve ret=" & Ret)
    
    If Ret = INVALID_SOCKET Then
        i = Err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        If frmMain.SUPERLOG.value = 1 Then Call LogCustom("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        Exit Sub
    End If
    NuevoSock = Ret

    If NewIndex <= MaxUsers Then
        If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Cargando Socket " & NewIndex & vbCrLf
        
               
        UserList(NewIndex).ip = GetAscIP(sa.sin_addr)
        
        If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventSockAccept:: GetAscIP=" & UserList(NewIndex).ip)
        
        'Busca si esta banneada la ip
       ' For i = 1 To BanIps.Count
            'If BanIps.Item(i) = UserList(NewIndex).ip Then
                'Call apiclosesocket(NuevoSock)
                'Call WSApiCloseSocket(NuevoSock)
                'Exit Sub
            'End If
       ' Next i
        
        
        Call LogApiSock("EventoSockAccept NewIndex: " & NewIndex & " NuevoSock: " & NuevoSock & " IP: " & UserList(NewIndex).ip)
        '=============================================
        If aDos.MaxConexiones(UserList(NewIndex).ip) Then
            UserList(NewIndex).ConnID = -1
            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "User slot reseteado " & NewIndex & vbCrLf
            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Socket unloaded" & NewIndex & vbCrLf
            'Call LogCriticEvent(UserList(NewIndex).ip & " intento crear mas de 3 conexiones.")
            Call aDos.RestarConexion(UserList(NewIndex).ip)
            'Call apiclosesocket(NuevoSock)
            Call WSApiCloseSocket(NuevoSock)
            'Exit Sub
        End If
        
        If NewIndex > LastUser Then LastUser = NewIndex
        
        UserList(NewIndex).SockPuedoEnviar = True
        UserList(NewIndex).ConnID = NuevoSock
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        Set UserList(NewIndex).ColaSalida = New Collection
        
        If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventoSockAccept:: Voy a llamar a agregaSlotSock(" & NuevoSock & "," & NewIndex & ")")
        Call AgregaSlotSock(NuevoSock, NewIndex)
        
'        Debug.Print "Conexion desde " & UserList(NewIndex).ip
        
        If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & UserList(NewIndex).ip & " logged." & vbCrLf & vbCrLf
    Else
        Call LogCriticEvent("No acepte conexion porque no tenia slots")
        If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventoSockAccept:: No tenia mas slots! sockid=" & SockID & " nuevosock=" & NuevoSock & " newindex=" & NewIndex)
        tStr = "ERRServer lleno." & ENDC
        Dim AAA As Long
        AAA = Send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
        If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventoSockAccept:: No tenia mas slots! send devuelve=" & AAA)
'        Ret = accept(SockID, sa, Tam)
'        If Ret = INVALID_SOCKET Then
'            Call LogCriticEvent("Error en Accept() API")
'            Exit Sub
'        End If

        'Call apiclosesocket(NuevoSock)
        Call WSApiCloseSocket(NuevoSock)
    End If
    
#End If
End Sub

Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos As String)
#If UsarQueSocket = 1 Then

Dim t() As String
Dim loopc As Long

Debug.Print "<<<< " & Datos

If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "EventoSockRead UI: " & Slot & " Datos: " & Datos & vbCrLf

'TCPESStats.BytesRecibidos = TCPESStats.BytesRecibidos + Len(Datos)

If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventoSockRead:: slot=" & Slot & " datos=" & Datos)

UserList(Slot).RDBuffer = UserList(Slot).RDBuffer & Datos

'If InStr(1, UserList(Slot).RDBuffer, Chr(2)) > 0 Then
'    UserList(Slot).RDBuffer = "CLIENTEVIEJO" & ENDC
'    Debug.Print "CLIENTEVIEJO"
'End If

t = Split(UserList(Slot).RDBuffer, ENDC)
If UBound(t) > 0 Then
    UserList(Slot).RDBuffer = t(UBound(t))
    
    For loopc = 0 To UBound(t) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            'If t(LoopC) <> "" Then If Not UserList(Slot).CommandsBuffer.Push(t(LoopC)) Then Call Cerrar_Usuario(Slot, 0)
            If t(loopc) <> "" Then If Not UserList(Slot).CommandsBuffer.Push(t(loopc)) Then Call CloseSocket(Slot)
            If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventoSockAccept:: Pude pushear los datos del slot " & Slot)
        Else ' no encolamos los comandos (MUY VIEJO)
              If UserList(Slot).ConnID <> -1 Then
                Call HandleData(Slot, t(loopc))
              Else
                Exit Sub
              End If
        End If
    Next loopc
End If

#End If
End Sub

Public Sub EventoSockClose(ByVal Slot As Integer)
#If UsarQueSocket = 1 Then

    If frmMain.SUPERLOG.value = 1 Then LogCustom ("EventoSockClose:: slot=" & Slot)
    If UserList(Slot).flags.UserLogged Then
        Call CloseSocketSL(Slot)
        Call Cerrar_Usuario(Slot)
    Else
        Call CloseSocket(Slot)
    End If
#End If
End Sub


Public Sub WSApiReiniciarSockets()
#If UsarQueSocket = 1 Then
Dim i As Long
    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Cierra todas las conexiones
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
            Call CloseSocket(i)
        End If
        
        'Call ResetUserSlot(i)
    Next i
    
    ' No 'ta el PRESERVE :p
    ReDim UserList(1 To MaxUsers)
    For i = 1 To MaxUsers
        UserList(i).ConnID = -1
        UserList(i).ConnIDValida = False
    Next i
    
    LastUser = 1
    NumUsers = 0
    
    Call LimpiaWsApi(frmMain.hwnd)
    Call Sleep(100)
    Call IniciaWsApi(frmMain.hwnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

'    'Inicia el socket de escucha
'    SockListen = ListenForConnect(Puerto, hWndMsg, "")
'
'    'Comprueba si el proc de la ventana es el correcto
'    Dim TmpWProc As Long
'    TmpWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)
'    If TmpWProc <> ActualWProc Then
'        MsgBox "Incorrecto proc de ventana (" & TmpWProc & " <> " & ActualWProc & ")"
'        Call LogApiSock("INCORRECTO PROC DE VENTANA")
'        OldWProc = TmpWProc
'        If OldWProc <> 0 Then
'            SetWindowLong frmMain.hWnd, GWL_WNDPROC, AddressOf WndProc
'            ActualWProc = GetWindowLong(frmMain.hWnd, GWL_WNDPROC)
'        End If
'    End If
#End If
End Sub


Public Sub WSApiCloseSocket(ByVal Socket As Long)
#If UsarQueSocket = 1 Then
Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
Call ShutDown(Socket, SD_BOTH)
#End If
End Sub

