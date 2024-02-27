Attribute VB_Name = "wskapiAO"
'**************************************************************
' wskapiAO.bas
'
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

''
' Modulo para manejar Winsock
'

#If SocketType = 1 Then


'Si la variable esta en TRUE , al iniciar el WsApi se crea
'una ventana LABEL para recibir los mensajes. Al detenerlo,
'se destruye.
'Si es FALSE, los mensajes se envian al form frmMain (o el
'que sea).
#Const WSAPI_CREAR_LABEL = True

Private Const SD_BOTH As Long = &H2

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)

Private Const SIZE_RCVBUF As Long = 8192
Private Const SIZE_SNDBUF As Long = 8192

''
'Esto es para agilizar la busqueda del slot a partir de un socket dado,
'sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.
'
' @param Sock sock
' @param slot slot
'
Public Type tSockCache
    Sock As Long
    Slot As Long
End Type

Public WSAPISock2Usr As Collection

' ====================================================================================
' ====================================================================================

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long

' ====================================================================================
' ====================================================================================

Public SockListen As Long

#Else
    Global Const WSAEWOULDBLOCK = 10035
#End If

Public LastSockListen As Long ' GSZAO

' ====================================================================================
' ====================================================================================


Public Sub IniciaWsApi(ByVal hwndParent As Long)
#If SocketType = 1 Then

Call LogApiSock("IniciaWsApi")
Debug.Print "IniciaWsApi"

#If WSAPI_CREAR_LABEL Then
hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
#Else
hWndMsg = hwndParent
#End If 'WSAPI_CREAR_LABEL

OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

Dim desc As String
Call StartWinsock(desc)

#End If
End Sub

Public Sub LimpiaWsApi()
#If SocketType = 1 Then

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

Public Function BuscaSlotSock(ByVal S As Long) As Long
#If SocketType = 1 Then
On Error GoTo hayerror
    
    If WSAPISock2Usr.Count <> 0 Then ' GSZAO
        BuscaSlotSock = WSAPISock2Usr.Item(CStr(S))
    Else
        BuscaSlotSock = -1
    End If
    
Exit Function
    
hayerror:
    BuscaSlotSock = -1
#End If

End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal Slot As Long)
Debug.Print "AgregaSockSlot"
#If (SocketType = 1) Then

    If WSAPISock2Usr.Count > MaxUsers Then
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
    
Public Sub BorraSlotSock(ByVal Sock As Long)
    #If (SocketType = 1) Then
        Dim cant As Long
        
        cant = WSAPISock2Usr.Count
        On Error Resume Next
        WSAPISock2Usr.Remove CStr(Sock)
        
        Debug.Print "BorraSockSlot " & cant & " -> " & WSAPISock2Usr.Count
    #End If
End Sub



Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#If SocketType = 1 Then

On Error Resume Next

    Dim ret As Long
    Dim Tmp() As Byte
    Dim S As Long
    Dim e As Long
    Dim n As Integer
    Dim UltError As Long
    
    Select Case msg
        Case 1025
            S = wParam
            e = WSAGetSelectEvent(lParam)
            
            Select Case e
                Case FD_ACCEPT
                    If S = SockListen Then
                        Call EventoSockAccept(S)
                    End If
                
            '    Case FD_WRITE
            '        N = BuscaSlotSock(s)
            '        If N < 0 And s <> SockListen Then
            '            'Call apiclosesocket(s)
            '            call WSApiCloseSocket(s)
            '            Exit Function
            '        End If
            '
            
            '        Call IntentarEnviarDatosEncolados(N)
            '
            '        Dale = UserList(N).ColaSalida.Count > 0
            '        Do While Dale
            '            Ret = WsApiEnviar(N, UserList(N).ColaSalida.Item(1), False)
            '            If Ret <> 0 Then
            '                If Ret = WSAEWOULDBLOCK Then
            '                    Dale = False
            '                Else
            '                    'y aca que hacemo' ?? help! i need somebody, help!
            '                    Dale = False
            '                    Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
            '                End If
            '            Else
            '            '    Debug.Print "Dato de la cola enviado"
            '                UserList(N).ColaSalida.Remove 1
            '                Dale = (UserList(N).ColaSalida.Count > 0)
            '            End If
            '        Loop
        
                Case FD_READ
                    n = BuscaSlotSock(S)
                    If n < 0 And S <> SockListen Then
                        'Call apiclosesocket(s)
                        Call WSApiCloseSocket(S)
                        Exit Function
                    End If
                    
                    'create appropiate sized buffer
                    ReDim Preserve Tmp(SIZE_RCVBUF - 1) As Byte
                    
                    ret = recv(S, Tmp(0), SIZE_RCVBUF, 0)
                    ' Comparo por = 0 ya que esto es cuando se cierra
                    ' "gracefully". (mas abajo)
                    If ret < 0 Then
                        UltError = Err.LastDllError
                        If UltError = WSAEMSGSIZE Then
                            Debug.Print "WSAEMSGSIZE"
                            ret = SIZE_RCVBUF
                        Else
                            Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
                            Call LogApiSock("Error en Recv: N=" & n & " S=" & S & " Str=" & GetWSAErrorString(UltError))
                            
                            'no hay q llamar a CloseSocket() directamente,
                            'ya q pueden abusar de algun error para
                            'desconectarse sin los 10segs. CREEME.
                            Call CloseSocketSL(n)
                            Call Cerrar_Usuario(n)
                            Exit Function
                        End If
                    ElseIf ret = 0 Then
                        Call CloseSocketSL(n)
                        Call Cerrar_Usuario(n)
                    End If
                    
                    ReDim Preserve Tmp(ret - 1) As Byte
                    
                    Call EventoSockRead(n, Tmp)
                
                Case FD_CLOSE
                    n = BuscaSlotSock(S)
                    If S <> SockListen Then Call apiclosesocket(S)
                    
                    If n > 0 Then
                        Call BorraSlotSock(S)
                        UserList(n).ConnID = -1
                        UserList(n).ConnIDValida = False
                        Call EventoSockClose(n)
                    End If
            End Select
        
        Case Else
            WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)
    End Select
#End If
End Function

'Retorna 0 cuando se envió o se metio en la cola,
'retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
Public Function WsApiEnviar(ByVal Slot As Integer, ByRef str As String) As Long
#If SocketType = 1 Or SocketType = 2 Then
    Dim ret As String
    Dim Retorno As Long
    Dim data() As Byte
    
    ReDim Preserve data(Len(str) - 1) As Byte

    data = StrConv(str, vbFromUnicode)
    
    Retorno = 0
    
    If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
        ret = send(ByVal UserList(Slot).ConnID, data(0), ByVal UBound(data()) + 1, ByVal 0)
        If ret < 0 Then
            ret = Err.LastDllError
            If ret = WSAEWOULDBLOCK Then
                               
                ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
                Call UserList(Slot).outgoingData.WriteASCIIStringFixed(str)
            End If
        End If
    ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then
        If Not UserList(Slot).Counters.Saliendo Then
            Retorno = -1
        End If
    End If
    
    WsApiEnviar = Retorno
#End If
End Function

Public Sub LogApiSock(ByVal str As String)
#If (SocketType = 1) Then

On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\Logs\WSApi.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

Exit Sub

Errhandler:

#End If
End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)
'Last Modification: 09/06/2012 - ^[GS]^
#If SocketType = 1 Then
'==========================================================
'USO DE LA API DE WINSOCK
'========================
    
    Dim NewIndex As Integer
    Dim ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
    Dim str As String
    Dim data() As Byte
    
    Tam = sockaddr_size
    
    '=============================================
    'SockID es en este caso es el socket de escucha,
    'a diferencia de socketwrench que es el nuevo
    'socket de la nueva conn
    
'Modificado por Maraxus
    'Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 0)
    ret = accept(SockID, sa, Tam)

    If ret = INVALID_SOCKET Then
        i = Err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        Exit Sub
    End If
    
    'If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
    '    Call WSApiCloseSocket(ret)
    '    Exit Sub
    'End If

    'If Ret = INVALID_SOCKET Then
    '    If Err.LastDllError = 11002 Then
    '        ' We couldn't decide if to accept or reject the connection
    '        'Force reject so we can get it out of the queue
    '        Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 1)
    '        Call LogCriticEvent("Error en WSAAccept() API 11002: No se pudo decidir si aceptar o rechazar la conexión.")
    '    Else
    '        i = Err.LastDllError
    '        Call LogCriticEvent("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
    '        Exit Sub
    '    End If
    'End If

    NuevoSock = ret

    If setsockopt(NuevoSock, SOL_SOCKET, SO_LINGER, 0, 4) <> 0 Then ' 0.13.3
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear Lingers." & i & ": " & GetWSAErrorString(i))
    End If
    
    

    'Seteamos el tamaño del buffer de entrada
    If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))
    End If
    'Seteamos el tamaño del buffer de salida
    If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   BIENVENIDO AL SERVIDOR!!!!!!!!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
    NewIndex = NextOpenUser ' Nuevo indice
    
    Call Socket_NewConnection(NewIndex, GetAscIP(sa.sin_addr), NuevoSock)
#End If
End Sub

Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos() As Byte)
#If SocketType = 1 Or SocketType = 2 Then

With UserList(Slot)
       
    Call .incomingData.WriteBlock(Datos)
    
    If .ConnID <> -1 Then
        Do While Protocol.HandleIncomingData(Slot) = True
        Loop
    Else
        Exit Sub
    End If
End With

#End If
End Sub

Public Sub EventoSockClose(ByVal Slot As Integer)
'Last Modification: 10/08/2011 - ^[GS]^
#If SocketType = 1 Then

    'Es el mismo user al que está revisando el centinela??
    'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
    Dim CentinelaIndex As Byte
    CentinelaIndex = UserList(Slot).flags.CentinelaIndex
        
    If CentinelaIndex <> 0 Then ' 0.13.3
        Call modCentinela.CentinelaUserLogout(CentinelaIndex)
    End If

    If UserList(Slot).flags.UserLogged Then
        Call CloseSocketSL(Slot)
        Call Cerrar_Usuario(Slot)
    Else
        Call CloseSocket(Slot)
    End If
#End If
End Sub


Public Sub WSApiReiniciarSockets()
    Dim i As Long
    
#If SocketType = 1 Then
    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
#ElseIf SocketType = 2 Then
    frmMain.wskListen.Close
#End If
    'Cierra todas las conexiones
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
            Call CloseSocket(i)
        End If
        
        'Call ResetUserSlot(i)
    Next i
    
    For i = 1 To MaxUsers
        Set UserList(i).incomingData = Nothing
        Set UserList(i).outgoingData = Nothing
    Next i
    
    ' No 'ta el PRESERVE :p
    ReDim UserList(1 To MaxUsers)
    For i = 1 To MaxUsers
        UserList(i).ConnID = -1
        UserList(i).ConnIDValida = False
        
        Set UserList(i).incomingData = New clsByteQueue
        Set UserList(i).outgoingData = New clsByteQueue
    Next i
    
    LastUser = 1
    NumUsers = 0
    
#If SocketType = 1 Then
    Call LimpiaWsApi
    Call Sleep(100)
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
#ElseIf SocketType = 2 Then
    frmMain.wskListen.Close
    frmMain.wskListen.LocalPort = Puerto
    frmMain.wskListen.listen
#End If
End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long, Optional ByVal UserIndex As Integer = 0) 'Alta negrada xD sorry by Mateo
    #If SocketType = 1 Then
        Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
        Call ShutDown(Socket, SD_BOTH)
    #ElseIf SocketType = 2 Then
        If UserIndex > 0 Then
            frmMain.wskClient(UserIndex).Close
        End If
    #End If
End Sub

