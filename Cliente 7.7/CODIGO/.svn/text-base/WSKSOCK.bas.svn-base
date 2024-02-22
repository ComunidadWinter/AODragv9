Attribute VB_Name = "WSKSOCK"
Option Explicit

Public Const WINSOCK_MESSAGE As Long = 1025
Public Const INADDR_NONE = &HFFFF
Public Const INADDR_ANY = &H0

Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Public Const sockaddr_size = 16
Public saZero As sockaddr
Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const AF_INET = 2
Public Const PF_INET = 2

'---Windows System Functions
    Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
    Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'---async notification constants
    Public Const SOL_SOCKET = &HFFFF&
    Public Const SO_LINGER = &H80&
    Public Const FD_READ = &H1&
    Public Const FD_WRITE = &H2&
    Public Const FD_OOB = &H4&
    Public Const FD_ACCEPT = &H8&
    Public Const FD_CONNECT = &H10&
    Public Const FD_CLOSE = &H20&
'---SOCKET FUNCTIONS
    Public Declare Function accept Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, addrlen As Long) As Long
    Public Declare Function bind Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
    Public Declare Function Connect Lib "wsock32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Public Declare Function recv Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
    Public Declare Function Send Lib "wsock32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
    Public Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'---DATABASE FUNCTIONS
    Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
'---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
    Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
    Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
    Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

    Public WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled


Sub EndWinsock()
    Dim Ret&
    If WSAIsBlocking() Then
        Ret = WSACancelBlockingCall()
    End If
    Ret = WSACleanup()
    WSAStartedUp = False
End Sub


Function getascip(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr&

    Dim nStr&

    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        getascip = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    getascip = retString
    If Err Then getascip = "255.255.255.255"
End Function



Public Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&) As Long
    Dim s&, Dummy&
    Dim SelectOps&

    Dim sockin As sockaddr
    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_SOCKET Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    s = socket(PF_INET, SOCK_STREAM, 0)
    If s < 0 Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    If bind(s, sockin, sockaddr_size) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, HWndToMsg, ByVal WINSOCK_MESSAGE, ByVal SelectOps) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ListenForConnect = SOCKET_ERROR
        Exit Function
    End If
    
    If listen(s, 1) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    ListenForConnect = s
End Function

Public Function SendDataw(ByVal s&, vMessage As Variant) As Long

    Dim TheMsg() As Byte, sTemp$
    TheMsg = ""
    Select Case VarType(vMessage)
        Case 8209   'byte array
            sTemp = vMessage
            TheMsg = sTemp
        Case 8      'string, if we recieve a string, its assumed we are linemode

            sTemp = StrConv(vMessage, vbFromUnicode)

        Case Else
            sTemp = CStr(vMessage)

            sTemp = StrConv(vMessage, vbFromUnicode)

    End Select
    TheMsg = sTemp
    If UBound(TheMsg) > -1 Then
        SendDataw = Send(s, TheMsg(0), (UBound(TheMsg) - LBound(TheMsg) + 1), 0)
    End If
End Function
Public Function StartWinsock(sDescription As String) As Boolean
    Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function






