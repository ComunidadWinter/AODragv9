Attribute VB_Name = "TCP"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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

Option Explicit

Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Public Declare Function send Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Function GetLongIp(ByVal IPS As String) As Long
    GetLongIp = inet_addr(IPS)
End Function

Public Function GetAscIP(ByVal inn As Long) As String
    #If Win32 Then
        Dim nStr&
    #Else
        Dim nStr%
    #End If
    Dim lpStr&
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        MemCopy ByVal retString, ByVal lpStr, nStr
        retString = Left$(retString, nStr)
        GetAscIP = retString
    Else
        GetAscIP = "255.255.255.255"
    End If
End Function

Public Sub Socket_NewConnection(ByVal UserIndex As Integer, ByVal ip As String, ByVal NuevoSock As Long)
    Dim i As Long
    Dim IPLong As Long
    Dim str As String
    Dim data() As Byte
    
    IPLong = GetLongIp(ip)
    
    If Not SecurityIp.IpSecurityAceptarNuevaConexion(IPLong) Then ' 0.13.3
        Call WSApiCloseSocket(NuevoSock, UserIndex)
        Exit Sub
    End If
    
    If SecurityIp.IPSecuritySuperaLimiteConexiones(IPLong) Then ' 0.13.3
        str = Protocol.PrepareMessageErrorMsg("Limite de conexiones para su IP alcanzado.")
        
        ReDim Preserve data(Len(str) - 1) As Byte
        
        data = StrConv(str, vbFromUnicode)
        
        Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
        Call WSApiCloseSocket(NuevoSock, UserIndex)
        Exit Sub
    End If
    
    If UserIndex <= MaxUsers Then
        
        'Make sure both outgoing and incoming data buffers are clean
        Call UserList(UserIndex).incomingData.ReadASCIIStringFixed(UserList(UserIndex).incomingData.length)
        Call UserList(UserIndex).outgoingData.ReadASCIIStringFixed(UserList(UserIndex).outgoingData.length)

        UserList(UserIndex).ip = ip
        UserList(UserIndex).IPLong = IPLong
        
        'Busca si esta banneada la ip
        For i = 1 To BanIPs.count
            If BanIPs.Item(i) = UserList(UserIndex).ip Then
                'Call apiclosesocket(NuevoSock)
                Call WriteErrorMsg(UserIndex, "Su IP se encuentra bloqueada en este servidor.")
                Call FlushBuffer(UserIndex)
                Call SecurityIp.IpRestarConexion(UserList(UserIndex).IPLong)
                Call WSApiCloseSocket(NuevoSock, UserIndex)
                Exit Sub
            End If
        Next i
         
        If UserIndex > LastUser Then LastUser = UserIndex
        
        UserList(UserIndex).ConnIDValida = True
        UserList(UserIndex).ConnID = NuevoSock
        
        Call AgregaSlotSock(NuevoSock, UserIndex)
    Else
        str = Protocol.PrepareMessageErrorMsg("El servidor se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
        
        ReDim Preserve data(Len(str) - 1) As Byte
        
        data = StrConv(str, vbFromUnicode)
        
        Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
        Call WSApiCloseSocket(NuevoSock, UserIndex)
    End If
End Sub

Sub DarCuerpo(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 14/03/2007
'Elije una cabeza para el usuario y le da un body
'*************************************************
Dim NewBody As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte
UserGenero = UserList(UserIndex).genero
UserRaza = UserList(UserIndex).raza
Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Enano
                NewBody = 291
            Case eRaza.Gnomo
                NewBody = 291
            Case Else
                NewBody = 226
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Enano
                NewBody = 291
            Case eRaza.Gnomo
                NewBody = 291
            Case Else
                NewBody = 226
        End Select
End Select
UserList(UserIndex).Char.body = NewBody
End Sub

Private Function ValidarCabeza(ByVal UserRaza As Byte, ByVal UserGenero As Byte, ByVal Head As Integer) As Boolean

Select Case UserGenero
    Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                ValidarCabeza = (Head >= 1 And _
                                Head <= 43)
            Case eRaza.Elfo
                ValidarCabeza = (Head >= 101 And _
                                Head <= 132)
            Case eRaza.Drow
                ValidarCabeza = (Head >= 201 And _
                                Head <= 230)
            Case eRaza.Enano
                ValidarCabeza = (Head >= 301 And _
                                Head <= 330)
            Case eRaza.Gnomo
                ValidarCabeza = (Head >= 401 And _
                                Head <= 430)
            Case eRaza.Orco
                ValidarCabeza = (Head >= 501 And _
                                Head <= 530)
            Case eRaza.NoMuerto
                ValidarCabeza = (Head >= 625 And _
                                Head <= 626)
        End Select
    
    Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                ValidarCabeza = (Head >= 70 And _
                                Head <= 100)
            Case eRaza.Elfo
                ValidarCabeza = (Head >= 170 And _
                                Head <= 200)
            Case eRaza.Drow
                ValidarCabeza = (Head >= 270 And _
                                Head <= 300)
            Case eRaza.Enano
                ValidarCabeza = (Head >= 370 And _
                                Head <= 399)
            Case eRaza.Gnomo
                ValidarCabeza = (Head >= 470 And _
                                Head <= 499)
            Case eRaza.Orco
                ValidarCabeza = (Head >= 570 And _
                                Head <= 599)
            Case eRaza.NoMuerto
                ValidarCabeza = (Head >= 650 And _
                                Head <= 651)
        End Select
End Select
        
End Function

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function

Sub ConnectNewUser(ByVal UserIndex As Integer, ByRef Name As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, _
                     ByVal CuentaName As String, ByVal Head As Integer, ByVal SerialHD As String)
'*************************************************
'Author: Unknown
'Last modified: 20/4/2007
'Conecta un nuevo Usuario
'23/01/2007 Pablo (ToxicWaste) - Agregué ResetFaccion al crear usuario
'24/01/2007 Pablo (ToxicWaste) - Agregué el nuevo mana inicial de los magos.
'12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
'20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
'09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
'*************************************************

    With UserList(UserIndex)
    
    
        If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
            Call WriteErrorMsg(UserIndex, "Nombre invalido.")
            Exit Sub
        End If
        
        If Len(Name) > 10 Then
            Call WriteErrorMsg(UserIndex, "El nombre debe tener menos de 10 letras.")
            Exit Sub
        End If
        
        If UserClase = Bard Then
            Call WriteErrorMsg(UserIndex, "Esta clase esta deshabilitada temporalmente. Disculpa las molestias.")
            Exit Sub
        End If
        
        If UserClase = Assasin Then
            Call WriteErrorMsg(UserIndex, "Esta clase esta deshabilitada temporalmente. Disculpa las molestias.")
            Exit Sub
        End If
        
        If .flags.UserLogged Then
            Call LogCheating("El usuario " & .Name & " ha intentado crear a " & Name & " desde la IP " & .ip)
            
            'Kick player ( and leave character inside :D )!
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            
            Exit Sub
        End If
        
        Dim LoopC As Long
        Dim totalskpts As Long
        
        '04/11/2015 Irongete: Compruebo si este nick ya está en uso
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT id_cuenta,id FROM personaje WHERE nombre = '" & Name & "' AND borrado = '0'")
        
        If Not (RS.EOF = True Or RS.BOF = True) Then
            If CInt(RS!id) > 0 Then
                Call WriteErrorMsg(UserIndex, "El nombre del personaje está en uso.")
                Exit Sub
            End If
        End If
        
       
        If Not ValidarCabeza(UserRaza, UserSexo, Head) Then
            Call LogCheating("El usuario " & Name & " ha seleccionado la cabeza " & Head & " desde la IP " & .ip)
            
            Call WriteErrorMsg(UserIndex, "Cabeza inválida, elija una cabeza seleccionable.")
            Exit Sub
        End If
        
        .flags.Muerto = 0
        .flags.Escondido = 0
        .flags.SerialHD = SerialHD
        
        .Reputacion.AsesinoRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.BurguesRep = 0
        .Reputacion.LadronesRep = 0
        .Reputacion.NobleRep = 1000
        .Reputacion.PlebeRep = 30
        
        .Reputacion.Promedio = 30 / 6
        
        
        .Name = Name
        .clase = UserClase
        .raza = UserRaza
        .genero = UserSexo
        
        '[Pablo (Toxic Waste) 9/01/08]
        .Stats.UserAtributos(eAtributos.Fuerza) = ModClase(UserClase).Fuerza + ModRaza(UserRaza).Fuerza
        .Stats.UserAtributos(eAtributos.Agilidad) = ModClase(UserClase).Agilidad + ModRaza(UserRaza).Agilidad
        .Stats.UserAtributos(eAtributos.Inteligencia) = ModClase(UserClase).Inteligencia + ModRaza(UserRaza).Inteligencia
        .Stats.UserAtributos(eAtributos.Energia) = ModClase(UserClase).Energia + ModRaza(UserRaza).Energia
        .Stats.UserAtributos(eAtributos.Constitucion) = ModClase(UserClase).Constitucion + ModRaza(UserRaza).Constitucion
        '[/Pablo (Toxic Waste)]
        
        '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
        
        .Char.heading = eHeading.SOUTH
        
        Call DarCuerpo(UserIndex)
        .Char.Head = Head
        .OrigChar = .Char
           
         
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
        
        Dim MiInt As Long
        MiInt = .Stats.UserAtributos(eAtributos.Constitucion)
        
        .Stats.MaxHP = 10 + MiInt
        .Stats.MinHP = 10 + MiInt
        
        If UserClase = eClass.Warrior Or UserClase = eClass.Hunter Or UserClase = eClass.Assasin Then
            MiInt = 70
        Else
            MiInt = 50
        End If
        
        .Stats.MaxSta = MiInt
        .Stats.MinSta = MiInt
        
        .Stats.MaxAGU = 100
        .Stats.MinAGU = 100
        
        .Stats.MaxHam = 100
        .Stats.MinHam = 100
        
        .Stats.ELO = 1000
        
        '<-----------------MANA----------------------->
        If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
            MiInt = .Stats.UserAtributos(eAtributos.Inteligencia) * 3
            .Stats.MaxMAN = MiInt
            .Stats.MinMAN = MiInt
        ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid _
            Or UserClase = eClass.Bard Or UserClase = eClass.Assasin Then
                .Stats.MaxMAN = 50
                .Stats.MinMAN = 50
        Else
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
        End If
        
         Select Case UserClase
            Case eClass.Cleric
                .Stats.UserHechizos(1) = 2
                .Stats.UserHechizos(2) = 1
            Case eClass.Mage
                .Stats.UserHechizos(1) = 2
            Case eClass.Assasin
                .Stats.UserHechizos(1) = 34
            Case eClass.Druid
                .Stats.UserHechizos(1) = 2
                .Stats.UserHechizos(2) = 16
            Case eClass.Paladin
                .Stats.UserHechizos(1) = 18
         End Select
        
        .Stats.MaxHIT = 2
        .Stats.MinHIT = 1
        
        .Stats.GLD = 0
        
        .Stats.Exp = 0
        .Stats.ELU = 300
        .Stats.ELV = 1
        
        Dim i As Byte
        
        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = 0
        Next i
        
        Call EquiparNewbie(UserIndex, UserClase, UserRaza)
         
        #If ConUpTime Then
            .LogOnTime = Now
            .UpTime = 0
        #End If
        
        'Valores Default de facciones al Activar nuevo usuario
        Call ResetFacciones(UserIndex)
        
        .Accounted = CuentaName
        
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT id FROM cuenta WHERE mail = '" & CuentaName & "'")
        UserList(UserIndex).CuentaId = RS!id
        
        '05/11/2015 Irongete: Cargar los objetos de la boveda
        Dim ln As String
        Set RS = SQL.Execute("SELECT slot,objeto FROM rel_cuenta_boveda WHERE id_cuenta = '" & .CuentaId & "'")
        UserList(UserIndex).BancoInvent.NroItems = CInt(RS.RecordCount)
        i = 1
        While Not RS.EOF
            ln = RS!Objeto
            UserList(UserIndex).BancoInvent.Object(i).ObjIndex = CInt(ReadField(1, ln, 45))
            UserList(UserIndex).BancoInvent.Object(i).Amount = CInt(ReadField(2, ln, 45))
            i = i + 1
            RS.MoveNext
        Wend
        
        Call SaveNewUser(UserIndex)
          
        'Open User

        Call ConnectUser(UserIndex, Name, SerialHD, UserList(UserIndex).CuentaId)
    End With
End Sub

Private Sub EquiparNewbie(ByVal UserIndex As Integer, ByVal UserClase As eClass, ByVal UserRaza As eRaza)
    With UserList(UserIndex)
    
            '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
        .Invent.NroItems = 9
        
        .Invent.Object(1).ObjIndex = 467
        .Invent.Object(1).Amount = 100
        
        .Invent.Object(2).ObjIndex = 468
        .Invent.Object(2).Amount = 100
        
        'Pociones rojas
        .Invent.Object(3).ObjIndex = 461
        .Invent.Object(3).Amount = 50
        
        'Amuleto de Ank
        '.Invent.Object(4).ObjIndex = 1023
        '.Invent.Object(4).Amount = 1
        
        '*******************************************
        'VESTIMENTAS POR RAZA
        Select Case UserRaza
            Case eRaza.Enano
                .Invent.Object(5).ObjIndex = 466
            Case eRaza.Gnomo
                .Invent.Object(5).ObjIndex = 466
            Case Else
                .Invent.Object(5).ObjIndex = 463
        End Select
        
        .Invent.Object(5).Amount = 1
        .Invent.Object(5).Equipped = 1
        '*******************************************
        
        Select Case UserClase
            Case eClass.Paladin
                .Invent.Object(6).ObjIndex = 1019
                .Invent.Object(6).Amount = 1
                .Invent.Object(7).ObjIndex = 1018 'Pociones amarillas
                .Invent.Object(7).Amount = 25
                .Invent.Object(8).ObjIndex = 462 'Pociones verdes
                .Invent.Object(8).Amount = 25
            Case eClass.Warrior
                .Invent.Object(6).ObjIndex = 464
                .Invent.Object(6).Amount = 1
                .Invent.Object(7).ObjIndex = 1018 'Pociones amarillas
                .Invent.Object(7).Amount = 25
                .Invent.Object(8).ObjIndex = 462
                .Invent.Object(8).Amount = 25
            Case eClass.Hunter
                .Invent.Object(6).ObjIndex = 1021
                .Invent.Object(6).Amount = 1
                .Invent.Object(7).ObjIndex = 1018 'Pociones amarillas
                .Invent.Object(7).Amount = 25
                .Invent.Object(8).ObjIndex = 462 'Pociones verdes
                .Invent.Object(8).Amount = 25
                .Invent.Object(9).ObjIndex = 1022
                .Invent.Object(9).Amount = 200
            Case eClass.Assasin
                .Invent.Object(6).ObjIndex = 460
                .Invent.Object(6).Amount = 1
                .Invent.Object(7).ObjIndex = 1018 'Pociones amarillas
                .Invent.Object(7).Amount = 25
                .Invent.Object(8).ObjIndex = 462 'Pociones verdes
                .Invent.Object(8).Amount = 25
             Case eClass.Mage, eClass.Cleric, eClass.Druid
                .Invent.Object(6).ObjIndex = 1020
                .Invent.Object(6).Amount = 1
                .Invent.Object(7).ObjIndex = 465 'Pocines azules
                .Invent.Object(7).Amount = 100
             Case eClass.Bard
                .Invent.Object(6).ObjIndex = 1019
                .Invent.Object(6).Amount = 1
                .Invent.Object(7).ObjIndex = 1018 'Pociones amarillas
                .Invent.Object(7).Amount = 25
                .Invent.Object(7).ObjIndex = 462 'Pociones verdes
                .Invent.Object(7).Amount = 25
                .Invent.Object(8).ObjIndex = 465 'Pocines azules
                .Invent.Object(8).Amount = 100
        End Select
        
        .Invent.ArmourEqpSlot = 4
        .Invent.ArmourEqpObjIndex = .Invent.Object(4).ObjIndex
    
    End With
End Sub

Sub CloseSocket(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: 27/07/2011 - ^[GS]^
'***************************************************

On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Call SecurityIp.IpRestarConexion(.IPLong)
        
        If .ConnID <> -1 Then
            Call CloseSocketSL(UserIndex)
        End If
        
        'Es el mismo user al que está revisando el centinela??
        'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
        ' y lo podemos loguear
        Dim CentinelaIndex As Byte
        CentinelaIndex = .flags.CentinelaIndex
        
        If CentinelaIndex <> 0 Then
            Call modCentinela.CentinelaUserLogout(CentinelaIndex)
        End If
        
        'mato los comercios seguros
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(.ComUsu.DestUsu)
                    Call FlushBuffer(.ComUsu.DestUsu)
                End If
            End If
        End If
        
        'Empty buffer for reuse
        Call .incomingData.ReadASCIIStringFixed(.incomingData.length)
        
        If .flags.UserLogged Then
            If NumUsers > 0 Then NumUsers = NumUsers - 1
            'Actualizo el frmMain. / maTih.-  |  02/03/2012
            If frmMain.Visible Then frmMain.CantUsuarios.Caption = CStr(NumUsers)
            
            Call CloseUser(UserIndex)
            
            'Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
        Else
            Call ResetUserSlot(UserIndex)
        End If
        
        .ConnID = -1
        .ConnIDValida = False
        
        
        If UserIndex = LastUser Then
            Do Until UserList(LastUser).ConnID <> -1
                LastUser = LastUser - 1
                If LastUser < 1 Then Exit Do
            Loop
        End If
    End With
    
Exit Sub

errHandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    Call ResetUserSlot(UserIndex)
    
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).ConnID <> -1
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)
End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************


If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
#If SocketType = 1 Then
    Call BorraSlotSock(UserList(UserIndex).ConnID)
    Call WSApiCloseSocket(UserList(UserIndex).ConnID, UserIndex)
#ElseIf SocketType = 2 Then
    frmMain.wskClient(UserIndex).Close
#End If
    UserList(UserIndex).ConnIDValida = False
End If

End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
' @remarks If UsarQueSocket is 3 it won`t use the clsByteQueue

Public Function EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String) As Long
'***************************************************
'Author: Unknownn
'Last Modification: 15/10/2011 - ^[GS]^
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
'***************************************************
On Error GoTo Err

#If SocketType = 1 Or SocketType = 2 Then '**********************************************
    
    Dim ret As Long
    
    ret = WsApiEnviar(UserIndex, Datos)
    
    If ret <> 0 And ret <> WSAEWOULDBLOCK Then
        ' Close the socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If
#End If '**********************************************

Exit Function
    
Err:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & UserIndex & "/" & IIf(UserIndex = 0, "nil", UserList(UserIndex).ConnID) & "/" & Datos)

End Function

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As worldPos) As Boolean

Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < XMaxMapSize + 1 And Y < YMaxMapSize + 1 Then
                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As worldPos, ObjIndex As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function
Function ValidateChr(ByVal UserIndex As Integer) As Boolean

ValidateChr = UserList(UserIndex).Char.Head <> 0 _
                And UserList(UserIndex).Char.body <> 0 _
                And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, ByRef Name As String, ByVal SerialHD As String, ByRef CuentaId As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/06/2009
'26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
'12/06/2009: ZaMa - Agrego chequeo de nivel al loguear
'***************************************************
Dim n As Integer
Dim tStr As String

    With UserList(UserIndex)
        
        .CuentaId = CuentaId
        
        
        If .flags.UserLogged Then
            Call LogCheating("El usuario " & .Name & " ha intentado loguear a " & Name & " desde la IP " & .ip)
            
            'Kick player ( and leave character inside :D )!
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            
            Exit Sub
        End If
        
        'Reseteamos los FLAGS
        .flags.Escondido = 0
        .flags.TargetNPC = 0
        .flags.TargetNpcTipo = eNPCType.Comun
        .flags.targetObj = 0
        .flags.targetUser = 0
        .Char.FX = 0
        .flags.SerialHD = SerialHD
        
        'Controlamos no pasar el maximo de usuarios
        If NumUsers >= MaxUsers Then
            Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        '¿Este IP ya esta conectado?
        If AllowMultiLogins = 0 Then
            If CheckForSameIP(UserIndex, .ip) = True Then
                Call WriteErrorMsg(UserIndex, "No es posible usar mas de un personaje al mismo tiempo.")
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
        End If
        
        '¿Este HD ya esta conectado?
        If AllowMultiLogins = 0 Then ' GSZAO
            If CheckForSameHD(UserIndex, SerialHD) = True Then
                Call WriteErrorMsg(UserIndex, "No es posible usar más de un personaje al mismo tiempo.")
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
        End If
    
        '¿Ya esta conectado el personaje?
        If CheckForSameName(Name) Then
            If UserList(NameIndex(Name)).Counters.Saliendo Then
                Call WriteErrorMsg(UserIndex, "El usuario está saliendo.")
            Else
                Call WriteErrorMsg(UserIndex, "Perdón, un usuario con el mismo nombre se ha logueado.")
            End If
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        'Reseteamos los privilegios
        .flags.Privilegios = 0
        
        'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
        If EsAdmin(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
            Call LogGM(Name, "Se conecto con ip:" & .ip)
        ElseIf EsDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
            Call LogGM(Name, "Se conecto con ip:" & .ip)
        ElseIf EsSemiDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
            Call LogGM(Name, "Se conecto con ip:" & .ip)
        ElseIf EsConsejero(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
            Call LogGM(Name, "Se conecto con ip:" & .ip)
        Else
            .flags.Privilegios = .flags.Privilegios Or PlayerType.User
            .flags.AdminPerseguible = True
        End If
        
        'Add RM flag if needed
        If EsRolesMaster(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
        End If
        
        If ServerSoloGMs > 0 Then
            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
                Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
        End If
        
        'Cargamos el personaje
        Dim Leer As New clsIniReader
        
        'Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")

        'Cargamos los datos del personaje
        Call LoadUserInit(UserIndex, Name)
        
        Call LoadUserStats(UserIndex, Name)
        
        Call LoadUserMonturas(UserIndex, Leer)
        
        If Not ValidateChr(UserIndex) Then
            Call WriteErrorMsg(UserIndex, "Error en el personaje.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        Call LoadUserReputacion(UserIndex, Name)
        
        Set Leer = Nothing
        
        If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
        If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
        If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
        
        If (.flags.Muerto = 0) Then
            .flags.SeguroResu = False
            Call WriteConsoleMsg(UserIndex, "Seguro de resurrección desactivado.", FontTypeNames.FONTTYPE_INFO)
        Else
            .flags.SeguroResu = True
            Call WriteConsoleMsg(UserIndex, "Seguro de resurrección activado.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call UpdateUserInv(True, UserIndex, 0)
        Call UpdateUserHechizos(True, UserIndex, 0)
        
        If .flags.Paralizado Then
            Call WriteParalizeOK(UserIndex)
        End If
        
        ''
        'TODO : Feo, esto tiene que ser parche cliente
        If .flags.Estupidez = 0 Then
            Call WriteDumbNoMore(UserIndex)
        End If
        
        '14/02/2016 Lorwik: Si el usuario conecta en el mapa de torneo se le lleva a nix.
        Call DesconectarDuelos(UserIndex)
        
        'Posicion de comienzo
        If .Pos.Map = 0 Then
                    .Pos.Map = 1
                    .Pos.X = 44
                    .Pos.Y = 88
        Else
            If Not MapaValido(.Pos.Map) Then
                Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
        End If
        
        'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
        'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex <> 0 Then
            Dim FoundPlace As Boolean
            Dim esAgua As Boolean
            Dim tX As Long
            Dim tY As Long
            
            FoundPlace = False
            esAgua = HayAgua(.Pos.Map, .Pos.X, .Pos.Y)
            
            For tY = .Pos.Y - 1 To .Pos.Y + 1
                For tX = .Pos.X - 1 To .Pos.X + 1
                    If esAgua Then
                        'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                        If LegalPos(.Pos.Map, tX, tY, True, False) Then
                            FoundPlace = True
                            Exit For
                        End If
                    Else
                        'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                        If LegalPos(.Pos.Map, tX, tY, False, True) Then
                            FoundPlace = True
                            Exit For
                        End If
                    End If
                Next tX
                
                If FoundPlace Then _
                    Exit For
            Next tY
            
            If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
                .Pos.X = tX
                .Pos.Y = tY
            Else
                'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Then
                    'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                    If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then
                        'Le avisamos al que estaba comerciando que se tuvo que ir.
                        If UserList(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                            Call FinComerciarUsu(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                            Call WriteConsoleMsg(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                            Call FlushBuffer(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                        End If
                        'Lo sacamos.
                        If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
                            Call FinComerciarUsu(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)
                            Call WriteErrorMsg(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                            Call FlushBuffer(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)
                        End If
                    End If
                    
                    Call CloseSocket(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)
                End If
            End If
        End If
        
        'Nombre de sistema
        .Name = Name
        
        .showName = True 'Por default los nombres son visibles
        
        'If in the water, and has a boat, equip it!
        If .Invent.BarcoObjIndex > 0 And _
                (HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.body)) Then
            Dim Barco As ObjData
            Barco = ObjData(.Invent.BarcoObjIndex)
            .Char.Head = 0
            If .flags.Muerto = 0 Then
        
                If .Faccion.ArmadaReal = 1 Or .Faccion.Legion = 1 Then
                    .Char.body = iFragataReal
                ElseIf .Faccion.FuerzasCaos = 1 Then
                    .Char.body = iFragataCaos
                Else
                    If criminal(UserIndex) Then
                        If Barco.Ropaje = iBarca Then .Char.body = iBarcaPk
                        If Barco.Ropaje = iGalera Then .Char.body = iGaleraPk
                        If Barco.Ropaje = iGaleon Then .Char.body = iGaleonPk
                    Else
                        If Barco.Ropaje = iBarca Then .Char.body = iBarcaCiuda
                        If Barco.Ropaje = iGalera Then .Char.body = iGaleraCiuda
                        If Barco.Ropaje = iGaleon Then .Char.body = iGaleonCiuda
                    End If
                End If
            Else
                .Char.body = iFragataFantasmal
            End If
            
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .flags.Navegando = 1
        End If
    
        '<Edurne>
        For n = 1 To NUMCASTILLOS
            Call WriteCastleAttack(UserIndex, n)
        Next n
        '/<Edurne>
    
        '**************Lorwik/Noche**************
        Call WriteNoche(UserIndex, DayStatus)
        '******************************************
            
        'Info
        Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
        Call WriteChangeMap(UserIndex, .Pos.Map, MapInfo(.Pos.Map).MapVersion) 'Carga el mapa
        
        '15/12/2018 Irongete: Le mando las zonas del mapa
        Call EnviarZonas(UserIndex, .Pos.Map)
        
        If .flags.Privilegios = PlayerType.Dios Then
            .flags.ChatColor = RGB(250, 250, 150)
        ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 0)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 255)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
            .flags.ChatColor = RGB(255, 128, 64)
        Else
            .flags.ChatColor = vbWhite
        End If
        
        
        ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
        #If ConUpTime Then
            .LogOnTime = Now
        #End If
        
        
        '¿Tiene equipado objetos de aumento de vida, mana, fuerza o agilidad?
        If .Invent.ArmourEqpObjIndex > 0 Then Call EquiparItemBonificador(UserIndex, .Invent.ArmourEqpObjIndex)
        If .Invent.CascoEqpObjIndex > 0 Then Call EquiparItemBonificador(UserIndex, .Invent.CascoEqpObjIndex)
        If .Invent.EscudoEqpObjIndex > 0 Then Call EquiparItemBonificador(UserIndex, .Invent.EscudoEqpObjIndex)
        If .Invent.WeaponEqpObjIndex > 0 Then Call EquiparItemBonificador(UserIndex, .Invent.WeaponEqpObjIndex)
        
        If .flags.Navegando = 1 Then
            Call WriteNavigateToggle(UserIndex)
            '.flags.Speed = Lorwik> Hay que modificar este sistema para darle velocidad
        End If
        
        '<Edurne>
        If .flags.QueMontura Then
           Call WriteMontateToggle(UserIndex)
           .flags.Speed = .flags.Montura(UserList(UserIndex).flags.QueMontura).Speed
        End If

        'Crea  el personaje del usuario
        Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
        
        Call WriteUserCharIndexInServer(UserIndex)
        ''[/el oso]
        
        Call CheckUserLevel(UserIndex)
        Call WriteUpdateUserStats(UserIndex)
        
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
        
        Call SendMOTD(UserIndex)
        
        If HaciendoBackup Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, WorldSave esta ejecutandose.", FontTypeNames.FONTTYPE_SERVER)
        End If
        
        If EnPausa Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
        End If
        
        If EnTesting And .Stats.ELV >= 18 Then
            Call WriteErrorMsg(UserIndex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        'Actualiza el Num de usuarios
        'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
        NumUsers = NumUsers + 1
        .flags.UserLogged = True
        
        '18/11/2015 Irongete: Lo pongo online en la base de datos
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("UPDATE personaje SET logged = '1' WHERE id = '" & UserList(UserIndex).id & "'")
        
        '18/11/2015 Irongete: Apunto su UserIndex en la base de datos
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("UPDATE personaje SET userindex = '" & UserIndex & "' WHERE id = '" & UserList(UserIndex).id & "'")
        
        Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
        
        MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        
        If NumUsers > recordusuarios Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
            recordusuarios = NumUsers
            Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
            
            Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
        End If
        
        If .NroMascotas > 0 And MapInfo(.Pos.Map).Pk Then
            Dim i As Integer
            For i = 1 To MAXMASCOTAS
                If .MascotasType(i) > 0 Then
                    .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)
                    
                    If .MascotasIndex(i) > 0 Then
                        NPCList(.MascotasIndex(i)).MaestroUser = UserIndex
                        Call FollowAmo(.MascotasIndex(i))
                    Else
                        .MascotasIndex(i) = 0
                    End If
                End If
            Next i
        End If
        
        '15/02/2016 Irongete: Si está en party le envio la Id de la misma
        .PartyId = 0
        Dim PartyId As Integer
        PartyId = GetPartyId(UserIndex)
        If PartyId > 0 Then
          .PartyId = PartyId
        End If
        Call WriteSetPartyId(UserIndex, 0, .PartyId)
        
        If .GuildIndex > 0 Then
           UserList(UserIndex).GuildIndex = .GuildIndex
        End If
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
        
        ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
        Call IntervaloPermiteSerAtacado(UserIndex, True)
        
        If .flags.Seguro = 0 Then
            Call WriteSegNoti(UserIndex, 1, False)
        Else
            Call WriteSegNoti(UserIndex, 1, True)
        End If
        
        Call WriteLoggedMessage(UserIndex)
        
        'Call modGuilds.SendGuildNews(UserIndex)
        
        'tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)
        
        'If LenB(tStr) <> 0 Then
        '    Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
        'End If
        
        'Load the user statistics
        Call Statistics.UserConnected(UserIndex)
        
        Call MostrarNumUsers
        
        If UserIndex = GranPoder Then
            Call OtorgarGranPoder(0)
            GranPoder = False
        End If
        
        n = FreeFile
        Open App.Path & "\logs\numusers.log" For Output As n
            Print #n, NumUsers
        Close #n
        
        n = FreeFile
        'Log
        Open App.Path & "\logs\Connect.log" For Append Shared As #n
            Print #n, .Name & " ha entrado al juego. UserIndex:" & UserIndex & " " & time & " " & Date
        Close #n
    End With

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
    Dim j As Long
    
    'Call WriteConsoleMsg(userIndex, "Mensajes de entrada:", FontTypeNames.FONTTYPE_INFO)
    For j = 1 To MaxLines
        Call WriteConsoleMsg(UserIndex, MOTD(j).texto, FontTypeNames.FONTTYPE_INFO)
    Next j
    
    '15/03/2016 Lorwik
    Call DueñosCastillos(UserIndex)
    
End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .Legion = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .FuerzasCaos = 0
        .FechaIngreso = "No ingresó a ninguna Facción"
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .NextRecompensa = 0
    End With
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'*************************************************
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        .bPuedeMeditar = False
        .Lava = 0
        .Mimetismo = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
        .Makro = 0
    End With
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex)
        .Name = vbNullString
        .desc = vbNullString
        .DescRM = vbNullString
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = vbNullString
        .clase = 0
        .genero = 0
        .raza = 0
        

        .PartyId = 0
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
        End With
        
    End With
End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0
    End With
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0
    End If
    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)
    End If
    UserList(UserIndex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 06/28/2008
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'06/28/2008 NicoNZ - Agrego el flag Inmovilizado
'*************************************************
    With UserList(UserIndex).flags
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .targetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .targetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .Vuela = 0
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .Silenciado = 0
        .CentinelaOK = False
        .CentinelaIndex = 0
        .AdminPerseguible = False
        .Anomalia = 0
        .DuelosClasicos = 0
        .Morph = 0
        .TimeDueloSet = 0
        .NoPuedeSerAtacado = 0
        .PuedeCambiarMapa = 0
        .AumentodeFuerza = 0
        .AumentodeAgilidad = 0
        .AumentodeVida = 0
        .AumentodeMana = 0
        .ArenaRinkel = False
    End With
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    UserList(UserIndex).NroMascotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasIndex(LoopC) = 0
        UserList(UserIndex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
    With UserList(UserIndex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)

UserList(UserIndex).ConnIDValida = False
UserList(UserIndex).ConnID = -1

Call LimpiarComercioSeguro(UserIndex)
Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetReputacion(UserIndex)
Call ResetGuildInfo(UserIndex)
Call ResetUserFlags(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)
Call ResetUserBanco(UserIndex)
With UserList(UserIndex).ComUsu
    .Acepto = False
    .cant = 0
    .DestNick = vbNullString
    .DestUsu = 0
    .Objeto = 0
End With

End Sub

Sub CloseUser(ByVal UserIndex As Integer)
'Call LogTarea("CloseUser " & UserIndex)
On Error GoTo errHandler

Dim n As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim Name As String
Dim i As Integer

Dim aN As Integer

'14/02/2016 Lorwik: Si sale pierde el gran poder y que rule...
If UserIndex = GranPoder Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha perdido el Gran Poder", FontTypeNames.FONTTYPE_WARNING))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FxGranPoder, 0))
    Call OtorgarGranPoder(0)
End If

aN = UserList(UserIndex).flags.AtacadoPorNpc
If aN > 0 Then
      NPCList(aN).Movement = NPCList(aN).flags.OldMovement
      NPCList(aN).Hostile = NPCList(aN).flags.OldHostil
      NPCList(aN).flags.AttackedBy = vbNullString
End If
aN = UserList(UserIndex).flags.NPCAtacado
If aN > 0 Then
    If NPCList(aN).flags.AttackedFirstBy = UserList(UserIndex).Name Then
        NPCList(aN).flags.AttackedFirstBy = vbNullString
    End If
End If
UserList(UserIndex).flags.AtacadoPorNpc = 0
UserList(UserIndex).flags.NPCAtacado = 0

Map = UserList(UserIndex).Pos.Map
Name = UCase$(UserList(UserIndex).Name)

UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.loops = 0
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))


UserList(UserIndex).flags.UserLogged = False
UserList(UserIndex).Counters.Saliendo = False

'Le devolvemos el body y head originales
If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

'Save statistics
Call Statistics.UserDisconnected(UserIndex)

' Grabamos el personaje del usuario
Call SaveUser(UserIndex)


'18/11/2015 Irongete: Lo pongo offline en la base de datos
Set RS = New ADODB.Recordset
Set RS = SQL.Execute("UPDATE personaje SET logged = '0' WHERE id = '" & UserList(UserIndex).id & "'")

'18/11/2015 Irongete: Pongo a 0 su UserIndex en la base de datos
'03/02/2016 Lorwik: Perdera el UserIndex cuando desconecte de la cuenta.
'Set RS = New ADODB.Recordset
'Set RS = SQL.Execute("UPDATE personaje SET userindex = '0' WHERE id = '" & UserList(UserIndex).Id & "'")

'Quitar el dialogo
'If MapInfo(Map).NumUsers > 0 Then
'    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
'End If

If MapInfo(Map).NumUsers > 0 Then
    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
End If

'Borrar el personaje
If UserList(UserIndex).Char.CharIndex > 0 Then
    Call EraseUserChar(UserIndex, UserList(UserIndex).flags.AdminInvisible = 1)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If NPCList(UserList(UserIndex).MascotasIndex(i)).flags.Active Then _
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)

Call ResetUserSlot(UserIndex)

Call MostrarNumUsers

n = FreeFile(1)
Open App.Path & "\logs\Connect.log" For Append Shared As #n
Print #n, Name & " ha dejado el juego. " & "User Index:" & UserIndex & " " & time & " " & Date
Close #n

Exit Sub

errHandler:
Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.Description)

End Sub

Sub ReloadSokcet()
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

On Error GoTo errHandler
#If SocketType = 1 Or SocketType = 2 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
        #If SocketType = 1 Then
            Call apiclosesocket(SockListen)
            SockListen = ListenForConnect(Puerto, hWndMsg, "")
        #ElseIf SocketType = 2 Then
            frmMain.wskListen.Close
            frmMain.wskListen.LocalPort = Puerto
            frmMain.wskListen.listen
        #End If
    End If
#End If

Exit Sub
errHandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.Description)

End Sub

Public Sub EnviarNoche(ByVal UserIndex As Integer)
    Call WriteSendNight(UserIndex, IIf(DeNoche And (MapInfo(UserList(UserIndex).Pos.Map).zona = Campo Or MapInfo(UserList(UserIndex).Pos.Map).zona = Ciudad), True, False))
    Call WriteSendNight(UserIndex, IIf(DeNoche, True, False))
End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios And PlayerType.User Then
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub
