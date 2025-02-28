Attribute VB_Name = "SecurityIp"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez


'**************************************************************
' General_IpSecurity.Bas - Maneja la seguridad de las IPs
'
' Escrito y dise�ado por DuNga (ltourrilhes@gmail.com)
'**************************************************************
Option Explicit

'*************************************************  *************
' General_IpSecurity.Bas - Maneja la seguridad de las IPs
'
' Escrito y dise�ado por DuNga (ltourrilhes@gmail.com)
'*************************************************  *************

Private IpTables()      As Long 'USAMOS 2 LONGS: UNO DE LA IP, SEGUIDO DE UNO DE LA INFO
Private EntrysCounter   As Long
Private MaxValue        As Long
Private Multiplicado    As Long 'Cuantas veces multiplike el EntrysCounter para que me entren?
Private Const IntervaloEntreConexiones As Long = 1100

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Declaraciones para maximas conexiones por usuario
'Agregado por EL OSO
Private MaxConTables()      As Long
Private MaxConTablesEntry   As Long     'puntero a la ultima insertada

Private Const LIMITECONEXIONESxIP As Long = 100

Private Enum e_SecurityIpTabla
    IP_INTERVALOS = 1
    IP_LIMITECONEXIONES = 2
End Enum

Public Sub InitIpTables(ByVal OptCountersValue As Long)
'*************************************************  *************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modification: 10/08/2011 - ^[GS]^
'
'*************************************************  *************
    EntrysCounter = OptCountersValue
    Multiplicado = 1

    ReDim IpTables(EntrysCounter * 2 - 1) As Long
    MaxValue = 0

    ReDim MaxConTables(Declaraciones.MaxUsers * 2 - 1) As Long
    MaxConTablesEntry = 0

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''FUNCIONES PARA INTERVALOS'''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub IpSecurityMantenimientoLista()
'*************************************************  *************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modification: 10/08/2011 - ^[GS]^
'*************************************************  *************
    'Las borro todas cada 1 hora, asi se "renuevan"
    EntrysCounter = EntrysCounter \ Multiplicado
    Multiplicado = 1
    ReDim IpTables(EntrysCounter * 2 - 1) As Long
    MaxValue = 0
End Sub

Public Function IpSecurityAceptarNuevaConexion(ByVal ip As Long) As Boolean
'*************************************************  *************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 27/07/2012 - ^[GS]^
'*************************************************  *************
    Dim IpTableIndex As Long
    Dim tmpTime As Long
    
    IpTableIndex = FindTableIp(ip, IP_INTERVALOS)
    tmpTime = (GetTickCount() And &H7FFFFFFF)
    
    If IpTableIndex >= 0 Then
        If IpTables(IpTableIndex + 1) + IntervaloEntreConexiones <= tmpTime Then   'No est� saturando de connects?
            IpTables(IpTableIndex + 1) = (tmpTime And &H3FFFFFFF)
            IpSecurityAceptarNuevaConexion = True
            Debug.Print "CONEXION ACEPTADA"
            'Exit Function
        Else
            IpSecurityAceptarNuevaConexion = False
            Debug.Print "CONEXION NO ACEPTADA"
            Exit Function
        End If
    Else
        IpTableIndex = Not IpTableIndex
        AddNewIpIntervalo ip, IpTableIndex
        IpTables(IpTableIndex + 1) = (tmpTime And &H3FFFFFFF)
        IpSecurityAceptarNuevaConexion = True
        'Exit Function
    End If

End Function


Private Sub AddNewIpIntervalo(ByVal ip As Long, ByVal Index As Long)
'*************************************************  *************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modification: 10/08/2011 - ^[GS]^
'
'*************************************************  *************
    '2) Pruebo si hay espacio, sino agrando la lista
    If MaxValue + 1 > EntrysCounter Then
        EntrysCounter = EntrysCounter \ Multiplicado
        Multiplicado = Multiplicado + 1
        EntrysCounter = EntrysCounter * Multiplicado
        
        ReDim Preserve IpTables(EntrysCounter * 2 - 1) As Long
    End If
    
    '4) Corro todo el array para arriba
    Call CopyMemory(IpTables(Index + 2), IpTables(Index), (MaxValue - Index \ 2) * 8)   '*4 (peso del long) * 2(cantidad de elementos por c/u)
    IpTables(Index) = ip
    
    '3) Subo el indicador de el maximo valor almacenado y listo :)
    MaxValue = MaxValue + 1
End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''FUNCIONES PARA LIMITES X IP''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function IPSecuritySuperaLimiteConexiones(ByVal ip As Long) As Boolean
Dim IpTableIndex As Long

    IpTableIndex = FindTableIp(ip, IP_LIMITECONEXIONES)
    
    If IpTableIndex >= 0 Then
        
        If MaxConTables(IpTableIndex + 1) < LIMITECONEXIONESxIP Then
            'LogIP ("Agregamos conexion a " & GetAscIP(IP) & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
            'Debug.Print "suma conexion a " & GetAscIP(IP) & " total " & MaxConTables(IpTableIndex + 1) + 1
            MaxConTables(IpTableIndex + 1) = MaxConTables(IpTableIndex + 1) + 1
            IPSecuritySuperaLimiteConexiones = False
        Else
            'LogIP ("rechazamos conexion de " & GetAscIP(IP) & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
            'Debug.Print "rechaza conexion a " & GetAscIP(IP)
            IPSecuritySuperaLimiteConexiones = True
        End If
    Else
        IPSecuritySuperaLimiteConexiones = False
        If MaxConTablesEntry < Declaraciones.MaxUsers Then  'si hay espacio..
            IpTableIndex = Not IpTableIndex
            AddNewIpLimiteConexiones ip, IpTableIndex    'iptableindex es donde lo agrego
            MaxConTables(IpTableIndex + 1) = 1
        Else
            Call LogCriticEvent("SecurityIP.IPSecuritySuperaLimiteConexiones: Se supero la disponibilidad de slots.")
        End If
    End If

End Function

Private Sub AddNewIpLimiteConexiones(ByVal ip As Long, ByVal Index As Long)
'*************************************************  *************
'Author: (EL OSO)
'Last Modification: 10/08/2011 - ^[GS]^
'05/21/10 - Pato: Saco el uso de buffer auxiliar
'*************************************************  *************
    'Debug.Print "agrega conexion a " & ip
    'Debug.Print "(modDeclaraciones.iniMaxUsuarios - index) = " & (modDeclaraciones.iniMaxUsuarios - Index)
    '4) Corro todo el array para arriba
    
    Call CopyMemory(MaxConTables(Index + 2), MaxConTables(Index), (MaxConTablesEntry - Index \ 2) * 8)    '*4 (peso del long) * 2(cantidad de elementos por c/u)
    MaxConTables(Index) = ip

    '3) Subo el indicador de el maximo valor almacenado y listo :)
    MaxConTablesEntry = MaxConTablesEntry + 1

End Sub

Public Sub IpRestarConexion(ByVal ip As Long)
'***************************************************
'Author: Unknownn
'Last Modification: 10/08/2011 - ^[GS]^
'
'***************************************************
On Error GoTo errhandler
    
    Dim key As Long
    'Debug.Print "resta conexion a " & ip
    
    key = FindTableIp(ip, IP_LIMITECONEXIONES)
    
    If key >= 0 Then
        If MaxConTables(key + 1) > 0 Then
            MaxConTables(key + 1) = MaxConTables(key + 1) - 1
        End If
        'Call LogIP("restamos conexion a " & ip & " key=" & key & ". Conexiones: " & MaxConTables(key + 1))
        'Comento esto, sino se nos va el HD en logs, jaja
        If MaxConTables(key + 1) <= 0 Then
            'la limpiamos
            MaxConTablesEntry = MaxConTablesEntry - 1
            
            If key + 2 < UBound(MaxConTables) Then
                Call CopyMemory(MaxConTables(key), MaxConTables(key + 2), (MaxConTablesEntry - (key \ 2)) * 8)
            End If
        End If
    Else 'Key < 0
        Call LogIP("restamos conexion a " & GetAscIP(ip) & " key=" & key & ". NEGATIVO!!")
        'LogCriticEvent "SecurityIp.IpRestarconexion obtuvo un valor negativo en key"
    End If
      
    Exit Sub

errhandler:

    Call LogError("Error en IpRestarConexion. Error: " & Err.Number & " - " & Err.Description & ". Ip: " & GetAscIP(ip) & " Key:" & key)

End Sub



' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''''''FUNCIONES GENERALES''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Function FindTableIp(ByVal ip As Long, ByVal Tabla As e_SecurityIpTabla) As Long
'*************************************************  *************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modification: 10/08/2011 - ^[GS]^
'Modified by Juan Mart�n Sotuyo Dodero (Maraxus) to use Binary Insertion
'*************************************************  *************
Dim First As Long
Dim Last As Long
Dim Middle As Long
    
    Select Case Tabla
        Case e_SecurityIpTabla.IP_INTERVALOS
            First = 0
            Last = MaxValue - 1
            Do While First <= Last
                Middle = (First + Last) \ 2
                
                If (IpTables(Middle * 2) < ip) Then
                    First = Middle + 1
                ElseIf (IpTables(Middle * 2) > ip) Then
                    Last = Middle - 1
                Else
                    FindTableIp = Middle * 2
                    Exit Function
                End If
            Loop
            FindTableIp = Not (First * 2)
        
        Case e_SecurityIpTabla.IP_LIMITECONEXIONES
            
            First = 0
            Last = MaxConTablesEntry - 1

            Do While First <= Last
                Middle = (First + Last) \ 2

                If MaxConTables(Middle * 2) < ip Then
                    First = Middle + 1
                ElseIf MaxConTables(Middle * 2) > ip Then
                    Last = Middle - 1
                Else
                    FindTableIp = Middle * 2
                    Exit Function
                End If
            Loop
            FindTableIp = Not (First * 2)
    End Select
    
End Function

Public Function DumpTables()
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

Dim i As Integer

    For i = 0 To MaxConTablesEntry * 2 - 1 Step 2
        Call LogCriticEvent(GetAscIP(MaxConTables(i)) & " > " & MaxConTables(i + 1))
    Next i

End Function

