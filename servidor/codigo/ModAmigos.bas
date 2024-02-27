Attribute VB_Name = "ModAmigos"
'****************************************
'SISTEMA DE AMIGOS
'****************************************

Public Sub AgregarAmigo(ByVal UserIndex As Integer, ByVal TName As String)

'*************************************************
'Llega la solicitud de una nueva amistad
'*************************************************

Dim tUser As Integer
Dim CantPeticiones As Byte
Dim i As Byte
Dim IndexFriend As Integer
Dim MisPeticiones As Byte

    With UserList(UserIndex)
        tUser = NameIndex(TName)

        IndexFriend = NameToIndex(TName)
        
        '///////////////////////////////////
        'COMPROBACIONES
        '//////////////////////////////////
        
        'Compruebo si se quiere agregar a si mismo
        If UCase(UserList(UserIndex).name) = UCase(TName) Then
            Call WriteConsoleMsg(UserIndex, "No puedes agregarte a ti mismo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Ya tienes a ese amigo
        For i = 1 To .flags.Amigos.CantFriend
            If UCase(SQLConsulta(.IndexPJ, "amigos", "Friend" & i)) = UCase(TName) Then
                Call WriteConsoleMsg(UserIndex, "El usuario ya se encuentra en tu lista de amigos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Next i

        'Comprobamos si ya supero el limite
        If .flags.Amigos.CantFriend >= 15 Then
            Call WriteConsoleMsg(UserIndex, "No puedes agregar mas amigos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        MisPeticiones = SQLConsulta(.IndexPJ, "amigos", "CantPeticiones")
        If Not MisPeticiones = 0 Then
            For i = 1 To MisPeticiones
                'Ya tienes a ese amigo
                If UCase(SQLConsulta(.IndexPJ, "amigos", "Peticion" & i)) = UCase(TName) Then
                    Call WriteConsoleMsg(UserIndex, "Ya hay una propuesta de amistad pendiente.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Next i
        End If
        
        CantPeticiones = SQLConsulta(NameToIndex(TName), "amigos", "CantPeticiones")
        
        'Comprobamos si ya supero el limite el amigo
        If CantPeticiones >= 15 Then
            Call WriteConsoleMsg(UserIndex, "Tu amigo no puede recibir mas peticiones por ahora.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Comprobamos si ya supero el maximo de amigos
        If .flags.Amigos.CantFriend >= 15 Then
            Call WriteConsoleMsg(UserIndex, "Tu amigo no puede aceptar mas amigos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '///////////////////////////////////
        'ESCRIBIMOS
        '//////////////////////////////////
        
        Debug.Print "Friend Antes: " & .flags.Amigos.CantFriend
        
        'Actualizamos el contador
        .flags.Amigos.CantFriend = .flags.Amigos.CantFriend + 1
        Debug.Print "Friend Despues: " & .flags.Amigos.CantFriend
        'Agregamos al nuevo amigo
        Call SQLSave(UserList(UserIndex).IndexPJ, "amigos", "TotalFriends", .flags.Amigos.CantFriend)
        Call SQLSave(UserList(UserIndex).IndexPJ, "amigos", "Friend" & .flags.Amigos.CantFriend, "'" & TName & "'")
        
        'Actualizamos la lista de amigos
        Call WriteEnviarAmigo(UserIndex, TName, True)
        
        'Añadimos el nombre del usuario solicitante a la lista de solicitudes del amigo que quiere agregar.
        Call SQLSave(IndexFriend, "amigos", "Peticion" & CantPeticiones + 1, "'" & .name & "'")
        Call SQLSave(IndexFriend, "amigos", "CantPeticiones", CantPeticiones + 1)
        
        '¿Esta online?
        If Not tUser <= 0 Then
            'Si el usuario esta online le decimos que le estan enviando una solicitud de amistad
            Call WriteConsoleMsg(tUser, .name & " te envio una solicitud de amistad.", FontTypeNames.FONTTYPE_INFO)
            'Actualizamos la lista de amigos del amigo al que queremos agregar para que le aparezca nuestro nombre.
            Call WriteEnviarAmigo(tUser, .name, False)
        End If
        
        Call WriteConsoleMsg(UserIndex, "Tu solicitud de amistad fue enviada, ahora solo debes esperar una respuesta.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub AceptarAmigo(ByVal UserIndex As Integer, ByVal TName As String)
'***********************************
'Aceptamos una solicitud de amistad
'**********************************

Dim CantPeticiones As Byte
Dim Slot As Byte
Dim PJIndexFriend As Integer
Dim i As Byte
Dim tUser As Integer

tUser = NameIndex(TName)

    With UserList(UserIndex)
    
        '///////////////////////////////////
        'COMPROBACIONES
        '//////////////////////////////////
        If TName = "" Then
            Call WriteConsoleMsg(UserIndex, "Selecciona un usuario.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Compruebo si se quiere agregar a si mismo (Prevenir Hacks)
        If UCase(UserList(UserIndex).name) = UCase(TName) Then
            Call WriteConsoleMsg(UserIndex, "No puedes agregarte a ti mismo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        CantPeticiones = SQLConsulta(.IndexPJ, "amigos", "CantPeticiones")
        
        If CantPeticiones = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡No tienes peticiones de amistad!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim Comprob As Byte
        For i = 1 To CantPeticiones
            If Not UCase(SQLConsulta(.IndexPJ, "amigos", "Peticion" & i)) = UCase(TName) Then
                Comprob = Comprob + 1
                If Comprob = CantPeticiones Then
                    Call WriteConsoleMsg(UserIndex, "¡Ese usuario no te envio ninguna solicitud de amistad!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        
             If UCase(SQLConsulta(.IndexPJ, "amigos", "Peticion" & i)) = UCase(TName) Then Slot = i
        Next i
        
        '///////////////////////////////////
        'ESCRIBIMOS
        '//////////////////////////////////
        
        'Eliminamos la peticion
        Call SQLSave(.IndexPJ, "amigos", "Peticion" & Slot, "''")
        'Bajamos la cantidad de peticiones y actualizamos
        CantPeticiones = CantPeticiones - 1
        Call SQLSave(.IndexPJ, "amigos", "CantPeticiones", CantPeticiones)
        
        'Actualizamos la lista de amigos
        .flags.Amigos.CantFriend = .flags.Amigos.CantFriend + 1
        Call SQLSave(.IndexPJ, "amigos", "Friend" & .flags.Amigos.CantFriend, "'" & TName & "'")
        Call SQLSave(.IndexPJ, "amigos", "Confirmed" & .flags.Amigos.CantFriend, 1)
        Call SQLSave(.IndexPJ, "amigos", "TotalFriends", .flags.Amigos.CantFriend)
        
        Dim Transpaso As String
        Transpaso = SQLConsulta(.IndexPJ, "amigos", "Peticion" & CantPeticiones + 1)
        Call SQLSave(.IndexPJ, "amigos", "Peticion" & CantPeticiones + 1, "''")
        Call SQLSave(.IndexPJ, "amigos", "Peticion" & Slot, "'" & Transpaso & "'")
        
        '########################################
        'Parte del amigo
        '########################################
        PJIndexFriend = NameToIndex(TName)
        CantPeticiones = SQLConsulta(PJIndexFriend, "amigos", "TotalFriends")

                    'Esto es TOTALFRIENDS
        For i = 1 To CantPeticiones
             If SQLConsulta(PJIndexFriend, "amigos", "Friend" & i) = .name Then Slot = i
        Next i
        
        Call SQLSave(PJIndexFriend, "amigos", "Confirmed" & Slot, 1)

        Call WriteEnviarAmigo(UserIndex, TName, True)
        Call WriteConsoleMsg(UserIndex, TName & " fue aceptado en tu lista de amigos.", FontTypeNames.FONTTYPE_INFO)
        '¿Esta online?
        If Not tUser <= 0 Then
            'Si el usuario esta online le decimos que le estan enviando una solicitud de amistad
            'Call WriteEnviarAmigo(tUser, .Name, True)
            Call ActualizarAmigosAll(tUser)
            Call WriteConsoleMsg(tUser, .name & " acepto tu solicitud de amistad.", FontTypeNames.FONTTYPE_INFO)
        End If
   End With
End Sub

Public Sub EliminarAmigo(ByVal UserIndex As Integer, ByVal TName As String)

Dim CantFriends As Byte
Dim Slot As Byte
Dim PJIndexFriend As Integer
Dim i As Byte
Dim N As Byte
Dim Transpaso As String
Dim tUser As Integer

    tUser = NameIndex(TName)
        
    With UserList(UserIndex)
        If .flags.Amigos.CantFriend = 0 And SQLConsulta(.IndexPJ, "amigos", "CantPeticiones") = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡No tienes ningun amigo!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Buscamos el slot donde se encuentra el amigo
        For i = 1 To 15
             If UCase(SQLConsulta(.IndexPJ, "amigos", "Friend" & i)) = UCase(TName) Then
                Slot = i
             ElseIf UCase(SQLConsulta(.IndexPJ, "amigos", "Peticion" & i)) = UCase(TName) Then
                Slot = i
                Call CancelarPeticion(UserIndex, TName, Slot)
                Exit Sub
             End If
        Next i
        
        '¿Se trata de una peticion que nosotros enviamos?
        If SQLConsulta(.IndexPJ, "amigos", "Confirmed" & Slot) = 0 Then
            Call EliminarPeticion(UserIndex, TName, Slot)
            Exit Sub
        End If
        
        If .flags.Amigos.CantFriend > 1 Then
            'Copio el ultimo pj de la lista
            Transpaso = SQLConsulta(.IndexPJ, "amigos", "Friend" & .flags.Amigos.CantFriend)
            
            'Paso al ultimo pj al slot libre que dejo el que borramos
            Call SQLSave(.IndexPJ, "amigos", "Friend" & Slot, "'" & Transpaso & "'")
            
            'Elimino al ultimo pj de la lista
            Call SQLSave(.IndexPJ, "amigos", "Friend" & .flags.Amigos.CantFriend, "''")
        
            'Borro la confirmacion del ultimo pj
            Call SQLSave(.IndexPJ, "amigos", "Confirmed" & .flags.Amigos.CantFriend, 0)
        Else
            'Elimino al pj de la lista
            Call SQLSave(.IndexPJ, "amigos", "Friend" & Slot, "''")
        
            'Borro la confirmacion del pj
            Call SQLSave(.IndexPJ, "amigos", "Confirmed" & Slot, 0)
        End If

        'Resto 1 a la cantidad de amigos
        .flags.Amigos.CantFriend = .flags.Amigos.CantFriend - 1
        Call SQLSave(.IndexPJ, "amigos", "TotalFriends", .flags.Amigos.CantFriend)
        
        Call WriteActualizarBorradoFriend(UserIndex, Slot)
        
        '########################################
        'Parte del amigo
        '########################################
        'buscamos el index de nuestor amigo
        PJIndexFriend = NameToIndex(TName)
        'Comprobamos su cantidad de amigos
        CantFriends = SQLConsulta(PJIndexFriend, "amigos", "TotalFriends")
        
        'Buscamos el slot donde nos encontramos en su lista
        For i = 1 To PJIndexFriend
             If UCase(SQLConsulta(PJIndexFriend, "amigos", "Friend" & i)) = UCase(TName) Then Slot = i
        Next i
        
        If CantFriends > 1 Then
            'Copio el ultimo pj de la lista
            Transpaso = SQLConsulta(PJIndexFriend, "amigos", "Friend" & CantFriends)
            
            'Paso al ultimo pj al slot libre que dejo el que borramos
            Call SQLSave(PJIndexFriend, "amigos", "Friend" & Slot, "'" & Transpaso & "'")
        End If
        
        'Elimino el pj de la lista
        Call SQLSave(PJIndexFriend, "amigos", "Friend" & CantFriends, "''")
        
        'Borro la confirmacion del ultimo pj
        Call SQLSave(PJIndexFriend, "amigos", "Confirmed" & CantFriends, 0)
        
        'Resto 1 a la cantidad de amigos
        Call SQLSave(PJIndexFriend, "amigos", "TotalFriends", CantFriends - 1)
        
        Call WriteConsoleMsg(UserIndex, TName & " fue eliminado de tu lista de amigos.", FontTypeNames.FONTTYPE_INFO)
        
        '¿Esta online?
        If Not tUser <= 0 Then
            'Si el usuario esta online le decimos que le estan enviando una solicitud de amistad
            UserList(tUser).flags.Amigos.CantFriend = UserList(tUser).flags.Amigos.CantFriend - 1
            Call WriteActualizarBorradoFriend(tUser, Slot)
            Call WriteConsoleMsg(tUser, .name & " te elimino de su lista de amigos.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Public Sub NotificarConexionAmigos(ByVal UserIndex As Integer, ByVal estado As Boolean)
    Dim AmigosIndex(1 To 15) As Integer
    Dim i As Byte
    
    With UserList(UserIndex)
        'Consultamos la cantidad de amigos que tiene agregado
        .flags.Amigos.CantFriend = SQLConsulta(.IndexPJ, "amigos", "TotalFriends")
        'Si no tiene amigos nos ahorramos todo esto
        If .flags.Amigos.CantFriend > 0 Then
            For i = 1 To .flags.Amigos.CantFriend
                AmigosIndex(i) = NameIndex(.flags.Amigos.Amigos(i))
                
                If .flags.Amigos.Confirmed(i) = True And Not AmigosIndex(i) <= 0 Then
                    'Notifico a mis amigos que me desconecte:
                    Call WriteEnviarAmigo(AmigosIndex(i), .name, True)
                    Call WriteConsoleMsg(AmigosIndex(i), .name & " se ha desconectado.", FontTypeNames.FONTTYPE_INFO)
                End If
            Next i
        End If
    End With
End Sub

Public Sub NotificarMapFriend(ByVal UserIndex As Integer)
    Dim AmigosIndex(1 To 15) As Integer
    Dim i As Byte
    
    With UserList(UserIndex)
        'Consultamos la cantidad de amigos que tiene agregado
        .flags.Amigos.CantFriend = SQLConsulta(.IndexPJ, "amigos", "TotalFriends")
        'Si no tiene amigos nos ahorramos todo esto
        If .flags.Amigos.CantFriend > 0 Then
            For i = 1 To .flags.Amigos.CantFriend
                AmigosIndex(i) = NameIndex(.flags.Amigos.Amigos(i))
                
                If .flags.Amigos.Confirmed(i) = True And Not AmigosIndex(i) <= 0 Then
                    'Notifico a mis amigos que me desconecte:
                    Call WriteEnviarAmigo(AmigosIndex(i), .name, True)
                End If
            Next i
        End If
    End With
End Sub

Public Sub ActualizarAmigosAll(ByVal UserIndex As Integer)
Dim i As Byte
Dim CantPeticiones As Byte
With UserList(UserIndex)
    '#######################################
    'AMIGOS CONFIRMADOS
    '#######################################
    For i = 1 To .flags.Amigos.CantFriend
        .flags.Amigos.Amigos(i) = SQLConsulta(.IndexPJ, "amigos", "Friend" & i)
        .flags.Amigos.Confirmed(i) = SQLConsulta(.IndexPJ, "amigos", "Confirmed" & i)
            'Actualizamos la lista de amigos
        Call WriteEnviarAmigo(UserIndex, SQLConsulta(UserList(UserIndex).IndexPJ, "amigos", "Friend" & i), True)
    Next i
    
    '######################################
    'PETICIONES
    '######################################
    CantPeticiones = SQLConsulta(.IndexPJ, "amigos", "CantPeticiones")
    If CantPeticiones > 0 Then
        For i = 1 To CantPeticiones
            Call WriteEnviarAmigo(UserIndex, SQLConsulta(.IndexPJ, "amigos", "Peticion" & i), False)
        Next i
    End If
End With
End Sub

Public Sub CancelarPeticion(ByVal UserIndex As Integer, ByVal TName As String, ByVal Slot As Byte)
Dim CantPeticiones As Byte
Dim Transpaso As String
Dim PJIndexFriend As Integer
Dim CantFriends As Byte
Dim i As Byte
Dim tUser As Integer
'Si llego hasta aqui es por que es una peticion que queremos eliminar
    With UserList(UserIndex)
    
        'Miramos cuantas peticiones tiene en total
        Call SQLConsulta(.IndexPJ, "amigos", "CantPeticiones")
        
        If CantPeticiones > 1 Then
            'Copio el ultimo pj de la lista
            Transpaso = SQLConsulta(.IndexPJ, "amigos", "Friend" & CantPeticiones)
                
            'Paso al ultimo pj al slot libre que dejo el que borramos
            Call SQLSave(.IndexPJ, "amigos", "Peticion" & Slot, "'" & Transpaso & "'")
            
            'Borramos la ultima peticion
            Call SQLSave(.IndexPJ, "amigos", "Peticion" & CantPeticiones, "''")
            
            'Le restamos 1 a la cantidad de peticiones total
            Call SQLSave(.IndexPJ, "amigos", "CantPeticiones", CantPeticiones - 1)
        Else
            'Borramos la peticion
            Call SQLSave(.IndexPJ, "amigos", "Peticion" & Slot, "''")
            'Le restamos 1 a la cantidad de peticiones total
            Call SQLSave(.IndexPJ, "amigos", "CantPeticiones", 0)
        End If
        
        Call WriteConsoleMsg(UserIndex, "Rechazaste la peticion de amistad de " & TName & ".", FontTypeNames.FONTTYPE_INFO)
        Call WriteActualizarBorradoFriend(UserIndex, Slot + .flags.Amigos.CantFriend)
        
        '########################################
        'Parte del amigo
        '########################################
        'buscamos el index de nuestor amigo
        PJIndexFriend = NameToIndex(TName)
        'Comprobamos su cantidad de amigos
        CantFriends = SQLConsulta(PJIndexFriend, "amigos", "TotalFriends")
            
        tUser = NameIndex(TName)
        
        'Buscamos el slot donde nos encontramos en su lista
        For i = 1 To PJIndexFriend
            If UCase(SQLConsulta(PJIndexFriend, "amigos", "Friend" & i)) = UCase(TName) Then Slot = i
        Next i
            
        If CantFriends > 1 Then
            'Copio el ultimo pj de la lista
            Transpaso = SQLConsulta(PJIndexFriend, "amigos", "Friend" & CantFriends)
                
            'Paso al ultimo pj al slot libre que dejo el que borramos
            Call SQLSave(PJIndexFriend, "amigos", "Friend" & Slot, "'" & Transpaso & "'")
            
            'Borramos la ultima peticion
            Call SQLSave(PJIndexFriend, "amigos", "Friend" & CantFriends, "''")
            
            'Resto 1 a la cantidad de amigos
            Call SQLSave(PJIndexFriend, "amigos", "TotalFriends", CantFriends - 1)
            
            'Si el antiguo slot estaba confirmado lo dejamos como estaba en el nuevo
            If SQLConsulta(PJIndexFriend, "amigos", "Friend" & CantFriends) = 1 Then
                Call SQLSave(PJIndexFriend, "amigos", "Confirmed" & Slot, 1)
            Else
                Call SQLSave(PJIndexFriend, "amigos", "Confirmed" & Slot, 0)
            End If
            
            'El ultimo slot tenemos que limpiarlo
            Call SQLSave(PJIndexFriend, "amigos", "Confirmed" & CantFriends, 0)
        Else
            'Borramos el pj
            Call SQLSave(PJIndexFriend, "amigos", "Friend" & Slot, "''")
            
            'Por si acaso
            Call SQLSave(PJIndexFriend, "amigos", "Confirmed" & Slot, 0)
            
            'El ultimo slot tenemos que limpiarlo
            Call SQLSave(PJIndexFriend, "amigos", "TotalFriends", 0)
        End If
            
        '¿Esta online?
        If Not tUser <= 0 Then
            'Si el usuario esta online le decimos que le estan enviando una solicitud de amistad
            UserList(tUser).flags.Amigos.CantFriend = UserList(tUser).flags.Amigos.CantFriend - 1
            Call WriteActualizarBorradoFriend(tUser, Slot)
            Call WriteConsoleMsg(tUser, .name & " rechazo tu solicitud de amistad.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    End With

End Sub

Public Sub EliminarPeticion(ByVal UserIndex As Integer, ByVal TName As String, ByVal Slot As Byte)
Dim CantPeticiones As Byte
Dim Transpaso As String
Dim PJIndexFriend As Integer
Dim i As Byte
Dim tUser As Integer
'Si llego hasta aqui es por que es una peticion que queremos eliminar
    With UserList(UserIndex)
            
        If .flags.Amigos.CantFriend > 1 Then
            'Copio el ultimo pj de la lista
            Transpaso = SQLConsulta(.IndexPJ, "amigos", "Friend" & .flags.Amigos.CantFriend)
                
            'Paso al ultimo pj al slot libre que dejo el que borramos
            Call SQLSave(.IndexPJ, "amigos", "Friend" & Slot, "'" & Transpaso & "'")
            
            'Borramos la ultima peticion
            Call SQLSave(.IndexPJ, "amigos", "Friend" & .flags.Amigos.CantFriend, "''")
            
            'Resto 1 a la cantidad de amigos
            Call SQLSave(.IndexPJ, "amigos", "TotalFriends", .flags.Amigos.CantFriend - 1)
            
            'Si el antiguo slot estaba confirmado lo dejamos como estaba en el nuevo
            If SQLConsulta(.IndexPJ, "amigos", "Friend" & .flags.Amigos.CantFriend) = 1 Then
                Call SQLSave(.IndexPJ, "amigos", "Confirmed" & Slot, 1)
            Else
                Call SQLSave(.IndexPJ, "amigos", "Confirmed" & Slot, 0)
            End If
            
            'El ultimo slot tenemos que limpiarlo
            Call SQLSave(.IndexPJ, "amigos", "Confirmed" & .flags.Amigos.CantFriend, 0)
        Else
            'Borramos el pj
            Call SQLSave(.IndexPJ, "amigos", "Friend" & Slot, "''")
            
            'Por si acaso
            Call SQLSave(.IndexPJ, "amigos", "Confirmed" & Slot, 0)
            
            'El ultimo slot tenemos que limpiarlo
            Call SQLSave(.IndexPJ, "amigos", "TotalFriends", 0)
        End If
        
        Call WriteConsoleMsg(UserIndex, "Cancelaste la peticion de amistad de " & TName & ".", FontTypeNames.FONTTYPE_INFO)
        Call WriteActualizarBorradoFriend(UserIndex, Slot)
        
        '########################################
        'Parte del amigo
        '########################################
        'buscamos el index de nuestor amigo
        PJIndexFriend = NameToIndex(TName)
        'Comprobamos su cantidad de amigos
        CantPeticiones = SQLConsulta(PJIndexFriend, "amigos", "CantPeticiones")
            
        tUser = NameIndex(TName)
        
        'Buscamos el slot donde nos encontramos en su lista
        For i = 1 To PJIndexFriend
            If UCase(SQLConsulta(PJIndexFriend, "amigos", "Peticion" & i)) = UCase(TName) Then Slot = i
        Next i
            
        If CantPeticiones > 1 Then
            'Copio el ultimo pj de la lista
            Transpaso = SQLConsulta(PJIndexFriend, "amigos", "Friend" & CantPeticiones)
                
            'Paso al ultimo pj al slot libre que dejo el que borramos
            Call SQLSave(PJIndexFriend, "amigos", "Peticion" & Slot, "'" & Transpaso & "'")
            
            'Borramos la ultima peticion
            Call SQLSave(PJIndexFriend, "amigos", "Peticion" & CantPeticiones, "''")
            
            'Resto 1 a la cantidad de amigos
            Call SQLSave(PJIndexFriend, "amigos", "CantPeticiones", CantPeticiones - 1)
        Else
            'Borramos el pj
            Call SQLSave(PJIndexFriend, "amigos", "Peticion" & Slot, "''")
            
            'El ultimo slot tenemos que limpiarlo
            Call SQLSave(PJIndexFriend, "amigos", "CantPeticiones", 0)
        End If
            
        '¿Esta online?
        If Not tUser <= 0 Then
            'Si el usuario esta online le decimos que le estan enviando una solicitud de amistad
            UserList(tUser).flags.Amigos.CantFriend = UserList(tUser).flags.Amigos.CantFriend - 1
            Call WriteActualizarBorradoFriend(tUser, Slot + CantPeticiones)
            Call WriteConsoleMsg(tUser, .name & " cancelo la solicitud de amistad que te envio.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    End With

End Sub
'****************************************
'//SISTEMA DE AMIGOS
'****************************************
