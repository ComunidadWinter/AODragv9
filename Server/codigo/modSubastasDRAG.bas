Attribute VB_Name = "modSubastasDRAG"
Option Explicit

Public Sub IniciarSubasta(ByVal UserIndex As Integer)
'18/11/2018
'Lorwik
'Envia la orden al cliente de abrir la ventana de subastas

On Error GoTo Errhandler

    'Actualizamos el dinero
    Call WriteUpdateUserStats(UserIndex)
    'Mostramos la ventana
    Call WriteSubInit(UserIndex)
    UserList(UserIndex).flags.Comerciando = True
    
Errhandler:

End Sub

Public Sub NuevaSubasta(ByVal UserIndex As Integer, ByVal ItemSelect As Byte, ByVal Precio As Long, ByVal Tiempo As Byte, ByVal Cantidad As Integer)
'18/11/2018
'Lorwik
'Publicamos una nueva subasta
Dim ObjIndex As Integer

    If Precio < 0 And Precio > 999999999 Then Exit Sub
    
    If ItemSelect > MAX_INVENTORY_SLOTS Then Exit Sub
    
    If Tiempo < 0 Or Tiempo > 3 Then Exit Sub
    
    ObjIndex = UserList(UserIndex).Invent.Object(ItemSelect).ObjIndex
    
    '¿Cumple los requisitos para subastar ese objeto?
    If PuedeSubastar(UserIndex, ObjIndex) = False Then Exit Sub
    
    Select Case Tiempo
        Case 1  '6 horas
            Tiempo = 6
        Case 2  '12 horas
            Tiempo = 12
        Case 3  '24 horas
            Tiempo = 24
    End Select

    'ATENCION!!!!!!!!!!!! TENGO QUE VER COMO CONTROLAR EL TIEMPO
    
    Set RS = New ADODB.Recordset
    
    'Registo la nueva subasta
    Set RS = SQL.Execute("INSERT INTO subasta (objeto_id, personaje_id, cantidad, buyout, fecha_creacion) VALUES ('" & ObjIndex _
    & "', '" & UserList(UserIndex).ID & "', '" & Cantidad & "', '" & Precio & "', '" & Ahora() & "')")
    
    Call QuitarUserInvItem(UserIndex, ItemSelect, Cantidad)

End Sub

Private Function PuedeSubastar(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    'Quizas por algun motivo se murio...
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.Muerto)
        PuedeSubastar = False
        Exit Function
    End If
    
    'No podra comerciar con objetos de newbie
    If ObjData(ObjIndex).Newbie = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.NoComerciar)
        PuedeSubastar = False
        Exit Function
    End If
    
    'Tampoco con objetos no comerciables
    If ObjData(ObjIndex).NoComerciable = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.NoComerciar)
        PuedeSubastar = False
        Exit Function
    End If
    
    PuedeSubastar = True
End Function

