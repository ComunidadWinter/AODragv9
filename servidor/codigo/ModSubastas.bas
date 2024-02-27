Attribute VB_Name = "ModSubastas"
'**************************************************************
'ModSubastas
'Ultima modificación: 07/03/2016
'Autor: Lorwik
'Descripción: Sistema de subastas por consola en tiempo real.
'**************************************************************

Option Explicit

'Item y cantidad que se esta subastando
Private ItemSub As Integer
Private CantidadSub As Integer

'Precio inicial para comenzar una subasta
Private Const PrecioInicialMinimo As Integer = 1
'Porcentaje de la comision dela subasta
Private Const ComisionSubasta As Byte = 3

'¿Hay una subasta en curso?
Public SubastaEnCurso As Boolean
'Precio de la oferta actual
Private OfertaActual As Long

Private UserIDSuba As Integer
Private UserIDPuja As Integer

'Tiempo que dura la subasta
Public TiempoSubasta As Byte

Public Function PuedeSubastar(ByVal UserIndex As Integer) As Boolean
'********************************************************************************************************
'Fecha: 07/08/2016
'Autor: Lorwik
'Descripción: Comprueba si es posible iniciar una nueva subasta.
'********************************************************************************************************
    If SubastaEnCurso Then
        Call WriteConsoleMsg(UserIndex, "¡Ya hay una subasta en curso, espera a que finalice!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    '¿Esta el user muerto? Si es asi no puede subastar
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.Muerto)
        Exit Function
    End If
    
    PuedeSubastar = True
End Function

Public Sub IniciarSubasta(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer, ByVal PrecioInicial As Long)
'********************************************************************************************************
'Fecha: 07/08/2016
'Autor: Lorwik
'Descripción: Si no hay ninguna subasta en curso, iniciamos una nueva.
'********************************************************************************************************
    Dim query As String
    
    'No se pueden subastar items que no se pueden comerciar ni de newbie
    If ObjData(Item).NoComerciable = 1 Or ObjData(Item).Newbie = 1 Then
        Call WriteConsoleMsg(UserIndex, "¡No puedes subastar ese objeto!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    'Precio inicial minimo
    If PrecioInicial < PrecioInicialMinimo Then
        Call WriteConsoleMsg(UserIndex, "¡El precio inicial minimo debe ser de " & PrecioInicialMinimo & " moneda de oro!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Iniciamos subasta
    Call WriteConsoleMsg(UserIndex, "La subasta ha dado comienzo. Esta subasta tiene una comisión de " & ComisionSubasta & "%.", FontTypeNames.FONTTYPE_INFO)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " esta subastando " & Cantidad & " " & ObjData(Item).Name & " con un precio inicial de " & PrecioInicial & " monedas de oro. Escribe /PUJAR (cantidad) para hacer una oferta.", FontTypeNames.FONTTYPE_INFOBOLD))
    
    SubastaEnCurso = True
    UserIDSuba = UserList(UserIndex).ID
    ItemSub = Item
    CantidadSub = Cantidad
    OfertaActual = PrecioInicial
    TiempoSubasta = 3
    
    Set RS = New ADODB.Recordset
    query = "UPDATE subasta SET "
    query = query & "UserIDSuba = '" & UserIDSuba & "', "
    query = query & "UserIDPuja = '0', "
    query = query & "ObjCant = '" & ItemSub & "-" & CantidadSub & "', "
    query = query & "TiempoSuba = '" & TiempoSubasta & "', "
    query = query & "Subastaenc = '1', "
    query = query & "OfertaA = '" & OfertaActual & "'"
    Set RS = SQL.Execute(query)
    
    Call QuitarObjetos(Item, Cantidad, UserIndex)
End Sub

Public Sub PujarSubasta(ByVal UserIndex As Integer, ByVal Oferta As Long)
'********************************************************************************************************
'Fecha: 07/08/2016
'Autor: Lorwik
'Descripción: Si hay una subasta en curso, hace una puja.
'********************************************************************************************************
    Set RS = New ADODB.Recordset
    Dim UserIndexPujador As Integer
    Dim UserNamePujador As String
    Dim UserNameSubastador As String
    
    'Información de pujador
    If UserIDPuja > 0 Then
        UserNamePujador = GetIdPorNombre(UserIDPuja)
        UserIndexPujador = GetPersonajeIndex(UserIDPuja)
    End If
    'Información del subastador
    UserNameSubastador = GetIdPorNombre(UserIDSuba)
    
    '¿Esta el user muerto? Si es asi no puede pujar
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.Muerto)
        Exit Sub
    End If
    
    If SubastaEnCurso = False Then
        Call WriteConsoleMsg(UserIndex, "¡No hay ninguna subasta en curso!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).Name = UserNameSubastador Then
        Call WriteConsoleMsg(UserIndex, "¡no puedes pujar en tu propia subasta!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Oferta < 1 Then
        Call WriteConsoleMsg(UserIndex, "¡Cantidad invalida!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).Stats.GLD < Oferta Then
        Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad dinero.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Oferta < OfertaActual Then
        Call WriteConsoleMsg(UserIndex, "¡Tu oferta es inferior a la puja actual!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Oferta < OfertaActual + 200 Then
        Call WriteConsoleMsg(UserIndex, "¡Debes hacer una oferta minima de " & OfertaActual + 200 & " monedas de oro!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserIndex = UserIndexPujador Then
        Call WriteConsoleMsg(UserIndex, "¡Acabas de pujar!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not UserNamePujador = "" Then
        If Not UserIndexPujador = 0 Then
            'Le devolvemos el oro al pujador anterior
            UserList(UserIndexPujador).Stats.GLD = UserList(UserIndexPujador).Stats.GLD + OfertaActual
            Call WriteUpdateGold(UserIndexPujador)
        Else
            Set RS = SQL.Execute("SELECT gld FROM `personaje` WHERE id = '" & UserIDPuja & "'")
            Set RS = SQL.Execute("UPDATE personaje SET gld = '" & RS!GLD + OfertaActual & "' WHERE id = '" & UserIDPuja & "'")
        End If
    End If
    
    UserIDPuja = UserList(UserIndex).ID
    OfertaActual = Oferta
    Set RS = SQL.Execute("UPDATE subasta SET UserIDPuja = '" & UserIDPuja & "', OfertaA = '" & OfertaActual & "'")
    
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Oferta
    Call WriteUpdateGold(UserIndex)
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha pujado por " & Oferta & " monedas de oro. Escribe /PUJAR (cantidad) para hacer una oferta.", FontTypeNames.FONTTYPE_INFOBOLD))
    Call WriteConsoleMsg(UserIndex, "Acabas de pujar en subasta, mientras seas el pujador mas alto no salgas del personaje ¡o no recibiras el item!", FontTypeNames.FONTTYPE_INFO)
End Sub

Public Sub InfoSubasta(ByVal UserIndex As Integer)
'********************************************************************************************************
'Fecha: 07/08/2016
'Autor: Lorwik
'Descripción: Comando para obtener información de la subasta actual.
'********************************************************************************************************
    Dim UserIndexPujador As Integer
    Dim UserNamePujador As String
    Dim UserNameSubastador As String
    
    If SubastaEnCurso = False Then
        Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta en estos momentos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Información de pujador
    If UserIDPuja > 0 Then
        UserNamePujador = GetIdPorNombre(UserIDPuja)
        UserIndexPujador = GetPersonajeIndex(UserIDPuja)
    End If
    'Información del subastador
    UserNameSubastador = GetIdPorNombre(UserIDSuba)
    'Comando /INFOSUB
    Call WriteConsoleMsg(UserIndex, "Subasta iniciada por: " & UserNameSubastador, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(UserIndex, "Item subastandose: " & ObjData(ItemSub).Name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(UserIndex, "Cantidad: " & CantidadSub, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(UserIndex, "Ultima oferta: " & OfertaActual & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(UserIndex, "Pujador actual: " & UserNamePujador, FontTypeNames.FONTTYPE_INFO)
End Sub

Public Sub FinalizarSubasta()
'********************************************************************************************************
'Fecha: 07/08/2016
'Autor: Lorwik
'Descripción: Finalizamos la subasta y entregamos el item al pujador y el oro al subastador.
'En caso de que nadie pujase, se le entrega el item al subastador.
'********************************************************************************************************
    Dim ObjGanado As Obj

    Dim UserIndexPujador As Integer
    Dim UserNamePujador As String
    Dim UserNameSubastador As String
    Dim UserIndexSubastador As Integer
    
    'Información de pujador
    If UserIDPuja > 0 Then
        UserNamePujador = GetIdPorNombre(UserIDPuja)
        UserIndexPujador = GetPersonajeIndex(UserIDPuja)
    End If
    
    'Información del subastador
    UserNameSubastador = GetIdPorNombre(UserIDSuba)
    UserIndexSubastador = GetPersonajeIndex(UserIDSuba)

    ObjGanado.ObjIndex = ItemSub
    ObjGanado.Amount = CantidadSub
        
    '¿Tuvo ofertas durante la subasta?
    If Not UserIDPuja = 0 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("La subasta finalizó con una puja maxima de " & OfertaActual & " realizada por " & UserNamePujador & ".", FontTypeNames.FONTTYPE_INFOBOLD))
        '¿El subastador esta online?
        If Not UserIndexSubastador = 0 Then
            UserList(UserIndexSubastador).Stats.GLD = UserList(UserIndexSubastador).Stats.GLD + Porcentaje(OfertaActual, ComisionSubasta)
            Call WriteUpdateGold(UserIndexSubastador)
            Call WriteConsoleMsg(UserIndexSubastador, "¡Recibes " & OfertaActual & " monedas de oro de la subasta!", FontTypeNames.FONTTYPE_oro)
            If ComisionSubasta > 0 Then _
            Call WriteConsoleMsg(UserIndexSubastador, "El centro de subastas se quedo con " & Porcentaje(OfertaActual, ComisionSubasta) & " monedas de oro de comisión.", FontTypeNames.FONTTYPE_INFOBOLD)
            UserList(UserIndexSubastador).Stats.GLD = UserList(UserIndexSubastador).Stats.GLD + (Porcentaje(OfertaActual, ComisionSubasta) - OfertaActual)
        Else
            Set RS = SQL.Execute("SELECT gld FROM `personaje` WHERE id = '" & UserIDSuba & "'")
            Dim GLDActual As Long
            GLDActual = RS!GLD
            Set RS = SQL.Execute("UPDATE personaje SET gld = '" & GLDActual + OfertaActual & "' WHERE id = '" & UserIDSuba & "'")
        End If
        
        If Not UserIndexPujador = 0 Then
            If Not MeterItemEnInventario(UserIndexPujador, ObjGanado) Then
                Call TirarItemAlPiso(UserList(UserIndexPujador).Pos, ObjGanado)
            Else
                Call UpdateUserInv(True, UserIndexPujador, 0)
            End If
        End If
    Else
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("La subasta finalizó sin pujas.", FontTypeNames.FONTTYPE_INFOBOLD))
        '¿El subastador esta online?
        If Not UserIndexSubastador = 0 Then
            If Not MeterItemEnInventario(UserIndexSubastador, ObjGanado) Then
                Call TirarItemAlPiso(UserList(UserIndexSubastador).Pos, ObjGanado)
            Else
                Call UpdateUserInv(True, UserIndexSubastador, 0)
            End If
        End If
    End If
    
    Call ResetearSubasta
End Sub

Private Sub ResetearSubasta()
'********************************************************************************************************
'Fecha: 07/08/2016
'Autor: Lorwik
'Descripción: Reseteamos los datos de las subastas.
'********************************************************************************************************
Dim query As String

    SubastaEnCurso = False
    UserIDPuja = 0
    UserIDSuba = 0
    OfertaActual = 0
    ItemSub = 0
    CantidadSub = 0
    
    Set RS = New ADODB.Recordset
    query = "UPDATE subasta SET "
    query = query & "UserIDSuba = '0', "
    query = query & "UserIDPuja = '0', "
    query = query & "ObjCant = '0-0', "
    query = query & "TiempoSuba = '0', "
    query = query & "Subastaenc = '0', "
    query = query & "OfertaA = '0'"
    Set RS = SQL.Execute(query)
    
End Sub

Public Sub ReanudarSubasta()
'********************************************************************************************************
'Fecha: 07/08/2016
'Autor: Lorwik
'Descripción: Cuando iniciamos el server comprobamos si habia alguna subasta en curso cuando cerro.
'********************************************************************************************************
    Dim ItemsCant As String
    
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT Subastaenc FROM `subasta`")
    
    If RS!Subastaenc = 0 Then Exit Sub
    
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT UserIDSuba,UserIDPuja,ObjCant,TiempoSuba,OfertaA FROM `subasta`")
    
    UserIDSuba = CLng(RS!UserIDSuba)
    UserIDPuja = CLng(RS!UserIDPuja)
    ItemsCant = CStr(RS!ObjCant)
    ItemSub = ReadField(1, ItemsCant, Asc("-"))
    CantidadSub = ReadField(2, ItemsCant, Asc("-"))
    OfertaActual = RS!OfertaA
    TiempoSubasta = RS!TiempoSuba
    SubastaEnCurso = True
    
End Sub
