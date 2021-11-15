Attribute VB_Name = "Mod_Cuentas"
Option Explicit

Public Type pjs
    NamePJ As String
    LvlPJ As Byte
    ClasePJ As eClass
End Type

Public Type Acc
    ID As Integer
    Name As String
    Pass As String
    
    CantPjs As Byte
    PJ(1 To 8) As pjs
    
End Type
Public Cuenta As Acc

Private Function VerificarCuenta(ByVal UserIndex As Integer, ByVal email As String, ByVal Pass As String, ByVal SerialHD As Long) As Boolean
'*************************************************************************************
'18/03/2016
'Lorwik
'Comprobamos que todos los datos de la cuenta esten correctos y que no esta baneada.
'*************************************************************************************

    '02/11/2015 Irongete: Cargo los datos de la cuenta.
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id,password,mail,bloqueada FROM cuenta WHERE mail = '" & email & "' AND password = '" & Pass & "'")
    
    With Cuenta

        '02/11/2015 Irongete: Compruebo si existe la cuenta.
        If RS.RecordCount = 0 Then
            Call WriteErrorMsg(UserIndex, "La cuenta no existe o la contraseña es incorrecta.")
            VerificarCuenta = False
            Exit Function
        End If
        
        If Not CInt(RS!ID) > 0 Then
            Call WriteErrorMsg(UserIndex, "La cuenta no existe o la contraseña es incorrecta.")
            VerificarCuenta = False
            Exit Function
        End If
        
        '02/11/2015 Irongete: Compruebo si la contraseña es correcta.
        If Not Pass = CStr(RS!Password) Then
            Call WriteErrorMsg(UserIndex, "La cuenta no existe o la contraseña es incorrecta.")
            VerificarCuenta = False
            Exit Function
        End If
        
        Set RS2 = New ADODB.Recordset
        Set RS2 = SQL.Execute("SELECT id FROM personaje WHERE id_cuenta = '" & RS!ID & "' AND logged = '1'")
        
        If RS2.RecordCount > 0 Then
            Call WriteErrorMsg(UserIndex, "Esta cuenta ya está logeada.")
            VerificarCuenta = False
            Exit Function
        End If
        
        '02/11/2015 Irongete: Compruebo si la cuenta está bloqueada.
        If CInt(RS!bloqueada) = 1 Then
            Call WriteErrorMsg(UserIndex, "Esta cuenta está bloqueada. Dirígete a www.aodrag.es/soporte para más información")
            VerificarCuenta = False
            Exit Function
        End If
        
        VerificarCuenta = True
        
    End With
End Function

Private Function IDAcc(ByVal email As String) As Integer
'***************************************************
'18/03/2016
'Lorwik
'Obtenemos la id de una cuenta desde su email
'***************************************************
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id FROM cuenta WHERE mail = '" & email & "'")
    
    IDAcc = RS!ID
    RS.Close
End Function

Public Sub ConectarCuenta(ByVal UserIndex As Integer, ByVal email As String, ByVal Pass As String, ByVal SerialHD As Long)
On Error Resume Next

    'Comprobamos que todo este correcto.
    If Not VerificarCuenta(UserIndex, email, Pass, SerialHD) Then Exit Sub
        
    UserList(UserIndex).CuentaId = IDAcc(email)
        
    '02/11/2015 Irongete: Enviar la lista de personajes
    Dim i As Integer
        
    With Cuenta
        
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT nombre FROM personaje WHERE id_cuenta = '" & UserList(UserIndex).CuentaId & "' AND borrado = '0'")
               
        .CantPjs = RS.RecordCount

        If Not .CantPjs = 0 Then
            i = 1
            While Not RS.EOF
                '18/03/2016 Lorwik: Si esta borrado pasamos de el.
                .PJ(i).NamePJ = RS!Nombre
                i = i + 1
                RS.MoveNext
            Wend
            
            Call WriteLimpiarACC(UserIndex)
            
            For i = 1 To .CantPjs
                Call EnviarCuenta(UserIndex, .PJ(i).NamePJ, i, .CantPjs, "1")
            Next
            
        Else
            Call EnviarCuenta(UserIndex, "", 0, 0, "1")
        End If

    End With
    
    RS.Close
    
    Debug.Print "Conectar cuenta: " & UserIndex
End Sub

Public Sub BorrarPersonaje(ByVal UserIndex As Integer, ByVal email As String, ByVal Pass As String, ByVal SerialHD As Long, ByVal NamePJ As String)
'***************************************************************************
'18/03/2016
'Lorwik
'Borramos un personaje.
'NOTAS:
'El Error 666 nos indicara que posiblemente se quiere borrar un PJ ajeno.
'***************************************************************************

Dim IDCuenta As Integer 'La ID de cuenta que obtenemos desde la cuenta.
Dim IDAccPJ As Integer 'La ID de cuenta que obtenemos desde el PJ.
Dim IDPJ As Integer

    If NamePJ = "" Then
        Call WriteErrorMsg(UserIndex, "Personaje invalido.")
        Exit Sub
    End If

    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("Select logged,borrado FROM personaje WHERE nombre = '" & NamePJ & "'")
    
    '¿El personaje que intenta borrar se encuentra conectado?
    If CByte(RS!logged) = 1 Then
        Call WriteErrorMsg(UserIndex, "El personaje que intentas borrar se encuentra conectado.")
        Exit Sub
    End If
    
    '¿El personaje que intenta borrar ya esta borrado?
    If CByte(RS!borrado) = 1 Then
        Call WriteErrorMsg(UserIndex, "El personaje que intentas borrar no existe.")
        Exit Sub
    End If
    
    'Comprobamos que todo este correcto.
    If Not VerificarCuenta(UserIndex, email, Pass, SerialHD) Then Exit Sub
    
    'Obtenemos su ID de cuenta
    IDCuenta = IDAcc(email)
    
    'Comprobamos que el ID Cuenta del PJ coincida con el de la Cuenta que nos facilito.
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id_cuenta,id FROM personaje WHERE nombre = '" & NamePJ & "'")
    
    IDAccPJ = RS!id_Cuenta
    
    If Not IDCuenta = IDAccPJ Then
        Call WriteErrorMsg(UserIndex, "ERROR: 666 - Ha ocurrido un error al borrar su PJ.")
        Exit Sub
    End If
    
    IDPJ = RS!ID
    
    'Comprobamos que no es fundador o lider de ningun clan.
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id FROM clan WHERE lider = " & IDPJ)
    
    If RS.RecordCount > 0 Then
        Call WriteErrorMsg(UserIndex, "Este personaje es lider de un clan, debes de disolver o abandonar el clan antes de borrarlo.")
        Exit Sub
    End If
    
    'Comprobamos que no es lider de ninguna party.
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id FROM party WHERE lider = " & IDPJ)
    
    If RS.RecordCount > 0 Then
        Call WriteErrorMsg(UserIndex, "Este personaje es lider de una clana, abandonar la party antes de borrarlo.")
        Exit Sub
    End If
    
    'Comprobamos que el ID Cuenta del PJ coincida con el de la Cuenta que nos facilito.
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("UPDATE personaje SET borrado = '1' WHERE nombre = '" & NamePJ & "'")
    
    Call ConectarCuenta(UserIndex, email, Pass, SerialHD)
End Sub
