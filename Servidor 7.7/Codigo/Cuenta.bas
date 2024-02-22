Attribute VB_Name = "Cuenta"
Type Acc
    mail As String
    NumPjs As Integer
    Pj() As String
    Logged As Boolean
    passwd As String
    Llave As Integer
    Naci As Byte
    Cajas As Byte
End Type

Public AccPath As String
Public Cuentas() As Acc

Function LlaveCuenta(UserIndex As Integer) As Integer
   On Error GoTo fallo
   LlaveCuenta = Cuentas(UserIndex).Llave
Exit Function
fallo:
Call LogError("LLAVECUENTA" & Err.number & " D: " & Err.Description)
End Function

Sub IniciaCuentas()
  On Error GoTo fallo
    AccPath = App.Path & "\Accounts\"
Exit Sub
fallo:
Call LogError("INICIACUENTA" & Err.number & " D: " & Err.Description)

End Sub
Sub ListarClaves(X As Integer)
On Error Resume Next

Dim Fn As String
Dim cad$
Dim n As Integer, k As Integer
'Dim tindex As Integer
'tindex = NameIndex("AoDraGBoT")
'If tindex <= 0 Then Exit Sub


Fn = App.Path & "\logs\RecuperaClaves.log"
Call SendData(ToIndex, X, 0, "|| Listando Recuperación de Claves..." & "´" & FontTypeNames.FONTTYPE_TALK)

If FileExist(Fn, vbArchive) Then

n = FreeFile
Open Fn For Input Shared As #n
    Do While Not EOF(n)
        k = k + 1
        Input #n, cad$
        
'Call SendData(ToIndex, tindex, 0, "|| " & k & "- " & cad$ & FONTTYPENAMES.FONTTYPE_TALK)
Call SendData(ToIndex, X, 0, "|| " & k & "- " & cad$ & "´" & FontTypeNames.FONTTYPE_TALK)
Loop
    Close #n
Else
Call SendData(ToIndex, X, 0, "|| No hay claves para mandar" & "´" & FontTypeNames.FONTTYPE_TALK)
End If
Call SendData(ToIndex, X, 0, "|| Fin del Listado" & "´" & FontTypeNames.FONTTYPE_TALK)

End Sub
Sub BorrarClaves(X As Integer)
Dim Fn As String
Fn = App.Path & "\logs\RecuperaClaves.log"

If FileExist(Fn, vbArchive) Then
Kill (Fn)
Call SendData(ToIndex, X, 0, "|| Fichero de Claves borrado" & "´" & FontTypeNames.FONTTYPE_TALK)
Else
Call SendData(ToIndex, X, 0, "|| Fichero de Claves no existe" & "´" & FontTypeNames.FONTTYPE_TALK)
End If
End Sub
Sub MandaPersonajes(UserIndex As Integer)
  On Error GoTo fallo
  Dim cc As String
  Dim Lla As Integer
  Dim LLa2 As String

    'Call SendData2(ToIndex, Userindex, 0, 75)
    If Cuentas(UserIndex).NumPjs <> 0 Then
        For X = 1 To Cuentas(UserIndex).NumPjs
        cc = cc & Cuentas(UserIndex).Pj(X) & ","
        Next X
    End If
        
    'pluto:6.0A
    Lla = BuscaLlave(Cuentas(UserIndex).Llave)
 
    If Lla = 0 Then
    Cuentas(UserIndex).Llave = 0
    LLa2 = ""
    Else
    LLa2 = ObjData(Lla).Name
    End If
    
    Call SendData2(ToIndex, UserIndex, 0, 73, Cuentas(UserIndex).NumPjs & "," & cc & LLa2)

   ' Call SendData2(ToIndex, Userindex, 0, 74)
Exit Sub
fallo:
Call LogError("MANDARPERSONAJES" & Err.number & " D: " & Err.Description)


End Sub

Function CuentaBaneada(mail As String) As Boolean
 On Error GoTo fallo
    If val(GetVar(AccPath & mail & ".acc", "DATOS", "BAN")) = 1 Then
        CuentaBaneada = True
    Else
        CuentaBaneada = False
    End If
Exit Function
fallo:
Call LogError("CUENTABANEADA" & Err.number & " D: " & Err.Description)


End Function

Function CuentaExiste(mail As String) As Boolean
 On Error GoTo fallo
    If FileExist(AccPath & mail & ".acc", vbArchive) Then
        CuentaExiste = True
    Else
        CuentaExiste = False
    End If
Exit Function
fallo:
Call LogError("CUENTAEXISTE" & Err.number & " D: " & Err.Description)

End Function

Function BuscaLlave(Clave As Integer) As Integer
  On Error GoTo fallo
    Dim X As Integer
    For X = 1 To NumObjDatas
        If ObjData(X).OBJType = OBJTYPE_LLAVES And ObjData(X).Clave = Clave Then
            BuscaLlave = X
            Exit Function
        End If
    Next X
    BuscaLlave = 0
Exit Function
fallo:
Call LogError("BUSCALLAVE" & Err.number & " D: " & Err.Description)

End Function

Sub ConectaCuenta(UserIndex As Integer, mail As String, pass As String)
  On Error GoTo fallo
    Dim X As Integer
    Dim a As String
    Dim Lla As String
    Dim Filex As String

    
    
    
a = GetVar(AccPath & mail & ".acc", "DATOS", "PASSWORD")
If GetVar(AccPath & mail & ".acc", "DATOS", "PASSWORD") <> pass Then
        'Call SendData2(ToIndex, UserIndex, 0, 43, "Contraseña incorrecta.")
        Call SendData2(ToIndex, UserIndex, 0, 95)
'pluto:6.8
'UserList(UserIndex).flags.Intentos = UserList(UserIndex).flags.Intentos + 1
Exit Sub
End If

'pluto:7.0 contraseña correcta cargamos datos
Filex = AccPath & mail & ".acc"
Dim Leer As New clsLeerInis
Leer.Abrir Filex
Cuentas(UserIndex).NumPjs = val(Leer.DarValor("DATOS", "NumPjs"))




    Cuentas(UserIndex).Logged = True
    Cuentas(UserIndex).mail = mail
    'Cuentas(UserIndex).NumPjs = val(GetVar(AccPath & mail & ".acc", "DATOS", "NumPjs"))
    Cuentas(UserIndex).NumPjs = val(Leer.DarValor("DATOS", "NumPjs"))
    Cuentas(UserIndex).passwd = Leer.DarValor("DATOS", "Password")
    'Cuentas(UserIndex).passwd = GetVar(AccPath & mail & ".acc", "DATOS", "Password")
    'Cuentas(UserIndex).Llave = val(GetVar(AccPath & mail & ".acc", "DATOS", "Llave"))
    Cuentas(UserIndex).Llave = val(Leer.DarValor("DATOS", "Llave"))
    'pluto:7.0
    
    Cuentas(UserIndex).Cajas = val(Leer.DarValor("DATOS", "Cajas"))
    If Cuentas(UserIndex).Cajas < 2 Then Cuentas(UserIndex).Cajas = 2
    
    'pluto:6.0A
    'Cuentas(UserIndex).Naci = val(GetVar(AccPath & mail & ".acc", "DATOS", "Naci"))
    
    
    If Cuentas(UserIndex).NumPjs <> 0 Then
        ReDim Cuentas(UserIndex).Pj(1 To Cuentas(UserIndex).NumPjs)
        For X = 1 To Cuentas(UserIndex).NumPjs
            'Cuentas(UserIndex).Pj(X) = GetVar(AccPath & mail & ".acc", "PERSONAJES", "PJ" & X)
            Cuentas(UserIndex).Pj(X) = Leer.DarValor("PERSONAJES", "PJ" & X)
        Next X
    End If
    
    'pluto:7.0 cargamos Cajas
    Call ResetUserBanco(UserIndex)
    Dim Ln2 As String
    For n = 1 To Cuentas(UserIndex).Cajas
        For X = 1 To 20
        Ln2 = Leer.DarValor("CAJA" & n, "Obj" & X)
        
        UserList(UserIndex).BancoInvent(n).Object(X).ObjIndex = val(ReadField(1, Ln2, 45))
        UserList(UserIndex).BancoInvent(n).Object(X).Amount = val(ReadField(2, Ln2, 45))

        Next X
    Next n
    
    
    
    
    
    'pluto:2.10
If Cuentas(UserIndex).mail = "aodragbot@aodrag.com" Then
Call ConnectUser(UserIndex, "AoDraGBoT", "hola", "bot", "MacBot")
Exit Sub
End If
 If Cuentas(UserIndex).mail = "aodragbot2@aodrag.com" Then
Call ConnectUser(UserIndex, "AoDraGBoT2", "hola", "bot", "MacBot")
Exit Sub
End If
    Call MandaPersonajes(UserIndex)



     'pluto:6.0A
    'Lla = ObjData(BuscaLlave(Cuentas(UserIndex).Llave)).Name
    'If Lla = "" Then Cuentas(UserIndex).Llave = 0
    
    'If Cuentas(UserIndex).Llave <> 0 Then Call SendData2(ToIndex, UserIndex, 0, 76, Lla) Else Call SendData2(ToIndex, UserIndex, 0, 76)
    Call ResetUserSlot(UserIndex)

Exit Sub
fallo:
Call LogError("CONECTACUENTA" & Err.number & " D: " & Err.Description)


End Sub

Function EstaUsandoCuenta(mail As String) As Boolean
   On Error GoTo fallo
'pluto:2.12
EstaUsandoCuenta = False
Dim uf As Integer
uf = DameIndexCuenta(mail)
If uf <> 0 Then
EstaUsandoCuenta = True
If EstaUsandoCuenta = True And UserList(uf).flags.UserLogged = False Then
Call DesconectaCuenta(uf)
Call CloseSocket(uf)
End If
End If

Exit Function
fallo:
Call LogError("ESTAUSANDOCUENTA" & Err.number & " D: " & Err.Description)

End Function

Function DameIndexCuenta(mail As String) As Integer
 On Error GoTo fallo
    Dim X As Integer
    For X = 1 To MaxUsers
        'pluto:2.8.0
        If UCase$(Cuentas(X).mail) = UCase$(mail) And Cuentas(X).Logged Then
            DameIndexCuenta = X
'pluto:2.12
'If UserList(X).Flags.UserLogged = False Then Call DesconectaCuenta(X)
'------------------------

            Exit Function
        End If
    Next X
    DameIndexCuenta = 0
Exit Function
fallo:
Call LogError("DAMEINDEXCUENTA" & Err.number & " D: " & Err.Description)

End Function


Sub DesconectaCuenta(UserIndex As Integer)
   On Error GoTo fallo
    Dim X As Integer
    Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "DATOS", "Password", Cuentas(UserIndex).passwd)
    Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "DATOS", "NumPjs", CStr(Cuentas(UserIndex).NumPjs))
    Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "DATOS", "Llave", CStr(Cuentas(UserIndex).Llave))
    Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "DATOS", "Naci", CStr(Cuentas(UserIndex).Naci))
'pluto:7.0
Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "DATOS", "Cajas", CStr(Cuentas(UserIndex).Cajas))


    For X = 1 To Cuentas(UserIndex).NumPjs
        Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "PERSONAJES", "PJ" & X, Cuentas(UserIndex).Pj(X))
        Cuentas(UserIndex).Pj(X) = ""
    Next
    'pluto:2.8.0
    If GetVar(AccPath & Cuentas(UserIndex).mail & ".acc", "PERSONAJES", "PJ" & Cuentas(UserIndex).NumPjs + 1) > "" Then
    Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "PERSONAJES", "PJ" & Cuentas(UserIndex).NumPjs + 1, "")
    End If
    
    
    'pluto:7.0 Grabamos cajas------------------------------
    
    For n = 1 To Cuentas(UserIndex).Cajas
        For X = 1 To 20
        Call WriteVar(AccPath & Cuentas(UserIndex).mail & ".acc", "CAJA" & n, "Obj" & X, UserList(UserIndex).BancoInvent(n).Object(X).ObjIndex & "-" & UserList(UserIndex).BancoInvent(n).Object(X).Amount)
        Next X
    Next n
    '---------------------------------------------------
    
    
    Cuentas(UserIndex).NumPjs = 0
    Cuentas(UserIndex).mail = ""
    Cuentas(UserIndex).passwd = ""
    Cuentas(UserIndex).Logged = False
    Cuentas(UserIndex).Llave = 0
    Cuentas(UserIndex).Naci = 0
    'pluto:7.0
    Cuentas(UserIndex).Cajas = 0
    
UserList(UserIndex).MacPluto2 = ""
UserList(UserIndex).Paquete = 0
UserList(UserIndex).MacClave = 0
Call ResetUserBanco(UserIndex)
Exit Sub
fallo:
Call LogError("DESCONECTA CUENTA" & Err.number & " D: " & Err.Description)

End Sub

Public Function CheckMailString(ByRef sString As String) As Boolean
        On Error GoTo fallo
        Dim lPos  As Long, lX    As Long
        Dim iAsc  As Integer
    
        '1er test: Busca un simbolo @
        lPos = InStr(sString, "@")
        If (lPos <> 0) Then
            '2do test: Busca un simbolo . después de @ + 1
            If Not (IIf((InStr(lPos, sString, ".", vbBinaryCompare) > (lPos + 1)), True, False)) Then _
                Exit Function
    
            '3er test: Valída el ultimo caracter
            If Not (CMSValidateChar_(Asc(Right(sString, 1)))) Then _
                Exit Function
    
            '4to test: Recorre todos los caracteres y los valída
            For lX = 0 To Len(sString) - 1 'el ultimo no porque ya lo probamos
                If Not (lX = (lPos - 1)) Then
                    iAsc = Asc(Mid(sString, (lX + 1), 1))
                    If Not (iAsc = 46 And lX > (lPos - 1)) Then _
                        If Not CMSValidateChar_(iAsc) Then _
                            Exit Function
                End If
            Next lX
    
            'Finale
            CheckMailString = True
        End If
    
Exit Function
fallo:
Call LogError("CHECKMAILSTRING" & Err.number & " D: " & Err.Description)

    End Function
    
Private Function CMSValidateChar_(ByRef iAsc As Integer) As Boolean
On Error GoTo fallo

CMSValidateChar_ = IIf( _
                    (iAsc >= 45 And iAsc <= 57) Or _
                    (iAsc >= 65 And iAsc <= 90) Or _
                    (iAsc >= 97 And iAsc <= 122) Or _
                    (iAsc = 95), True, False)

Exit Function
fallo:
Call LogError("CMSVALIDATECHAR" & Err.number & " D: " & Err.Description)


End Function

