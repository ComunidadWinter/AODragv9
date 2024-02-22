Attribute VB_Name = "BaseDeDatos"


Dim conn As ADODB.Connection
Dim BD As Boolean

'Public Function BDDAddAcount(mail As String) As Boolean
    'If Bd = False Then Exit Function
    'Dim passwd As String
    'Dim BDD As ADODB.Recordset
   ' Set BDD = New ADODB.Recordset
    'If BDD Is Nothing Then GoTo mal
    'BDD.Open "SELECT * FROM cuentas WHERE mail='" & mail & "'", conn
    'If BDD.EOF = True Then
       ' passwd = Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122))
        'BDDAddAcount = True
        'conn.Execute "INSERT INTO cuentas VALUES('" & mail & "','" & passwd & "',0,'" & RandomNumber(1111, 60000) & "',0)"
    'Else
        'BDDAddAcount = False
    'End If
    'BDD.Close
    'Set BDD = Nothing
    'Exit Function
'mal:
'End Function

'Public Function BDDAddRecovery(mail As String) As Boolean
    '''an Error GoTo mal
    'Dim password As String
    'If Bd = False Then Exit Function
    'Dim BDD As ADODB.Recordset
    'Set BDD = New ADODB.Recordset
    'If BDD Is Nothing Then GoTo mal
    'BDD.Open "SELECT * FROM recovery WHERE mail='" & mail & "'", conn
    'If BDD.EOF Then
        'password = Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122)) & Chr$(RandomNumber(97, 122))
        'conn.Execute "INSERT INTO recovery VALUES('" & mail & "','" & password & "','" & RandomNumber(1111, 60000) & "',0,0)"
        'BDDAddRecovery = True
   ' Else
        'BDDAddRecovery = False
    'End If
    'BDD.Close
    'Set BDD = Nothing
    'Exit Function
'mal:
'End Function

'Public Sub BDDCmpCuentas()
    ''''an error GoTo mal
    'Dim file As String
    'If Bd = False Then Exit Sub
   ' Dim BDD As ADODB.Recordset
    'Set BDD = New ADODB.Recordset
    'If BDD Is Nothing Then GoTo mal
    'BDD.Open "SELECT * FROM cuentas WHERE activada=1", conn, adOpenKeyset, adLockOptimistic
    'If Not BDD.EOF Then
        'file = AccPath & BDD!mail & ".acc"
        'If Not CuentaExiste(BDD!mail) Then
            'Call WriteVar(file, "DATOS", "NumPjs", "0")
            'Call WriteVar(file, "DATOS", "Ban", "0")
            'Call WriteVar(file, "DATOS", "Password", MD5String(BDD!pass))
            'Call WriteVar(file, "DATOS", "Llave", "0")
        'End If
        'BDD!activada = 2
        'BDD.Update
   ' End If
    'BDD.Close
    
    'BDD.Open "SELECT * FROM recovery WHERE activada=1", conn, adOpenKeyset, adLockOptimistic
    'If Not BDD.EOF Then
        'file = AccPath & BDD!mail & ".acc"
        'If EstaUsandoCuenta(BDD!mail) Then
            'Call CloseSocket(DameIndexCuenta(BDD!mail))
        'End If
        'If CuentaExiste(BDD!mail) Then
            'Call WriteVar(file, "DATOS", "Password", MD5String(BDD!pass))
        'End If
        'BDD!activada = 2
        'BDD.Update
    'End If
    'BDD.Close

    'Set BDD = Nothing
    'Exit Sub
'mal:
'End Sub

'Public Sub BDDSetGMState(user As String, estado As Integer)
    ''''an error GoTo mal
    'If Bd = False Then Exit Sub
    'Dim BDD As ADODB.Recordset
    'Set BDD = New ADODB.Recordset
    'If BDD Is Nothing Then GoTo mal
'    BDD.Open "SELECT * FROM gms WHERE nombre='" & user & "'", conn, adOpenKeyset, adLockOptimistic
'    If BDD.EOF = False Then
 '       BDD!online = estado
  '      BDD.Update
   ' End If
    'BDD.Close
    'Set BDD = Nothing
    'Exit Sub
'mal:
'End Sub

'Public Function BDDGetHash(user As String) As String
    '''an error GoTo mal
    'If Bd = False Then Exit Function
    'Dim BDD As ADODB.Recordset
    'Set BDD = New ADODB.Recordset
    'If BDD Is Nothing Then GoTo mal
    'BDD.Open "SELECT * FROM gms WHERE nombre='" & user & "'", conn
    'If BDD.EOF = True Then
        'BDDGetHash = ""
    'Else
        'BDDGetHash = BDD!hash
    'End If
    'BDD.Close
    'Set BDD = Nothing
    'Exit Function
'mal:
'End Function
'pluto:2.20
Public Sub BDDSetUsersOnline()
  On Error GoTo mal
    If BD = False Then Exit Sub
    Dim loopc As Integer
    Dim tStr As String
   ' pluto:2.24-------------
If ServerPrimario = 1 Then
    conn.Execute "DELETE FROM online"
    conn.Execute "INSERT INTO online VALUES('" & NumUsers & "','" & ReNumUsers & "','" & AyerReNumUsers & "')"
Else
   conn.Execute "DELETE FROM online2"
   conn.Execute "INSERT INTO online2 VALUES('" & NumUsers & "','" & ReNumUsers & "','" & AyerReNumUsers & "')"
End If
  '--------------------
Exit Sub
mal:
'MsgBox ("vamos mal")
End Sub
'--------------------------------------------------
Public Sub BorraPjBD(nombrecito As String)
If BD = False Then Exit Sub
conn.Execute "DELETE FROM estadis WHERE nombre='" & nombrecito & "'"
End Sub
Public Sub BDDSetCastillos()
 On Error GoTo mal
'Set SQLResult = SQLLink.MyExecute("SELECT * FROM `castillos`")
'castillo1 = SQLResult("norte")
'castillo2 = SQLResult("sur")
'castillo3 = SQLResult("este")
'castillo4 = SQLResult("oeste")
'castillo5 = SQLResult("fortaleza")

  'MsgBox (G190)
'pluto:2.24-------------------
'quitar esto
Exit Sub
If ServerPrimario = 1 Then
    conn.Execute "DELETE FROM castillos"
    conn.Execute "INSERT INTO castillos VALUES('norte','" & castillo1 & "')"
    conn.Execute "INSERT INTO castillos VALUES('sur','" & castillo2 & "')"
    conn.Execute "INSERT INTO castillos VALUES('este','" & castillo3 & "')"
   conn.Execute "INSERT INTO castillos VALUES('oeste','" & castillo4 & "')"
   conn.Execute "INSERT INTO castillos VALUES('fortaleza','" & fortaleza & "')"
Else
conn.Execute "DELETE FROM castillos2"
    conn.Execute "INSERT INTO castillos2 VALUES('norte','" & castillo1 & "')"
    conn.Execute "INSERT INTO castillos2 VALUES('sur','" & castillo2 & "')"
    conn.Execute "INSERT INTO castillos2 VALUES('este','" & castillo3 & "')"
   conn.Execute "INSERT INTO castillos2 VALUES('oeste','" & castillo4 & "')"
   conn.Execute "INSERT INTO castillos2 VALUES('fortaleza','" & fortaleza & "')"
End If
'------------------------
Exit Sub
mal:
End Sub
Public Sub EstadisticasPjs(UserIndex As Integer)
  Dim str As String
'conn.Execute "INSERT INTO estadis VALUES('" & UserList(Userindex).Name & "','" & UserList(Userindex).Stats.GLD & "','" & UserList(Userindex).Stats.Banco & "','" & UserList(Userindex).Remort & "','" & UserList(Userindex).Stats.MaxHP & "','" & UserList(Userindex).GuildInfo.GuildName & "','" & UserList(Userindex).Stats.MaxHIT & "','" & UserList(Userindex).Stats.Fama & "','" & UserList(Userindex).Stats.Elu & "','" & UserList(Userindex).Stats.ELV & "','" & UserList(Userindex).Genero & "','" & UserList(Userindex).clase & "','" & UserList(Userindex).raza & "','" & UserList(Userindex).Stats.NPCsMuertos & "')"



If UserList(UserIndex).BD = 1 Then

str = "UPDATE `estadis` SET"
str = str & " gld=" & UserList(UserIndex).Stats.GLD
str = str & ",banco=" & UserList(UserIndex).Stats.Banco
str = str & ",remort=" & UserList(UserIndex).Remort
str = str & ",maxhp=" & UserList(UserIndex).Stats.MaxHP
str = str & ",clan='" & UserList(UserIndex).GuildInfo.GuildName & "'"
str = str & ",maxhit=" & UserList(UserIndex).Stats.MaxHIT
str = str & ",fama=" & UserList(UserIndex).Stats.Fama
str = str & ",elu=" & UserList(UserIndex).Stats.Elu
str = str & ",elv=" & UserList(UserIndex).Stats.ELV
str = str & ",genero='" & UserList(UserIndex).Genero & "'"
str = str & ",clase='" & UserList(UserIndex).clase & "'"
str = str & ",raza='" & UserList(UserIndex).raza & "'"
str = str & ",muertes=" & UserList(UserIndex).Stats.NPCsMuertos
str = str & " WHERE nombre='" & UserList(UserIndex).Name & "'"
Call conn.Execute(str)
Else
conn.Execute "INSERT INTO estadis VALUES('" & UserList(UserIndex).Name & "','" & UserList(UserIndex).Stats.GLD & "','" & UserList(UserIndex).Stats.Banco & "','" & UserList(UserIndex).Remort & "','" & UserList(UserIndex).Stats.MaxHP & "','" & UserList(UserIndex).GuildInfo.GuildName & "','" & UserList(UserIndex).Stats.MaxHIT & "','" & UserList(UserIndex).Stats.Fama & "','" & UserList(UserIndex).Stats.Elu & "','" & UserList(UserIndex).Stats.ELV & "','" & UserList(UserIndex).Genero & "','" & UserList(UserIndex).clase & "','" & UserList(UserIndex).raza & "','" & UserList(UserIndex).Stats.NPCsMuertos & "')"
UserList(UserIndex).BD = 1
Call WriteVar(CharPath & Left$(UCase$(UserList(UserIndex).Name), 1) & "\" & UCase$(UserList(UserIndex).Name) & ".chr", "INIT", "BD", val(UserList(UserIndex).BD))
End If

End Sub
Public Sub BDDConquistanCastillo(cual As String, clanx As String)

    Dim rs As ADODB.Recordset
   On Error GoTo mal
    'pluto:2.24
If ServerPrimario = 1 Then
    Set rs = New ADODB.Recordset
    If BD = False Then Exit Sub
    If rs Is Nothing Then GoTo mal
    rs.Open "SELECT * FROM castillos WHERE castillo='" & cual & "'", conn, adOpenKeyset, adLockOptimistic
    rs!clan = clanx
    rs.Update
    rs.Close
    Set rs = Nothing
Else
 Set rs = New ADODB.Recordset
    If BD = False Then Exit Sub
    If rs Is Nothing Then GoTo mal
    rs.Open "SELECT * FROM castillos2 WHERE castillo='" & cual & "'", conn, adOpenKeyset, adLockOptimistic
    rs!clan = clanx
    rs.Update
    rs.Close
    Set rs = Nothing
 End If
    Exit Sub
mal:
End Sub

'Public Sub BDDResetGMsos()
    ''''an error GoTo mal
    'If Bd = False Then Exit Sub
    'conn.Execute "DELETE FROM sos"
    'Exit Sub
'mal:
'End Sub

'Public Sub BDDAddGMsos(user As String, razon As String)
    '''an error GoTo mal
    'If Bd = False Then Exit Sub
    'conn.Execute "INSERT INTO sos VALUES('" & user & "','" & razon & "')"
    'Exit Sub
'mal:
'End Sub

'Public Sub BDDDelGMsos(user As String)
    '''an error GoTo mal
    'If Bd = False Then Exit Sub
    'conn.Execute "DELETE FROM sos WHERE user='" & user & "'"
    'Exit Sub
'mal:
'End Sub

'Public Sub BDDAddBanIP(by As String, ip As String)
    'conn.Execute "INSERT INTO banip VALUES('" & by & "','" & ip & "')"
'End Sub

'Public Sub BDDDelBanIP(ip As String)
    'conn.Execute "DELETE FROM banip WHERE ip LIKE '" & ip & "'"
'End Sub

'Function BDDIsBanIP(ip As String) As Boolean
'quitar esto
'Exit Function
    
    'Dim BDD As ADODB.Recordset
   ' Set BDD = New ADODB.Recordset
    'BDD.Open "SELECT * FROM banip WHERE '" & ip & "' LIKE ip", conn
    'If BDD.EOF Then
        'BDDIsBanIP = False
    'Else
        'BDDIsBanIP = True
    'End If
    'BDD.Close
    'Set BDD = Nothing
'End Function

Public Sub BDDConnect()



On Error GoTo mal


'Quitar ESTO
If BaseDatos = 0 Then Exit Sub
'Exit Sub

'pluto:2.24
If ServerPrimario = 1 Then
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "DRIVER={MySQL ODBC 3.51 Driver};" _
      & "SERVER=localhost;" _
      & "DATABASE=drag;" _
      & "UID=desiree;PWD=gwh28dgcmp; OPTION=3"
    conn.Open
    conn.Execute "CREATE TABLE IF NOT EXISTS online(numero int,rhoy int, rayer int)"
   ' conn.Execute "DROP TABLE IF EXISTS sos"
   ' conn.Execute "CREATE TABLE sos(user text, razon text)"
    conn.Execute "DROP TABLE IF EXISTS castillos"
    conn.Execute "CREATE TABLE castillos(castillo text, clan text)"
   ' conn.Execute "CREATE TABLE IF NOT EXISTS gms(nombre text, hash int, online int)"
   ' conn.Execute "UPDATE gms SET online=0"
   ' conn.Execute "CREATE TABLE IF NOT EXISTS cuentas(mail text, pass text, activada int, hash int, sndmail int)"
   ' conn.Execute "CREATE TABLE IF NOT EXISTS recovery(mail text, pass text, hash int, activada int, sndmail int)"
   ' conn.Execute "CREATE TABLE IF NOT EXISTS banip(quien text, ip text)"
    BD = True
'frmMain.Cuentas.Enabled = True

Else
Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "DRIVER={MySQL ODBC 3.51 Driver};" _
      & "SERVER=localhost;" _
      & "DATABASE=drag;" _
      & "UID=desiree;PWD=gwh28dgcmp; OPTION=3"
    conn.Open
    conn.Execute "CREATE TABLE IF NOT EXISTS online2(numero int,rhoy int, rayer int)"
   ' conn.Execute "DROP TABLE IF EXISTS sos"
   ' conn.Execute "CREATE TABLE sos(user text, razon text)"
    conn.Execute "DROP TABLE IF EXISTS castillos2"
    conn.Execute "CREATE TABLE castillos2(castillo text, clan text)"
   ' conn.Execute "CREATE TABLE IF NOT EXISTS gms(nombre text, hash int, online int)"
   ' conn.Execute "UPDATE gms SET online=0"
   ' conn.Execute "CREATE TABLE IF NOT EXISTS cuentas(mail text, pass text, activada int, hash int, sndmail int)"
   ' conn.Execute "CREATE TABLE IF NOT EXISTS recovery(mail text, pass text, hash int, activada int, sndmail int)"
   ' conn.Execute "CREATE TABLE IF NOT EXISTS banip(quien text, ip text)"
    BD = True
'frmMain.Cuentas.Enabled = True
   

End If

Exit Sub
mal:
MsgBox ("NO SE PUDO CONECTAR A LA BASE DE DATOS")
BD = False
End
End Sub



