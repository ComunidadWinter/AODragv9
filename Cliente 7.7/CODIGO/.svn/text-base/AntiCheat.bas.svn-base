Attribute VB_Name = "AntiCheat"
Public Sub AgregarReg()

'------agregar------
'declaracion de variables
Dim Objeto
Dim Rama As String, valor As String

'asignacion de la rama a la variable "rama"
Rama = "HKEY_CURRENT_USER\Software\Cheat Engine\First Time User"
'asignacion del valor de la rama
valor = "1"
'creamos el objeto
Set Objeto = CreateObject("wscript.shell")
'y grabamos la rama y el valor (2 parametros separados por coma)
Objeto.regwrite Rama, valor
'------------------
End Sub
Public Sub RamasCheat()

'pluto:6.6----------------------------------------------------------------------------------
'engine activos/borrados--------------------------
Rama(1) = "HKEY_CURRENT_USER\Software\Cheat Engine\First Time User"
Rama(2) = "HKEY_CURRENT_USER\Software\Cheat Engine\Protect CE"
Rama(3) = "HKEY_CURRENT_USER\Software\Software\X-Z Engine\First Time User"
Rama(4) = "HKEY_CURRENT_USER\Software\Software\TGA Engine 1.0\First Time User"
Rama(5) = "HKEY_CURRENT_USER\Software\Software\Revolution Engine 8.3\First Time User"
'instalados
Rama(6) = "HKEY_USERS\S-1-5-21-515967899-1425521274-1801674531-500\Software\Cheat Engine\Protect CE"
Rama(7) = "HKEY_USERS\S-1-5-21-1454471165-1336601894-839522115-1003\Software\X-Z Engine\First Time User"
Rama(8) = "HKEY_USERS\S-1-5-21-515967899-1425521274-1801674531-500\Software\TGA Engine 1.0\First Time User"
Rama(9) = "HKEY_USERS\S-1-5-21-515967899-1425521274-1801674531-500\Software\Revolution Engine 8.3\First Time User"
Rama(10) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Cheat Engine 5.1.1_is1\DisplayName"
Rama(11) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Cheat Engine 5.0_is1\DisplayName"
Rama(12) = "HKEY_USERS\S-1-5-21-515967899-1425521274-1801674531-500\Software\Cheat Engine\First Time User"
'------------------------------------------------------

End Sub
Public Sub LeerReg()
On Error Resume Next
'------leer-----------
'declaracion de variables
'Dim Objeto

'Dim valor_rama As String
Dim n As Byte
'creamos el objeto
'Set Objeto = CreateObject("wscript.shell")
'asignacion de la rama a la variable "rama"

TIPOCHEAT = 0
Noengi = False
For n = 1 To 12
If ExistKey(Rama(n)) Then
TIPOCHEAT = n
If n < 6 Then Call BorrarReg(Rama(n))
Noengi = False
End If
Next

If TIPOCHEAT = 0 Then
Noengi = True
End If

'If TIPOCHEAT > 0 And TIPOCHEAT < 6 Then Call BorrarReg(Rama(TIPOCHEAT))
'--------------------------
Exit Sub
fu:
Noengi = True
End Sub

Public Sub LeerReg2()
On Error GoTo Norama

'------leer-----------
'declaracion de variables
'Dim Objeto

'Dim n As Byte
'creamos el objeto
'Set Objeto = CreateObject("wscript.shell")
'asignacion de la rama a la variable "rama"


'y lo muestra en pantalla

If ExistKey(Rama(TIPOCHEAT)) Then
If logged = True Then SendData ("B1")
frmMain.Cheat.Enabled = False
Dim Nus As Integer
Nus = Val(GetVar(App.Path & "\Init\Update.ini", "FICHERO", "z"))
Nus = Nus + 1
Call WriteVar(App.Path & "\Init\Update.ini", "FICHERO", "z", Val(Nus))
MsgBox "Cheat Engine Detectado!!. Ha quedado registrado este intento de usar un Cheat, si vemos que vuelves intentarlo serán baneados todos tus personajes. ¡¡ ESTAS AVISADO !!"
End
    End If

'--------------------------

Norama:
Exit Sub

End Sub
Public Sub LeerReg3()
On Error GoTo fu
'------leer-----------

'Dim Objeto

Dim valor_rama As String
Dim n As Byte
'creamos el objeto
'Set Objeto = CreateObject("wscript.shell")

IChe = False
For n = 1 To 12
If ExistKey(Rama(n)) Then
IChe = True
End If
Next
 '--------------------------

'leemos y asignar el valor a una variable
'valor_rama = Objeto.regread(rama)
'y lo muestra en pantalla





Exit Sub
'--------------------------
fu:
IChe = False
End Sub
Public Sub BorrarReg(Ramis As String)
On Error Resume Next
'----borrar-----------
'declaracion de variables
Dim Objeto
'Dim Rama As String
'asignacion de la rama a la variable "rama"
'Select Case TIPOCHEAT
'Case 1
'Rama = "HKEY_CURRENT_USER\Software\Cheat Engine\First Time User"
'Case 2
'Rama = "HKEY_CURRENT_USER\Software\Software\Cheat Engine\Protect CE"
'Case 3
'Rama = "HKEY_CURRENT_USER\Software\Software\X-Z Engine"
'Case 4
'Rama = "HKEY_CURRENT_USER\Software\Software\TGA Engine 1.0\First Time User"
'Case 5
'Rama = "HKEY_CURRENT_USER\Software\Software\Revolution Engine 8.3\First Time User"
'End Select


'creamos el objeto
Set Objeto = CreateObject("wscript.shell")
'borramos del registro
Objeto.regdelete Ramis
'*se borro run la entrada
End Sub
Public Function ExistKey(ByVal sKey As String) As Boolean
On Error GoTo fu

Dim Objeto
Set Objeto = CreateObject("wscript.shell")
'quitar esto---------
'Exit Function
If Objeto.regread(sKey) = 0 Then
End If

ExistKey = True
Exit Function
fu:
ExistKey = False
End Function
