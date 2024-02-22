Attribute VB_Name = "mdlCOmercioConUsuario"
'Modulo para comerciar con otro usuario
'Por Alejo (Alejandro Santos)
'
'
'[Alejo]

Option Explicit

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    Objeto As Integer 'Indice del inventario a comerciar, que objeto desea dar
    
    'El tipo de datos de Cant ahora es Long (antes Integer)
    'asi se puede comerciar con oro > 32k
    '[CORREGIDO]
    Cant As Long 'Cuantos comerciar, cuantos objetos desea dar
    '[/CORREGIDO]
    Acepto As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(Origen As Integer, Destino As Integer)
On Error GoTo fallo

'Actualiza el inventario del usuario
Call UpdateUserInv(True, Origen, 0)

'Decirle al origen que abra la ventanita.
Call SendData(ToIndex, Origen, 0, "CU")
UserList(Origen).flags.Comerciando = True

'si es el receptor, enviamos el objeto del otro usu
'If UserList(UserList(Origen).ComUsu.DestUsu).ComUsu.DestUsu = Origen Then
If UserList(Origen).ComUsu.DestUsu = Destino Then
    Call EnviarObjetoTransaccion(Origen)
End If

Exit Sub
fallo:
Call LogError("iniciarcomerciousuario " & Err.number & " D: " & Err.Description)


End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(AQuien As Integer)
On Error GoTo errhandler

If AQuien = 0 Then Exit Sub
'Dim Object As UserOBJ
Dim ObjInd As Integer
Dim ObjCant As Long
'pluto:2.9.0
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = 0 Then Exit Sub
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = 1281 Then Exit Sub
'[Alejo]: En esta funcion se centralizaba el problema
'         de no poder comerciar con mas de 32k de oro.
'         Ahora si funciona!!!

'Object.Amount = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
    'Object.ObjIndex = iORO
    ObjInd = iORO
Else
    'Object.ObjIndex = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
    ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
End If

'If Object.ObjIndex > 0 And Object.Amount > 0 Then
'    Call SendData(ToIndex, AQuien, 0, "COMUSUINV" & 1 & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
'    & ObjData(Object.ObjIndex).ObjType & "," _
'    & ObjData(Object.ObjIndex).MaxHIT & "," _
'    & ObjData(Object.ObjIndex).MinHIT & "," _
'    & ObjData(Object.ObjIndex).MaxDef & "," _
'    & ObjData(Object.ObjIndex).Valor \ 3)
'End If



If ObjInd > 0 And ObjCant > 0 Then
'pluto:2.12---------------------------
   Dim flu As String
   flu = ObjData(ObjInd).Name
   If ObjData(ObjInd).OBJType = 60 Then
   Dim flu2 As Byte
   flu2 = ObjInd - 887
   flu = ObjData(ObjInd).Name & "  Niv: " & UserList(UserList(AQuien).ComUsu.DestUsu).Montura.Nivel(flu2) & " Exp: " & UserList(UserList(AQuien).ComUsu.DestUsu).Montura.exp(flu2)
   End If
 '------------------------------------
 
 
      'pluto:2.3
    Call SendData2(ToIndex, AQuien, 0, 72, 1 & "," & ObjInd & "," & flu & "," & ObjCant & "," & 0 & "," & ObjData(ObjInd).GrhIndex & "," _
    & ObjData(ObjInd).OBJType & "," _
    & ObjData(ObjInd).MaxHIT & "," _
    & ObjData(ObjInd).MinHIT & "," _
    & ObjData(ObjInd).MaxDef & "," _
    & ObjData(ObjInd).Valor \ 3 & "," _
    & ObjData(ObjInd).SubTipo)
    
End If
Exit Sub
errhandler:
 Call LogError("Enviarobjetotransaccion")
End Sub

Public Sub FinComerciarUsu(UserIndex As Integer)
On Error GoTo fallo
If UserIndex = 0 Then Exit Sub
UserList(UserIndex).ComUsu.Acepto = False
UserList(UserIndex).ComUsu.Cant = 0
UserList(UserIndex).ComUsu.DestUsu = 0
UserList(UserIndex).ComUsu.Objeto = 0

UserList(UserIndex).flags.Comerciando = False
'pluto:2.7.0
Call SendData(ToIndex, UserIndex, 0, "||Ha finalizado el Comercio." & "´" & FontTypeNames.FONTTYPE_COMERCIO)

Call SendData(ToIndex, UserIndex, 0, "CF")

Exit Sub
fallo:
Call LogError("fincomerciarusu " & Err.number & " D: " & Err.Description)

End Sub

Public Sub AceptarComercioUsu(UserIndex As Integer)

On Error GoTo errhandler
Dim ii As Byte
'quitar todos los avis= son indicadores
Dim avis As Byte
avis = 0
If UserIndex = 0 Then Exit Sub
If UserList(UserIndex).ComUsu.DestUsu <= 0 Then Exit Sub
If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then Exit Sub


UserList(UserIndex).ComUsu.Acepto = True

If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False Then
    Call SendData(ToIndex, UserIndex, 0, "||El otro usuario aun no ha aceptado tu oferta." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
    Exit Sub
End If
avis = 1
Dim Obj1 As obj, Obj2 As obj
Dim OtroUserIndex As Integer
Dim TerminarAhora As Boolean

TerminarAhora = False
OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
'pluto:2.10
If UserList(UserIndex).ComUsu.Objeto = FLAGORO And UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
Call SendData(ToIndex, UserIndex, 0, "||No podéis intercambiar Oro" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, OtroUserIndex, 0, "||No podéis intercambiar Oro" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
TerminarAhora = True
GoTo fuera
End If





'[Alejo]: Creo haber podido erradicar el bug de
'         no poder comerciar con mas de 32k de oro.
'         Las lineas comentadas en los siguientes
'         2 grandes bloques IF (4 lineas) son las
'         que originaban el problema.

If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
    'Obj1.Amount = UserList(UserIndex).ComUsu.Cant
    Obj1.ObjIndex = iORO
    'If Obj1.Amount > UserList(UserIndex).Stats.GLD Then
    If UserList(UserIndex).ComUsu.Cant > UserList(UserIndex).Stats.GLD Then
        Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
        TerminarAhora = True
    End If
Else
'pluto:2.7.0
avis = 2
Dim chorizo As Integer
chorizo = UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).ObjIndex
   
    
    Obj1.Amount = UserList(UserIndex).ComUsu.Cant
    Obj1.ObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).ObjIndex
    If Obj1.Amount > UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
        TerminarAhora = True
    End If
End If
avis = 3
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = iORO
    'If Obj2.Amount > UserList(OtroUserIndex).Stats.GLD Then
    If UserList(OtroUserIndex).ComUsu.Cant > UserList(OtroUserIndex).Stats.GLD Then
        Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
        TerminarAhora = True
    End If

'pluto:2.7.0
If UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).Equipped = 1 Then
Call SendData(ToIndex, UserIndex, 0, "||Tienes ese objeto Equipado" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, OtroUserIndex, 0, "||El otro user tiene ese objeto Equipado." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
TerminarAhora = True
End If



Dim i2 As Byte
i2 = 0
For ii = 1 To MAX_INVENTORY_SLOTS
If UserList(UserIndex).Invent.Object(ii).ObjIndex = 0 Then i2 = i2 + 1
Next ii
If i2 < 2 Then
TerminarAhora = True
Call Encarcelar(UserIndex, 30)
Call LogCasino("/CARCEL AUTOMATICO COMERCIO" & UserList(UserIndex).Name)
End If
avis = 4
Else
  'pluto:2.7.0
Dim chorizo2 As Integer
chorizo2 = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex
   
    Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex
    If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
        TerminarAhora = True
    End If

If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then GoTo ee

'pluto:2.7.0
If UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Equipped = 1 Then
Call SendData(ToIndex, OtroUserIndex, 0, "||Tienes ese objeto Equipado" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, UserIndex, 0, "||El otro user tiene ese objeto Equipado." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
TerminarAhora = True
End If



'pluto:2.9.0
Dim i1 As Byte
avis = 5
i1 = 0

For ii = 1 To MAX_INVENTORY_SLOTS
If UserList(OtroUserIndex).Invent.Object(ii).ObjIndex = 0 Then i1 = i1 + 1
Next ii
If i1 < 2 Then
TerminarAhora = True
Call Encarcelar(OtroUserIndex, 30)
Call LogCasino("/CARCEL AUTOMATICO COMERCIO" & UserList(OtroUserIndex).Name)
End If






'pluto:2.9.0
ee:
If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then GoTo ee2
If UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).Equipped = 1 Then
Call SendData(ToIndex, UserIndex, 0, "||Tienes ese objeto Equipado" & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, OtroUserIndex, 0, "||El otro user tiene ese objeto Equipado." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
TerminarAhora = True
End If


ee2:
End If

avis = 6

'PLuto:2.10
If (chorizo > 887 And chorizo < 900) And (chorizo2 > 887 And chorizo2 < 900) Then
Call SendData(ToIndex, UserIndex, 0, "||No se puede comerciar una mascota por otra." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, OtroUserIndex, 0, "||No se puede comerciar una mascota por otra." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
TerminarAhora = True
End If
'pluto:2.14
    If ObjData(Obj1.ObjIndex).Caos > 0 Or ObjData(Obj1.ObjIndex).Real > 0 Or ObjData(Obj2.ObjIndex).Real > 0 Or ObjData(Obj2.ObjIndex).Caos > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||No se puede comerciar con Ropas de Armadas." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, OtroUserIndex, 0, "||No se puede comerciar con Ropas de Armadas." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
TerminarAhora = True
End If
'pluto:2.15
If (chorizo > 887 And chorizo < 900) Then
If UserList(UserIndex).Montura.Elu(chorizo - 887) = 0 Then
Call SendData(ToGM, 0, 0, "|| Comercio Mascota Bugueada: " & UserList(UserIndex).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call LogMascotas("BUG comercioMASCOTA Serie: " & UserList(UserIndex).Serie & " IP: " & UserList(UserIndex).ip & " Nom: " & UserList(UserIndex).Name)

TerminarAhora = True
End If
End If

If (chorizo2 > 887 And chorizo2 < 900) Then
If UserList(OtroUserIndex).Montura.Elu(chorizo2 - 887) = 0 Then
Call SendData(ToGM, 0, 0, "|| Comercio Mascota Bugueada: " & UserList(OtroUserIndex).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call LogMascotas("BUG comercioMASCOTA Serie: " & UserList(OtroUserIndex).Serie & " IP: " & UserList(OtroUserIndex).ip & " Nom: " & UserList(OtroUserIndex).Name)

TerminarAhora = True
End If
End If
'-----------------------
'pluto:6.0A
'If UserList(UserIndex).Nmonturas > 2 Or UserList(OtroUserIndex).Nmonturas > 2 Then
'Call SendData(ToIndex, UserIndex, 0, "||No se puede tener más de Tres Mascotas." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
'Call SendData(ToIndex, OtroUserIndex, 0, "||No se puede tener más de Tres Mascotas." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
'TerminarAhora = True
'End If


fuera:
'Por si las moscas...
If TerminarAhora = True Then
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If

'pluto:2.7.0

'---jugador 1-----
If chorizo > 887 And chorizo < 900 Then
avis = 7
Dim userfile As String
Dim userfile2 As String
userfile = CharPath & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".chr"
userfile2 = CharPath & Left$(UserList(OtroUserIndex).Name, 1) & "\" & UserList(OtroUserIndex).Name & ".chr"
Dim n As Byte
For n = 1 To 3
If val(GetVar(userfile2, "MONTURA" & n, "TIPO")) = chorizo - 887 Then
Call SendData(ToIndex, UserIndex, 0, "||Ese Pj ya tiene ese tipo de Mascota." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, OtroUserIndex, 0, "||Ya tienes ese tipo de Mascota." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call FinComerciarUsu(UserIndex)
Call FinComerciarUsu(OtroUserIndex)
Exit Sub
End If
Next n

'pluto:6.0A
If UserList(OtroUserIndex).Nmonturas > 2 Then
Call SendData(ToIndex, UserIndex, 0, "||Ese Personaje ya tiene Tres Mascotas." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, OtroUserIndex, 0, "||No se puede tener más de Tres Mascotas." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call FinComerciarUsu(UserIndex)
Call FinComerciarUsu(OtroUserIndex)
Exit Sub
End If
'If val(GetVar(userfile2, "MONTURA", "NIVEL" & chorizo - 887)) > 0 Then
'Call SendData(ToIndex, Userindex, 0, "||Ese Pj ya tiene ese tipo de Mascota." & FONTTYPENAMES.FONTTYPE_COMERCIO)
'Call SendData(ToIndex, OtroUserIndex, 0, "||Ya tienes ese tipo de Mascota." & FONTTYPENAMES.FONTTYPE_COMERCIO)
'Call FinComerciarUsu(Userindex)
'Call FinComerciarUsu(OtroUserIndex)
'Exit Sub
'End If


'Dim xx, x1, x2, x3, x4, x5 As Integer
Dim x1 As Byte
Dim x2 As Long
Dim x3 As Long
Dim x4 As Integer
Dim x5 As Integer
Dim xx As Integer
Dim x6 As String
'pluto:6.0A
Dim x7 As Byte
Dim x8 As Byte
Dim x9 As Byte
Dim x10 As Byte
Dim x11 As Byte
Dim x12 As Byte
Dim x13 As Byte
Dim x14 As Byte
Dim x15 As Byte
Dim x16 As Byte
xx = chorizo - 887
x1 = UserList(UserIndex).Montura.Nivel(xx)
x2 = UserList(UserIndex).Montura.exp(xx)
x3 = UserList(UserIndex).Montura.Elu(xx)
x4 = UserList(UserIndex).Montura.Vida(xx)
x5 = UserList(UserIndex).Montura.Golpe(xx)
x6 = UserList(UserIndex).Montura.Nombre(xx)
x7 = UserList(UserIndex).Montura.AtCuerpo(xx)
x8 = UserList(UserIndex).Montura.Defcuerpo(xx)
x9 = UserList(UserIndex).Montura.AtFlechas(xx)
x10 = UserList(UserIndex).Montura.DefFlechas(xx)
x11 = UserList(UserIndex).Montura.AtMagico(xx)
x12 = UserList(UserIndex).Montura.DefMagico(xx)
x13 = UserList(UserIndex).Montura.Evasion(xx)
x14 = UserList(UserIndex).Montura.Libres(xx)
x15 = UserList(UserIndex).Montura.Tipo(xx)

x16 = UserList(UserIndex).Montura.index(xx)

Call LogMascotas("Comercio " & UserList(UserIndex).Name & " ofrece Mascota: " & x6 & " tiene " & UserList(UserIndex).Nmonturas)
Call LogMascotas("Comercio " & UserList(OtroUserIndex).Name & " acepta Mascota: " & x6 & " tiene " & UserList(OtroUserIndex).Nmonturas)


'tomamos valores del user1 al user2 excepto el index (x16)
UserList(OtroUserIndex).Montura.Nivel(xx) = val(x1)
UserList(OtroUserIndex).Montura.exp(xx) = val(x2)
UserList(OtroUserIndex).Montura.Elu(xx) = val(x3)
UserList(OtroUserIndex).Montura.Vida(xx) = val(x4)
UserList(OtroUserIndex).Montura.Golpe(xx) = val(x5)
UserList(OtroUserIndex).Montura.Nombre(xx) = x6
UserList(OtroUserIndex).Montura.AtCuerpo(xx) = val(x7)
UserList(OtroUserIndex).Montura.Defcuerpo(xx) = val(x8)
UserList(OtroUserIndex).Montura.AtFlechas(xx) = val(x9)
UserList(OtroUserIndex).Montura.DefFlechas(xx) = val(x10)
UserList(OtroUserIndex).Montura.AtMagico(xx) = val(x11)
UserList(OtroUserIndex).Montura.DefMagico(xx) = val(x12)
UserList(OtroUserIndex).Montura.Evasion(xx) = val(x13)
UserList(OtroUserIndex).Montura.Libres(xx) = val(x14)
UserList(OtroUserIndex).Montura.Tipo(xx) = val(x15)


'buscamos el index
For n = 1 To 3
If val(GetVar(userfile2, "MONTURA" & n, "TIPO")) = 0 Then GoTo gb
Next
Call LogMascotas("Comercio NO INDEX LIBRE en " & UserList(OtroUserIndex).Name)
gb:
'guardamos el index pero no hace falta grabarlo
UserList(OtroUserIndex).Montura.index(xx) = n
Call LogMascotas("Comercio metemos en INDEX: " & n & " una Mascota: " & x6 & " al user " & UserList(OtroUserIndex).Name)
'guardamos en ficha user2
Call WriteVar(userfile2, "MONTURA" & n, "NIVEL", val(x1))
Call WriteVar(userfile2, "MONTURA" & n, "EXP", val(x2))
Call WriteVar(userfile2, "MONTURA" & n, "ELU", val(x3))
Call WriteVar(userfile2, "MONTURA" & n, "VIDA", val(x4))
Call WriteVar(userfile2, "MONTURA" & n, "GOLPE", val(x5))
Call WriteVar(userfile2, "MONTURA" & n, "NOMBRE", x6)
Call WriteVar(userfile2, "MONTURA" & n, "ATCUERPO", val(x7))
Call WriteVar(userfile2, "MONTURA" & n, "DEFCUERPO", val(x8))
Call WriteVar(userfile2, "MONTURA" & n, "ATFLECHAS", val(x9))
Call WriteVar(userfile2, "MONTURA" & n, "DEFFLECHAS", val(x10))
Call WriteVar(userfile2, "MONTURA" & n, "ATMAGICO", val(x11))
Call WriteVar(userfile2, "MONTURA" & n, "DEFMAGICO", val(x12))
Call WriteVar(userfile2, "MONTURA" & n, "EVASION", val(x13))
Call WriteVar(userfile2, "MONTURA" & n, "LIBRES", val(x14))
Call WriteVar(userfile2, "MONTURA" & n, "TIPO", val(x15))

'ponemos a cero la mascota del user1
Call ResetMontura(UserIndex, xx)
'ponemos a cero la ficha mascota user 1
Call WriteVar(userfile, "MONTURA" & x16, "NIVEL", 0)
Call WriteVar(userfile, "MONTURA" & x16, "EXP", 0)
Call WriteVar(userfile, "MONTURA" & x16, "ELU", 0)
Call WriteVar(userfile, "MONTURA" & x16, "VIDA", 0)
Call WriteVar(userfile, "MONTURA" & x16, "GOLPE", 0)
Call WriteVar(userfile, "MONTURA" & x16, "NOMBRE", "")
Call WriteVar(userfile, "MONTURA" & x16, "ATCUERPO", 0)
Call WriteVar(userfile, "MONTURA" & x16, "DEFCUERPO", 0)
Call WriteVar(userfile, "MONTURA" & x16, "ATFLECHAS", 0)
Call WriteVar(userfile, "MONTURA" & x16, "DEFFLECHAS", 0)
Call WriteVar(userfile, "MONTURA" & x16, "ATMAGICO", 0)
Call WriteVar(userfile, "MONTURA" & x16, "DEFMAGICO", 0)
Call WriteVar(userfile, "MONTURA" & x16, "EVASION", 0)
Call WriteVar(userfile, "MONTURA" & x16, "LIBRES", 0)
Call WriteVar(userfile, "MONTURA" & x16, "TIPO", 0)
Call LogMascotas("Comercio INDEX : " & x16 & " a cero en " & UserList(UserIndex).Name)

'sumamos y restamos mascotas
UserList(UserIndex).Nmonturas = UserList(UserIndex).Nmonturas - 1
UserList(OtroUserIndex).Nmonturas = UserList(OtroUserIndex).Nmonturas + 1
Call WriteVar(userfile, "MONTURAS", "NroMonturas", val(UserList(UserIndex).Nmonturas))
Call WriteVar(userfile2, "MONTURAS", "NroMonturas", val(UserList(OtroUserIndex).Nmonturas))
Call LogMascotas("Comercio " & UserList(UserIndex).Name & " resta 1 y ahora tiene " & UserList(UserIndex).Nmonturas)
Call LogMascotas("Comercio " & UserList(OtroUserIndex).Name & " suma 1 y ahora tiene " & UserList(OtroUserIndex).Nmonturas)

End If 'jugador 1

'---jugador 2-----
If chorizo2 > 887 And chorizo2 < 900 Then
avis = 8
userfile2 = CharPath & Left$(UserList(OtroUserIndex).Name, 1) & "\" & UserList(OtroUserIndex).Name & ".chr"
userfile = CharPath & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".chr"

For n = 1 To 3
If val(GetVar(userfile, "MONTURA" & n, "TIPO")) = chorizo2 - 887 Then
Call SendData(ToIndex, OtroUserIndex, 0, "||Ese Pj ya tiene ese tipo de Mascota." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, UserIndex, 0, "||Ya tienes ese tipo de Mascota." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call FinComerciarUsu(UserIndex)
Call FinComerciarUsu(OtroUserIndex)
Exit Sub
End If
Next n
'pluto:6.0A
If UserList(UserIndex).Nmonturas > 2 Then
Call SendData(ToIndex, OtroUserIndex, 0, "||Ese Personaje ya tiene Tres Mascotas." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call SendData(ToIndex, UserIndex, 0, "||No se puede tener más de Tres Mascotas." & "´" & FontTypeNames.FONTTYPE_COMERCIO)
Call FinComerciarUsu(UserIndex)
Call FinComerciarUsu(OtroUserIndex)
Exit Sub
End If

'If val(GetVar(userfile, "MONTURA", "NIVEL" & chorizo2 - 887)) > 0 Then
'Call SendData(ToIndex, OtroUserIndex, 0, "||Ese Pj ya tiene ese tipo de Mascota." & FONTTYPENAMES.FONTTYPE_COMERCIO)
'Call SendData(ToIndex, Userindex, 0, "||Ya tienes ese tipo de Mascota." & FONTTYPENAMES.FONTTYPE_COMERCIO)
'Call FinComerciarUsu(OtroUserIndex)
'Call FinComerciarUsu(Userindex)
'Exit Sub
'End If
xx = chorizo2 - 887
x1 = UserList(OtroUserIndex).Montura.Nivel(xx)
x2 = UserList(OtroUserIndex).Montura.exp(xx)
x3 = UserList(OtroUserIndex).Montura.Elu(xx)
x4 = UserList(OtroUserIndex).Montura.Vida(xx)
x5 = UserList(OtroUserIndex).Montura.Golpe(xx)
x6 = UserList(OtroUserIndex).Montura.Nombre(xx)
x7 = UserList(OtroUserIndex).Montura.AtCuerpo(xx)
x8 = UserList(OtroUserIndex).Montura.Defcuerpo(xx)
x9 = UserList(OtroUserIndex).Montura.AtFlechas(xx)
x10 = UserList(OtroUserIndex).Montura.DefFlechas(xx)
x11 = UserList(OtroUserIndex).Montura.AtMagico(xx)
x12 = UserList(OtroUserIndex).Montura.DefMagico(xx)
x13 = UserList(OtroUserIndex).Montura.Evasion(xx)
x14 = UserList(OtroUserIndex).Montura.Libres(xx)
x15 = UserList(OtroUserIndex).Montura.Tipo(xx)

x16 = UserList(OtroUserIndex).Montura.index(xx)

UserList(UserIndex).Montura.Nivel(xx) = val(x1)
UserList(UserIndex).Montura.exp(xx) = val(x2)
UserList(UserIndex).Montura.Elu(xx) = val(x3)
UserList(UserIndex).Montura.Vida(xx) = val(x4)
UserList(UserIndex).Montura.Golpe(xx) = val(x5)
UserList(UserIndex).Montura.Nombre(xx) = x6
UserList(UserIndex).Montura.AtCuerpo(xx) = val(x7)
UserList(UserIndex).Montura.Defcuerpo(xx) = val(x8)
UserList(UserIndex).Montura.AtFlechas(xx) = val(x9)
UserList(UserIndex).Montura.DefFlechas(xx) = val(x10)
UserList(UserIndex).Montura.AtMagico(xx) = val(x11)
UserList(UserIndex).Montura.DefMagico(xx) = val(x12)
UserList(UserIndex).Montura.Evasion(xx) = val(x13)
UserList(UserIndex).Montura.Libres(xx) = val(x14)
UserList(UserIndex).Montura.Tipo(xx) = val(x15)


'buscamos el index
For n = 1 To 3
If val(GetVar(userfile, "MONTURA" & n, "TIPO")) = 0 Then GoTo gb2
Next
Call LogMascotas("Comercio NO INDEX LIBRE en " & UserList(UserIndex).Name)
gb2:

'guardamos el index pero no hace falta grabarlo
UserList(UserIndex).Montura.index(xx) = n
Call LogMascotas("Comercio metemos en INDEX: " & n & " una Mascota: " & x6 & " al user " & UserList(UserIndex).Name)
'guardamos en ficha user1
Call WriteVar(userfile, "MONTURA" & n, "NIVEL", val(x1))
Call WriteVar(userfile, "MONTURA" & n, "EXP", val(x2))
Call WriteVar(userfile, "MONTURA" & n, "ELU", val(x3))
Call WriteVar(userfile, "MONTURA" & n, "VIDA", val(x4))
Call WriteVar(userfile, "MONTURA" & n, "GOLPE", val(x5))
Call WriteVar(userfile, "MONTURA" & n, "NOMBRE", x6)
Call WriteVar(userfile, "MONTURA" & n, "ATCUERPO", val(x7))
Call WriteVar(userfile, "MONTURA" & n, "DEFCUERPO", val(x8))
Call WriteVar(userfile, "MONTURA" & n, "ATFLECHAS", val(x9))
Call WriteVar(userfile, "MONTURA" & n, "DEFFLECHAS", val(x10))
Call WriteVar(userfile, "MONTURA" & n, "ATMAGICO", val(x11))
Call WriteVar(userfile, "MONTURA" & n, "DEFMAGICO", val(x12))
Call WriteVar(userfile, "MONTURA" & n, "EVASION", val(x13))
Call WriteVar(userfile, "MONTURA" & n, "LIBRES", val(x14))
Call WriteVar(userfile, "MONTURA" & n, "TIPO", val(x15))

'ponermos a cero user2
Call ResetMontura(OtroUserIndex, xx)

'ponemos a cero ficha user2

Call WriteVar(userfile2, "MONTURA" & x16, "NIVEL", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "EXP", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "ELU", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "VIDA", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "GOLPE", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "NOMBRE", "")
Call WriteVar(userfile2, "MONTURA" & x16, "ATCUERPO", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "DEFCUERPO", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "ATFLECHAS", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "DEFFLECHAS", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "ATMAGICO", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "DEFMAGICO", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "EVASION", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "LIBRES", 0)
Call WriteVar(userfile2, "MONTURA" & x16, "TIPO", 0)
Call LogMascotas("Comercio INDEX : " & x16 & " a cero en " & UserList(OtroUserIndex).Name)

'sumamos y restamos mascotas
'If UserList(OtroUserIndex).Nmonturas < 1 Then GoTo noo
UserList(OtroUserIndex).Nmonturas = UserList(OtroUserIndex).Nmonturas - 1
noo:
UserList(UserIndex).Nmonturas = UserList(UserIndex).Nmonturas + 1
Call WriteVar(userfile2, "MONTURAS", "NroMonturas", val(UserList(OtroUserIndex).Nmonturas))
Call WriteVar(userfile, "MONTURAS", "NroMonturas", val(UserList(UserIndex).Nmonturas))

Call LogMascotas("Comercio " & UserList(OtroUserIndex).Name & " resta 1 y ahora tiene " & UserList(OtroUserIndex).Nmonturas)
Call LogMascotas("Comercio " & UserList(UserIndex).Name & " suma 1 y ahora tiene " & UserList(UserIndex).Nmonturas)

End If 'jugador 2








'[CORREGIDO]
'Desde acá corregí el bug que cuando se ofrecian mas de
'10k de oro no le llegaban al destinatario.
avis = 9
'pone el oro directamente en la billetera
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Cant
    Call SendUserStatsOro(OtroUserIndex)
    'y se la doy al otro
    'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Cant
    Call AddtoVar(UserList(UserIndex).Stats.GLD, UserList(OtroUserIndex).ComUsu.Cant, MAXORO)
    Call SendUserStatsOro(UserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(UserIndex, Obj2) = False Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj2)
    End If
    Call QuitarObjetos(Obj2.ObjIndex, Obj2.Amount, OtroUserIndex)
End If
avis = 10
'pone el oro directamente en la billetera
If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.Cant
    Call SendUserStatsOro(UserIndex)
    'y se la doy al otro
    'UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.Cant
  Call AddtoVar(UserList(OtroUserIndex).Stats.GLD, UserList(UserIndex).ComUsu.Cant, MAXORO)

    Call SendUserStatsOro(OtroUserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(OtroUserIndex, Obj1) = False Then
        Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, Obj2)
    End If
    Call QuitarObjetos(Obj1.ObjIndex, Obj1.Amount, UserIndex)
End If
avis = 11
'[/CORREGIDO] :p

Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserInv(True, OtroUserIndex, 0)

Call FinComerciarUsu(UserIndex)
Call FinComerciarUsu(OtroUserIndex)
 
 Exit Sub
errhandler:
 Call LogError("aceptarcomerciousu " & UserList(UserIndex).Name & " y " & UserList(OtroUserIndex).Name & " Obj: " & Obj1.ObjIndex & " / " & Obj2.ObjIndex & " Cant: " & Obj1.Amount & " / " & Obj2.Amount & " " & avis)
End Sub

'[/Alejo]

