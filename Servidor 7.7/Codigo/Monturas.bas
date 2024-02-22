Attribute VB_Name = "Monturas"
Sub EnviarMontura(ByVal UserIndex As Integer, ByVal MON As Byte)
On Error GoTo errhandler
Dim i As Integer
Dim cad$

'Dim xx As Integer
'xx = UserList(UserIndex).flags.ClaseMontura
'tope level
If PMascotas(MON).TopeLevel = UserList(UserIndex).Montura.Nivel(MON) Then UserList(UserIndex).Montura.Elu(MON) = 1

cad$ = UserList(UserIndex).Montura.Nivel(MON) & "," & UserList(UserIndex).Montura.exp(MON) & "," & UserList(UserIndex).Montura.Elu(MON) & "," & UserList(UserIndex).Montura.Vida(MON) & "," & UserList(UserIndex).Montura.Golpe(MON) & "," & UserList(UserIndex).Montura.Nombre(MON) & "," & str$(MON) & "," & UserList(UserIndex).Montura.AtCuerpo(MON) & "," & UserList(UserIndex).Montura.DefCuerpo(MON) & "," & UserList(UserIndex).Montura.AtFlechas(MON) & "," & UserList(UserIndex).Montura.DefFlechas(MON) & "," & UserList(UserIndex).Montura.AtMagico(MON) & "," & UserList(UserIndex).Montura.DefMagico(MON) & "," & UserList(UserIndex).Montura.Evasion(MON) & "," & UserList(UserIndex).Montura.Libres(MON)
Call SendData2(ToIndex, UserIndex, 0, 35, cad$)
 Exit Sub

errhandler:
    Call LogError("Error en EnviarMontura Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " MON:" & MON & " N: " & Err.number & " D: " & Err.Description)
    'Call LogError("Error en EnviarMontura User:" & UserIndex & " MON:" & MON)
End Sub
Sub ResetMontura(ByVal UserIndex As Integer, ByVal xx As Byte)
UserList(UserIndex).Montura.Nivel(xx) = 0
UserList(UserIndex).Montura.exp(xx) = 0
UserList(UserIndex).Montura.Elu(xx) = 0
UserList(UserIndex).Montura.Vida(xx) = 0
UserList(UserIndex).Montura.Golpe(xx) = 0
UserList(UserIndex).Montura.Nombre(xx) = ""
UserList(UserIndex).Montura.AtCuerpo(xx) = 0
UserList(UserIndex).Montura.DefCuerpo(xx) = 0
UserList(UserIndex).Montura.AtFlechas(xx) = 0
UserList(UserIndex).Montura.DefFlechas(xx) = 0
UserList(UserIndex).Montura.AtMagico(xx) = 0
UserList(UserIndex).Montura.DefMagico(xx) = 0
UserList(UserIndex).Montura.Evasion(xx) = 0
UserList(UserIndex).Montura.Tipo(xx) = 0
UserList(UserIndex).Montura.index(xx) = 0
UserList(UserIndex).Montura.Libres(xx) = 0
End Sub
Sub CheckMonturaLevel(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim xx As Integer

xx = UserList(UserIndex).flags.ClaseMontura
If xx = 0 Then Exit Sub




If PMascotas(xx).TopeLevel < UserList(UserIndex).Montura.Nivel(xx) Then Exit Sub
'Si exp >= then Exp para subir de nivel entonce subimos el nivel
If UserList(UserIndex).Montura.exp(xx) >= UserList(UserIndex).Montura.Elu(xx) Then
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_NIVEL)
    Call SendData(ToIndex, UserIndex, 0, "||¡Has subido de nivel tu Mascota !" & "´" & FontTypeNames.FONTTYPE_info)
    
    UserList(UserIndex).Montura.Nivel(xx) = UserList(UserIndex).Montura.Nivel(xx) + 1
    UserList(UserIndex).Montura.exp(xx) = 0
    UserList(UserIndex).Montura.Elu(xx) = PMascotas(xx).exp(UserList(UserIndex).Montura.Nivel(xx)) 'UserList(UserIndex).Montura.Elu(xx) * 1.5
     'pluto:6.0A
 Call SendData(ToIndex, UserIndex, 0, "H5" & xx & "," & UserList(UserIndex).Montura.Nivel(xx) & "," & UserList(UserIndex).Montura.Nombre(xx) & "," & (UserList(UserIndex).Montura.Elu(xx) - UserList(UserIndex).Montura.exp(xx)))

'PMascotas(xx).TopeLevel = UserList(UserIndex).Montura.Nivel(xx) Then
'pluto:6.0a
 If xx = 5 Or xx = 6 Then
 UserList(UserIndex).Montura.Libres(xx) = UserList(UserIndex).Montura.Libres(xx) + 4
Else
UserList(UserIndex).Montura.Libres(xx) = UserList(UserIndex).Montura.Libres(xx) + 1

End If
'pluto:2.17
Dim X As Integer
Dim Y As Integer
'Dim Expmascota As Integer
'Expmascota = UserList(UserIndex).Montura.Elu(xx) / UserList(UserIndex).Montura.Nivel(xx)
X = RandomNumber(CInt(PMascotas(xx).VidaporLevel / 2), CInt(PMascotas(xx).VidaporLevel))
Y = RandomNumber(CInt(PMascotas(xx).GolpeporLevel / 2), CInt(PMascotas(xx).GolpeporLevel))
    UserList(UserIndex).Montura.Vida(xx) = UserList(UserIndex).Montura.Vida(xx) + X
    UserList(UserIndex).Montura.Golpe(xx) = UserList(UserIndex).Montura.Golpe(xx) + Y
End If
Exit Sub
errhandler:
    LogError ("Error en la subrutina CheckMonturaLevel")
End Sub


Public Sub UsaMontura(ByVal UserIndex As Integer, ByRef Montura As ObjData)
On Error GoTo errhandler
Dim X As Integer
Dim Y As Integer

If UserList(UserIndex).Bebe > 0 Then Exit Sub
If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
If UserList(UserIndex).flags.Comerciando = True Then Exit Sub
'pluto:2.14
If UserList(UserIndex).flags.Estupidez = 1 Or UserList(UserIndex).Counters.Ceguera = 1 Then
Call SendData(ToIndex, UserIndex, 0, "||¡No puedes usar Mascotas en tu estado !" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'pluto:2.17
'If Montura.SubTipo = 5 And UserList(Userindex).Remort = 0 Then
'Call SendData(ToIndex, Userindex, 0, "||¡Mascota sólo para Remorts !" & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
'End If


If Not TieneObjetos(960, 1, UserIndex) And Montura.SubTipo <> 6 Then
Call SendData(ToIndex, UserIndex, 0, "P4")
Exit Sub
End If
'pluto:6.2
If UserList(UserIndex).Stats.ELV > 29 And Montura.SubTipo = 6 Then
Call SendData(ToIndex, UserIndex, 0, "||¡Tienes demasiado Nivel para usar un Jabato !" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'-----------


If UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Demonio > 0 Or UserList(UserIndex).flags.Muerto > 0 Then Exit Sub
If UserList(UserIndex).Pos.Map = 164 Or UserList(UserIndex).Pos.Map = 171 Or UserList(UserIndex).Pos.Map = 177 Then Exit Sub
'pluto:6.0A
If UserList(UserIndex).Pos.Map = mapi Or UserList(UserIndex).Pos.Map = 250 Then Exit Sub

If UserList(UserIndex).flags.Montura = 2 Then Exit Sub

If UserList(UserIndex).flags.Montura = 0 Then


'UserList(UserIndex).Char.Head = 0
'UserList(UserIndex).flags.DragCredito1 = 1
'pluto:6.9 dragon negro sms
If Montura.Ropaje = 306 Then
If UserList(UserIndex).flags.DragCredito1 = 1 Then Montura.Ropaje = 408
'pluto:6.5 dragon rojo sms
If UserList(UserIndex).flags.DragCredito1 = 2 Then Montura.Ropaje = 409
'pluto:6.5 dragon azul sms
If UserList(UserIndex).flags.DragCredito1 = 3 Then Montura.Ropaje = 420
'pluto:6.5 dragon violeta sms
If UserList(UserIndex).flags.DragCredito1 = 4 Then Montura.Ropaje = 421
'pluto:6.5 dragon blanco sms
If UserList(UserIndex).flags.DragCredito1 = 5 Then Montura.Ropaje = 419
End If

'pluto:6.5 uni dorado sms
If UserList(UserIndex).flags.DragCredito2 = 1 And Montura.Ropaje = 275 Then Montura.Ropaje = 422
'pluto:6.5 uni rojo sms
If UserList(UserIndex).flags.DragCredito2 = 2 And Montura.Ropaje = 275 Then Montura.Ropaje = 423
'pluto:6.9
'If UserList(UserIndex).flags.DragCredito1 = 6 Then Montura.Ropaje = 365
'------------------------
UserList(UserIndex).Char.Body = Montura.Ropaje

UserList(UserIndex).flags.ClaseMontura = Montura.SubTipo
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + (UserList(UserIndex).flags.ClaseMontura * 100)
Call SendUserStatsPeso(UserIndex)
 'pluto:6.0A
 Call SendData(ToIndex, UserIndex, 0, "H5" & Montura.SubTipo & "," & UserList(UserIndex).Montura.Nivel(Montura.SubTipo) & "," & UserList(UserIndex).Montura.Nombre(Montura.SubTipo) & "," & (UserList(UserIndex).Montura.Elu(Montura.SubTipo) - UserList(UserIndex).Montura.exp(Montura.SubTipo)))

   
   ' UserList(userindex).Char.ShieldAnim = NingunEscudo
    'UserList(userindex).Char.WeaponAnim = NingunArma
    'UserList(userindex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).Char.Botas = NingunBota
    UserList(UserIndex).flags.Montura = 1
    If UserList(UserIndex).Montura.Nivel(Montura.SubTipo) = 1 Then
    'UserList(UserIndex).flags.Estupidez = 1
    Call SendData2(ToIndex, UserIndex, 0, 3)
    End If
Else '<>montura=0
    UserList(UserIndex).flags.Estupidez = 0
    Call SendData2(ToIndex, UserIndex, 0, 56)
    UserList(UserIndex).flags.Montura = 0
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax - (UserList(UserIndex).flags.ClaseMontura * 100)
Call SendUserStatsPeso(UserIndex)
'pluto:6.0A
 Call SendData(ToIndex, UserIndex, 0, "H7")

    
    If UserList(UserIndex).flags.Muerto = 0 Then
    UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
            
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
        If UserList(UserIndex).Invent.BotaEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.Botas = ObjData(UserList(UserIndex).Invent.BotaEqpObjIndex).Botas

    
    Else 'muerto
        If Not Criminal(UserIndex) Then UserList(UserIndex).Char.Body = iCuerpoMuerto Else UserList(UserIndex).Char.Body = iCuerpoMuerto2
        If Not Criminal(UserIndex) Then UserList(UserIndex).Char.Head = iCabezaMuerto Else UserList(UserIndex).Char.Head = iCabezaMuerto2
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
        UserList(UserIndex).Char.Botas = NingunBota
       
    End If

End If

Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)

'pluto:6.0A silueta mascota
  'If UserInventory(iX).OBJType = 60 Then
  'frmMain.LogoMascota.Picture = LoadPicture(App.Path & "\graficos\" & val(UserInventory(iX).SubTipo) & ".jpg")
  'frmMain.LogoMascota.Visible = True
  'End If
  '----------------------------
  
  

'Call SendData(ToIndex, UserIndex, 0, "NAVEG")
 Exit Sub

errhandler:
    Call LogError("Error en UsaMontura")
End Sub

Sub DarMontura(ByVal UserIndex As Integer, ByVal rdata As String)
On Error GoTo errhandler
Dim userfile As String
Dim userfile2 As String
Dim Name As String

'pluto:6.3
If rdata = "" Then Exit Sub
Name = rdata & "$"
Tindex = NameIndex(Name)
    If Tindex > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "|| Ese usuario está Online, usa el /comerciar para pasarle la mascota." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
    End If

userfile = CharPath & Left$(rdata, 1) & "\" & rdata & ".chr"
userfile2 = CharPath & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".chr"



'modifica ficha
If FileExist(userfile, vbArchive) Then 'And FileExist(userfile2, vbArchive) Then


Dim x1 As Byte
Dim x2 As Long
Dim x3 As Long
Dim x4 As Integer
Dim x5 As Integer
Dim x6 As String
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

xx = UserList(UserIndex).flags.ClaseMontura

Dim xxx As Byte 'index de la mascota 1 a 3

   'buscamos un hueco
For n = 1 To 3
If val(GetVar(userfile, "MONTURA" & n, "TIPO")) = 0 Then
xxx = n 'index de la mascota 1 a 3
Exit For
End If
Next n


    'salimos sin hueco
If xxx = 0 Then
Call SendData(ToIndex, UserIndex, 0, "||Ese Pj ya tiene el tope de Mascotas" & "´" & FontTypeNames.FONTTYPE_info)

Exit Sub
End If


'pluto:6.0A
If val(GetVar(userfile, "INIT", "Bebe")) > 0 Then
Call SendData(ToIndex, UserIndex, 0, "||Los bebes no usan mascotas." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

'miramos que no repita mascota
For n = 1 To 3
If val(GetVar(userfile, "MONTURA" & n, "TIPO")) = xx Then
Call SendData(ToIndex, UserIndex, 0, "||Ese Personaje ya tiene esa clase de mascota.." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
Next n

'Call LogMascotas("Dar mascota: " & " Metemos Tipo " & xx & " EN INDEX " & xxx & " del user " & rdata)

'carga en las variables Xn las caracteristicas de la montura
x1 = UserList(UserIndex).Montura.Nivel(xx)
x2 = UserList(UserIndex).Montura.exp(xx)
x3 = UserList(UserIndex).Montura.Elu(xx)
x4 = UserList(UserIndex).Montura.Vida(xx)
x5 = UserList(UserIndex).Montura.Golpe(xx)
x6 = UserList(UserIndex).Montura.Nombre(xx)
x7 = UserList(UserIndex).Montura.AtCuerpo(xx)
x8 = UserList(UserIndex).Montura.DefCuerpo(xx)
x9 = UserList(UserIndex).Montura.AtFlechas(xx)
x10 = UserList(UserIndex).Montura.DefFlechas(xx)
x11 = UserList(UserIndex).Montura.AtMagico(xx)
x12 = UserList(UserIndex).Montura.DefMagico(xx)
x13 = UserList(UserIndex).Montura.Evasion(xx)
x14 = UserList(UserIndex).Montura.Libres(xx)
x15 = UserList(UserIndex).Montura.Tipo(xx)
'Graba en la ficha del Pj receptor la mascota con sus caracteristicas
Call WriteVar(userfile, "MONTURA" & xxx, "NIVEL", val(x1))
Call WriteVar(userfile, "MONTURA" & xxx, "EXP", val(x2))
Call WriteVar(userfile, "MONTURA" & xxx, "ELU", val(x3))
Call WriteVar(userfile, "MONTURA" & xxx, "VIDA", val(x4))
Call WriteVar(userfile, "MONTURA" & xxx, "GOLPE", val(x5))
Call WriteVar(userfile, "MONTURA" & xxx, "NOMBRE", x6)
Call WriteVar(userfile, "MONTURA" & xxx, "ATCUERPO", val(x7))
Call WriteVar(userfile, "MONTURA" & xxx, "DEFCUERPO", val(x8))
Call WriteVar(userfile, "MONTURA" & xxx, "ATFLECHAS", val(x9))
Call WriteVar(userfile, "MONTURA" & xxx, "DEFFLECHAS", val(x10))
Call WriteVar(userfile, "MONTURA" & xxx, "ATMAGICO", val(x11))
Call WriteVar(userfile, "MONTURA" & xxx, "DEFMAGICO", val(x12))
Call WriteVar(userfile, "MONTURA" & xxx, "EVASION", val(x13))
Call WriteVar(userfile, "MONTURA" & xxx, "LIBRES", val(x14))
Call WriteVar(userfile, "MONTURA" & xxx, "TIPO", val(x15))
Dim Nmascorecep As Byte
Call LogMascotas("Dar mascota: " & UserList(UserIndex).Name & " da su " & x6 & " a " & rdata & " EN INDEX " & xxx)


Nmascorecep = val(GetVar(userfile, "MONTURAS", "NroMonturas"))
Call LogMascotas("Dar mascota: " & rdata & " tenia " & Nmascorecep)
Nmascorecep = Nmascorecep + 1
Call WriteVar(userfile, "MONTURAS", "NroMonturas", val(Nmascorecep))
Call LogMascotas("Dar mascota: " & rdata & " ahora tiene " & Nmascorecep)


'Elimina la mascota del registro del dueño original
UserList(UserIndex).Montura.Nivel(xx) = 0
UserList(UserIndex).Montura.exp(xx) = 0
UserList(UserIndex).Montura.Elu(xx) = 0
UserList(UserIndex).Montura.Vida(xx) = 0
UserList(UserIndex).Montura.Golpe(xx) = 0
UserList(UserIndex).Montura.Nombre(xx) = ""
UserList(UserIndex).Montura.AtCuerpo(xx) = 0
UserList(UserIndex).Montura.AtFlechas(xx) = 0
UserList(UserIndex).Montura.AtMagico(xx) = 0
UserList(UserIndex).Montura.DefCuerpo(xx) = 0
UserList(UserIndex).Montura.DefFlechas(xx) = 0
UserList(UserIndex).Montura.DefMagico(xx) = 0
UserList(UserIndex).Montura.Evasion(xx) = 0
UserList(UserIndex).Montura.Tipo(xx) = 0
UserList(UserIndex).Montura.Libres(xx) = 0
UserList(UserIndex).Montura.index(xx) = 0
'UserFile = App.Path & "\charfile\" & UCase$(UserList(UserIndex).name) & ".chr"

'Elimina la mascota de la ficha del dueño original
For n = 1 To 3
If val(GetVar(userfile2, "MONTURA" & n, "TIPO")) = xx Then
zzz = n 'index mascota 1-3 dueño
End If
Next n
Call LogMascotas("Dar mascota: " & UserList(UserIndex).Name & " CERO EN INDEX " & zzz)

Call WriteVar(userfile2, "MONTURA" & zzz, "NIVEL", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "EXP", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "ELU", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "VIDA", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "GOLPE", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "NOMBRE", "")
Call WriteVar(userfile2, "MONTURA" & zzz, "ATCUERPO", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "DEFCUERPO", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "ATFLECHAS", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "DEFFLECHAS", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "ATMAGICO", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "DEFMAGICO", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "EVASION", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "LIBRES", 0)
Call WriteVar(userfile2, "MONTURA" & zzz, "TIPO", 0)

Call QuitarObjetos(UserList(UserIndex).flags.ClaseMontura + 887, 1, UserIndex)
Call LogMascotas("Dar mascota: " & UserList(UserIndex).Name & " quitar objeto " & UserList(UserIndex).flags.ClaseMontura + 887)


'quita
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
            Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
            Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
            UserList(UserIndex).MascotasIndex(i) = 0
            UserList(UserIndex).MascotasType(i) = 0
        End If
    End If
Next i
UserList(UserIndex).Nmonturas = UserList(UserIndex).Nmonturas - 1
UserList(UserIndex).flags.Montura = 0
Call WriteVar(userfile2, "MONTURAS", "NroMonturas", val(UserList(UserIndex).Nmonturas))
Call LogMascotas("Dar mascota: " & UserList(UserIndex).Name & " ahora tiene " & UserList(UserIndex).Nmonturas)
Call SendData(ToIndex, UserIndex, 0, "||La Mascota ha sido enviada a " & rdata & "´" & FontTypeNames.FONTTYPE_info)

'si esta online
'[Tite]Soluciona el bug que duplicaba mascotas
'Name = rdata
'Name = rdata & "$"
'[\Tite]
'If Name = "" Then Exit Sub
 '   Tindex = NameIndex(Name)
'If Tindex <= 0 Then GoTo yap
'UserList(Tindex).Montura.Nivel(xx) = val(x1)
'UserList(Tindex).Montura.exp(xx) = val(x2)
'UserList(Tindex).Montura.Elu(xx) = val(x3)
'UserList(Tindex).Montura.Vida(xx) = val(x4)
'UserList(Tindex).Montura.Golpe(xx) = val(x5)
'UserList(Tindex).Montura.Nombre(xx) = x6
'UserList(Tindex).Montura.AtCuerpo(xx) = val(x7)
'UserList(Tindex).Montura.DefCuerpo(xx) = val(x8)
'UserList(Tindex).Montura.AtFlechas(xx) = val(x9)
'UserList(Tindex).Montura.DefFlechas(xx) = val(x10)
'UserList(Tindex).Montura.AtMagico(xx) = val(x11)
'UserList(Tindex).Montura.DefMagico(xx) = val(x12)
'UserList(Tindex).Montura.Evasion(xx) = val(x13)
'UserList(Tindex).Montura.Libres(xx) = val(x14)
'UserList(Tindex).Montura.Tipo(xx) = val(x15)
'UserList(Tindex).Montura.index(xx) = zzz
'UserList(Tindex).Nmonturas = UserList(Tindex).Nmonturas + 1
yap:
Else
Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe" & "´" & FontTypeNames.FONTTYPE_info)
End If
 Exit Sub

errhandler:
    Call LogError("Error en DarMontura: " & UserList(UserIndex).Name)
'End If
End Sub
Sub DomarMontura(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
On Error GoTo errhandler

Dim n As Byte
Dim tc As Integer

Dim userfile As String



userfile = CharPath & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".chr"

tc = Npclist(NpcIndex).flags.Domable + 387
'If Npclist(NpcIndex).Numero < 621 Then
'tc = Npclist(NpcIndex).Numero + 272
'Else
'tc = Npclist(npcinde).Numero + 224
'End If

Dim nPos As WorldPos
Dim MiObj As obj
MiObj.Amount = 1
MiObj.ObjIndex = tc
If TieneObjetos(tc, 1, UserIndex) Then
NoDomarMontura = True
Exit Sub
End If
'miramos que no repita mascota
For n = 1 To 3
If val(GetVar(userfile, "MONTURA" & n, "TIPO")) = Npclist(NpcIndex).flags.Domable - 500 Then
Call SendData(ToIndex, UserIndex, 0, "||Ya tienes esa clase de mascota, ve a la cuidadora de mascotas en Banderbill a recuperarla." & "´" & FontTypeNames.FONTTYPE_info)
NoDomarMontura = True
Exit Sub
End If
Next n



'Dim K As Integer
'K = RandomNumber(1, 1000)
'If Npclist(npcindex).Flags.Domable <= 1000 Then 'CalcularPoderDomador(UserIndex) And K > 500 Then
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call SendData(ToIndex, UserIndex, 0, "P5")
NoDomarMontura = True
Exit Sub
End If
'pluto:6.5
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'---------------------------


Call SendData(ToIndex, UserIndex, 0, "||La criatura te ha aceptado como su amo." & "´" & FontTypeNames.FONTTYPE_info)
Call LogMascotas("Domar: " & UserList(UserIndex).Name & " doma un " & Npclist(NpcIndex).Name)
Call SubirSkill(UserIndex, Domar)
Call QuitarNPC(NpcIndex)
Dim xx As Integer
Dim X As Integer
Dim Y As Integer

Dim Expmascota As Integer
xx = tc - 887


X = RandomNumber(CInt(PMascotas(xx).VidaporLevel / 2), PMascotas(xx).VidaporLevel)
Y = RandomNumber(CInt(PMascotas(xx).GolpeporLevel / 2), PMascotas(xx).GolpeporLevel)

UserList(UserIndex).Montura.Nivel(xx) = 1
UserList(UserIndex).Montura.exp(xx) = 0
UserList(UserIndex).Montura.Elu(xx) = PMascotas(xx).exp(1)
UserList(UserIndex).Montura.Vida(xx) = X
UserList(UserIndex).Montura.Golpe(xx) = Y
UserList(UserIndex).Montura.Nombre(xx) = PMascotas(xx).Tipo
'pluto:6.0A
UserList(UserIndex).Montura.AtCuerpo(xx) = 0
UserList(UserIndex).Montura.DefCuerpo(xx) = 0
UserList(UserIndex).Montura.AtFlechas(xx) = 0
UserList(UserIndex).Montura.DefFlechas(xx) = 0
UserList(UserIndex).Montura.AtMagico(xx) = 0
UserList(UserIndex).Montura.DefMagico(xx) = 0
UserList(UserIndex).Montura.Evasion(xx) = 0

'pluto:6.3
If xx = 5 Then
 UserList(UserIndex).Montura.Libres(xx) = UserList(UserIndex).Montura.Libres(xx) + 3
ElseIf xx = 6 Then
UserList(UserIndex).Montura.Libres(xx) = UserList(UserIndex).Montura.Libres(xx) + 4
Else
UserList(UserIndex).Montura.Libres(xx) = UserList(UserIndex).Montura.Libres(xx) + 1
End If

'If xx <> 5 Then
'UserList(UserIndex).Montura.Libres(xx) = 4
'Else
'UserList(UserIndex).Montura.Libres(xx) = 3
'End If

UserList(UserIndex).Montura.Tipo(xx) = xx
UserList(UserIndex).Nmonturas = UserList(UserIndex).Nmonturas + 1
Dim xxx As Byte
For n = 1 To 3
    If val(GetVar(userfile, "MONTURA" & n, "TIPO")) = 0 Then
    xxx = n
    Exit For
    End If
Next n
Call WriteVar(userfile, "MONTURAS", "NroMonturas", val(UserList(UserIndex).Nmonturas))
Call LogMascotas("Domar: " & UserList(UserIndex).Name & " ahora tiene " & UserList(UserIndex).Nmonturas & " la metemos en index " & xxx)

Call WriteVar(userfile, "MONTURA" & xxx, "NOMBRE", UserList(UserIndex).Montura.Nombre(xx))
Call WriteVar(userfile, "MONTURA" & xxx, "NIVEL", val(UserList(UserIndex).Montura.Nivel(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "EXP", val(UserList(UserIndex).Montura.exp(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "ELU", val(UserList(UserIndex).Montura.Elu(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "VIDA", val(UserList(UserIndex).Montura.Vida(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "GOLPE", val(UserList(UserIndex).Montura.Golpe(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "TIPO", val(UserList(UserIndex).Montura.Tipo(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "ATCUERPO", val(UserList(UserIndex).Montura.AtCuerpo(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "DEFCUERPO", val(UserList(UserIndex).Montura.DefCuerpo(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "ATFLECHAS", val(UserList(UserIndex).Montura.AtFlechas(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "DEFFLECHAS", val(UserList(UserIndex).Montura.DefFlechas(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "ATMAGICO", val(UserList(UserIndex).Montura.AtMagico(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "DEFMAGICO", val(UserList(UserIndex).Montura.DefMagico(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "EVASION", val(UserList(UserIndex).Montura.Evasion(xx)))
Call WriteVar(userfile, "MONTURA" & xxx, "LIBRES", val(UserList(UserIndex).Montura.Libres(xx)))
'pluto:6.0A
UserList(UserIndex).Montura.index(xx) = xxx


'If Not MeterItemEnInventario(userindex, MiObj) Then
'Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
'End If

'Else
'Call SendData(ToIndex, UserIndex, 0, "P3")
'End If
'fin pluto:2.3
 Exit Sub

errhandler:
    Call LogError("Error en DomarMontura")

End Sub

Sub MontarSoltar(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo errhandler
'pluto:2.3
If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = 60 Then

'pluto:2.17
'If ObjData(UserList(Userindex).Invent.Object(Slot).ObjIndex).SubTipo = 5 And UserList(Userindex).Remort = 0 Then
'Call SendData(ToIndex, Userindex, 0, "||¡Mascota sólo para Remorts !" & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
'End If
'--------------------------

'pluto:6.0A
If UserList(UserIndex).flags.Muerto = 1 Or UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
If MapInfo(UserList(UserIndex).Pos.Map).Monturas = 1 Then Exit Sub
'pluto:6.8
If UserList(UserIndex).Bebe > 0 Then Exit Sub
'ropa cabalgar y no jabato
If Not TieneObjetos(960, 1, UserIndex) And ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo <> 6 Then
Call SendData(ToIndex, UserIndex, 0, "P4")
Exit Sub
End If

'pluto:6.9
If UserList(UserIndex).Stats.ELV > 29 And ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo = 6 Then
Call SendData(ToIndex, UserIndex, 0, "||¡Tienes demasiado Nivel para usar un Jabato !" & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'-----------


If UserList(UserIndex).flags.Montura = 2 Then
Dim a As Integer
a = UserList(UserIndex).Stats.Peso

Dim i As Integer
For i = 1 To MAXMASCOTAS
If UserList(UserIndex).MascotasIndex(i) > 0 Then
If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
UserList(UserIndex).MascotasIndex(i) = 0
UserList(UserIndex).MascotasType(i) = 0
End If
End If
Next i
UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
UserList(UserIndex).flags.Montura = 0
'UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax - (UserList(UserIndex).Flags.ClaseMontura * 100)
UserList(UserIndex).flags.ClaseMontura = 0
Call UseInvItem(UserIndex, Slot)

Exit Sub
End If

Dim ind As Integer, index As Integer
If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
ind = SpawnNpc(ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Clave, UserList(UserIndex).Pos, False, False)


If ind = MAXNPCS Then
Call SendData(ToIndex, UserIndex, 0, "||No hay sitio para tu mascota." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
        
        If ind < MAXNPCS Then
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            
            index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(index) = ind
            UserList(UserIndex).MascotasType(index) = Npclist(ind).numero
            
            Npclist(ind).MaestroUser = UserIndex
           
           If UserList(UserIndex).flags.Montura = 1 Then Call UseInvItem(UserIndex, Slot)
           UserList(UserIndex).flags.ClaseMontura = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo
           UserList(UserIndex).flags.Montura = 2
        
            Npclist(ind).Stats.MinHP = UserList(UserIndex).Montura.Vida(UserList(UserIndex).flags.ClaseMontura)
            Npclist(ind).Stats.MaxHP = UserList(UserIndex).Montura.Vida(UserList(UserIndex).flags.ClaseMontura)
'pluto:2.4
'UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + (UserList(UserIndex).Flags.ClaseMontura * 100)

            Call FollowAmo(ind)
    Else
        Exit Sub
    End If
End If
 Exit Sub
 End If
' pluto:2.3
 Exit Sub

errhandler:
    Call LogError("Error en MontarSoltar")

End Sub

