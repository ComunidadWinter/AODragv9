Attribute VB_Name = "ModNacimiento"
'pluto:2.15
Function ComprobarNombreBebe(Namebebe As String, UserIndex As Integer, Genero As String) As Boolean
'Dim Namebebe As String

'pluto:2.24
Namebebe = Trim$(Namebebe)

If Not NombrePermitido(Namebebe) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Los nombres de los personajes deben pertencer a la fantasia, el nombre indicado es invalido.")
    Exit Function
End If

'[Tite]Bug clon fichas. Se copiaba una ficha a la cuenta de papa o mama poniendo un espacio delante del nick al darle nombre.
'Dim i As Integer
'i = 1
'Do While Right$(Left$(Namebebe, i), 1) = Chr(32) And i <= Len(Namebebe)
 '   If Left$(Namebebe, 1) = Chr(32) Then
  '  Namebebe = Right$(Namebebe, Len(Namebebe) - 1)
   ' Else
    'i = i + 1
    'End If
'Loop

'[\Tite]

If Len(Namebebe) > 15 Or Len(Namebebe) < 4 Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Nombre demasiado largo o demasiado corto.")
    Exit Function
End If

If Not AsciiValidos(Namebebe) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Nombre invalido.")
    Exit Function
End If
If PersonajeExiste(Namebebe) Then
    Call SendData2(ToIndex, UserIndex, 0, 43, "Ya existe el personaje.")
    Exit Function
End If
ComprobarNombreBebe = True
Call Nacimiento(UserIndex, Namebebe, Genero)
End Function

Sub Nacimiento(UserIndex As Integer, Namebebe As String, Genero As String)

Dim Qui As Byte
Dim Dueño As String
Dim Tindex As Integer
Dim raza As String
'Dim Genero As String
Dim py As Byte
Dim px As Byte
Dim pmap As Integer

Tindex = NameIndex(UserList(UserIndex).Esposa)

Qui = RandomNumber(1, 10)

If Qui > 5 Then
Dueño = UserList(UserIndex).Email
py = UserList(UserIndex).Pos.Y
px = UserList(UserIndex).Pos.X
pmap = py = UserList(UserIndex).Pos.Map
Else
Dueño = UserList(Tindex).Email
py = UserList(Tindex).Pos.Y
px = UserList(Tindex).Pos.X
pmap = py = UserList(Tindex).Pos.Map
End If

Dim archiv As String
archiv = App.Path & "\Accounts\" & Dueño & ".acc"

'SEGUIMOS UNA VEZ SABEMOS A QUE FICHA INTRODUCIR EL PJ (DUEÑO)

'si es para la mamá que debe estar online
 If Qui > 5 Then
Cuentas(UserIndex).NumPjs = Cuentas(UserIndex).NumPjs + 1
ReDim Preserve Cuentas(UserIndex).Pj(1 To Cuentas(UserIndex).NumPjs)
Cuentas(UserIndex).Pj(Cuentas(UserIndex).NumPjs) = Namebebe
'Call MandaPersonajes(UserIndex)
Else
'si es para el papá que debe estar online
If NameIndex(UserList(UserIndex).Esposa) > 0 Then
Cuentas(Tindex).NumPjs = Cuentas(Tindex).NumPjs + 1
ReDim Preserve Cuentas(Tindex).Pj(1 To Cuentas(Tindex).NumPjs)
Cuentas(Tindex).Pj(Cuentas(Tindex).NumPjs) = Namebebe

'Call MandaPersonajes(tindex)
End If
End If
'----si no ta online
Dim Num8 As Byte
Dim num9 As Byte
UserList(UserIndex).Embarazada = 0
UserList(UserIndex).Nhijos = UserList(UserIndex).Nhijos + 1
UserList(UserIndex).Hijo(val(UserList(UserIndex).Nhijos)) = Namebebe
UserList(UserIndex).NombreDelBebe = ""
UserList(Tindex).Nhijos = UserList(Tindex).Nhijos + 1
UserList(Tindex).Hijo(val(UserList(Tindex).Nhijos)) = Namebebe
UserList(Tindex).NombreDelBebe = ""


'Dim num7 As Byte
'Num8 = val(GetVar(archiv, "DATOS", "Numpjs"))
'Call WriteVar(archiv, "DATOS", "NumPjs", CStr(Num8 + 1))
'Call WriteVar(archiv, "PERSONAJES", "PJ" & CStr(Num8 + 1), Namebebe)
'num7 = val(GetVar(archiv, "INIT", "Nhijos")) + 1
'Call WriteVar(archiv, "INIT", "Nhijos", val(num7))
'Call WriteVar(archiv, "INIT", "Hijo" & num7, Namebebe)

'aleatorios papá y mamá
Num8 = RandomNumber(1, 2)
'num9 = RandomNumber(1, 2)
If Num8 = 1 Then raza = UserList(UserIndex).raza Else raza = UserList(Tindex).raza
'If num9 = 1 Then Genero = "Hombre" Else Genero = "Mujer"
Dim pa As String
If Qui > 5 Then pa = "Madre " & UserList(UserIndex).Name Else pa = "Padre " & UserList(Tindex).Name
Call SendData(ToIndex, UserIndex, 0, "!! La matrona ha decidido en esta ocasión que la custodia del bebé corresponde a su " & pa & " que a partir de esos momentos será el encargado de su entrenamiento.")
Call SendData(ToIndex, Tindex, 0, "!! La matrona ha decidido en esta ocasión que la custodia del bebé corresponde a su " & pa & " que a partir de esos momentos será el encargado de su entrenamiento.")

Call CreaBebe(Namebebe, Dueño, raza, Genero, py, pmap, px, UserList(UserIndex).Stats.ELV, UserList(Tindex).Stats.ELV, UserList(UserIndex).raza, UserList(Tindex).raza, UserList(Tindex).Name, UserList(UserIndex).Name)
End Sub
 
 
Sub CreaBebe(Namebebe As String, Dueño As String, raza As String, Genero As String, py As Byte, pmap As Integer, px As Byte, ax3 As Byte, ax4 As Byte, ax1 As String, ax2 As String, a5 As String, a6 As String)

On Error GoTo errhandler
Dim loopc As Integer
Dim userfile As String




userfile = CharPath & Left$(Namebebe, 1) & "\" & UCase$(Namebebe) & ".chr"

Call WriteVar(userfile, "FLAGS", "Muerto", 0)
Call WriteVar(userfile, "FLAGS", "Escondido", 0)

Call WriteVar(userfile, "FLAGS", "Hambre", 0)
Call WriteVar(userfile, "FLAGS", "Sed", 0)
Call WriteVar(userfile, "FLAGS", "Desnudo", 1)
Call WriteVar(userfile, "FLAGS", "Ban", 0)
Call WriteVar(userfile, "FLAGS", "Navegando", 0)

Call WriteVar(userfile, "FLAGS", "Montura", 0)
Call WriteVar(userfile, "FLAGS", "ClaseMontura", 0)

Call WriteVar(userfile, "FLAGS", "Envenenado", 0)
Call WriteVar(userfile, "FLAGS", "Paralizado", 0)
Call WriteVar(userfile, "FLAGS", "Morph", 0)
'pluto:hoy
Call WriteVar(userfile, "QUEST", "Estado", 0)
Call WriteVar(userfile, "QUEST", "Numero", 0)
Call WriteVar(userfile, "QUEST", "Level", 0)
Call WriteVar(userfile, "QUEST", "Entrega", 0)
Call WriteVar(userfile, "QUEST", "Cantidad", 0)
Call WriteVar(userfile, "QUEST", "Objeto", 0)
Call WriteVar(userfile, "QUEST", "Enemigo", 0)
Call WriteVar(userfile, "QUEST", "Clase", "")


Call WriteVar(userfile, "FLAGS", "Angel", 0)
Call WriteVar(userfile, "FLAGS", "Demonio", 0)

Call WriteVar(userfile, "COUNTERS", "Pena", 0)

Call WriteVar(userfile, "FACCIONES", "EjercitoReal", 0)
Call WriteVar(userfile, "FACCIONES", "EjercitoCaos", 0)
Call WriteVar(userfile, "FACCIONES", "CiudMatados", 0)
Call WriteVar(userfile, "FACCIONES", "CrimMatados", 0)
Call WriteVar(userfile, "FACCIONES", "rArCaos", 0)
Call WriteVar(userfile, "FACCIONES", "rArReal", 0)

Call WriteVar(userfile, "FACCIONES", "rArLegion", 0)
Call WriteVar(userfile, "FACCIONES", "rExCaos", 0)
Call WriteVar(userfile, "FACCIONES", "rExReal", 0)
Call WriteVar(userfile, "FACCIONES", "recCaos", 0)
Call WriteVar(userfile, "FACCIONES", "recReal", 0)


Call WriteVar(userfile, "GUILD", "EsGuildLeader", 0)
Call WriteVar(userfile, "GUILD", "Echadas", 0)
Call WriteVar(userfile, "GUILD", "Solicitudes", 0)
Call WriteVar(userfile, "GUILD", "SolicitudesRechazadas", 0)
Call WriteVar(userfile, "GUILD", "VecesFueGuildLeader", 0)
Call WriteVar(userfile, "GUILD", "YaVoto", 0)
Call WriteVar(userfile, "GUILD", "FundoClan", 0)

Call WriteVar(userfile, "STATS", "PClan", 0)
Call WriteVar(userfile, "STATS", "GTorneo", 0)

Call WriteVar(userfile, "GUILD", "GuildName", "")
Call WriteVar(userfile, "GUILD", "ClanFundado", "")
Call WriteVar(userfile, "GUILD", "ClanesParticipo", "")
Call WriteVar(userfile, "GUILD", "GuildPts", "")


Dim Jur As Byte
Dim Jar As Byte
Dim Pote As Byte
'calculo potencial del bebé
'media de niveles papis
Jar = (ax3 + ax4) / 2
'suma bonus por niveles papis
Pote = 1
If Jar > 10 Then Pote = Pote + 1
If Jar > 15 Then Pote = Pote + 1
If Jar > 20 Then Pote = Pote + 1
If Jar > 25 Then Pote = Pote + 1
If Jar > 30 Then Pote = Pote + 1
If Jar > 35 Then Pote = Pote + 1
If Jar > 40 Then Pote = Pote + 1
If Jar > 45 Then Pote = Pote + 1
If Jar > 50 Then Pote = Pote + 1
If Jar > 55 Then Pote = Pote + 1

'calcula atributos
For loopc = 1 To 5
Jur = 12
'Jur = 8 + RandomNumber(1, 3)

'suma bonus por raza
Select Case UCase$(raza)
    Case "HUMANO"
        If loopc = 1 Then Jur = Jur + 2
        If loopc = 2 Then Jur = Jur + 1
        If loopc = 5 Then Jur = Jur + 2
        If loopc = 4 Then Jur = Jur + 1
    Case "ELFO"
        If loopc = 2 Then Jur = Jur + 2
        If loopc = 3 Then Jur = Jur + 1
        If loopc = 4 Then Jur = Jur + 2
    Case "ELFO OSCURO"
        If loopc = 1 Then Jur = Jur + 1
        If loopc = 2 Then Jur = Jur + 2
        If loopc = 3 Then Jur = Jur + 2
        If loopc = 4 Then Jur = Jur + 2
    Case "ENANO"
        If loopc = 1 Then Jur = Jur + 3
        If loopc = 5 Then Jur = Jur + 3
        If loopc = 3 Then Jur = Jur - 6
            
    Case "GNOMO"
           If loopc = 1 Then Jur = Jur - 5
        If loopc = 2 Then Jur = Jur + 3
        If loopc = 3 Then Jur = Jur + 3
           
    Case "ORCO"
        If loopc = 1 Then Jur = Jur + 4
        If loopc = 5 Then Jur = Jur + 3
        If loopc = 3 Then Jur = Jur - 6
        If loopc = 2 Then Jur = Jur - 2
          
   Case "VAMPIRO"
        If loopc = 1 Then Jur = Jur + 2
        If loopc = 5 Then Jur = Jur + 1
        If loopc = 3 Then Jur = Jur + 1
        If loopc = 2 Then Jur = Jur + 2
            End Select
Call WriteVar(userfile, "ATRIBUTOS", "AT" & loopc, val(Jur))
Next


For loopc = 1 To 20
Call WriteVar(userfile, "SKILLS", "SK" & loopc, 0)
Next


Call WriteVar(userfile, "CONTACTO", "Email", Dueño)
'pluto:2.10
Call WriteVar(userfile, "CONTACTO", "EmailActual", Dueño)

Call WriteVar(userfile, "INIT", "Genero", Genero)
Call WriteVar(userfile, "INIT", "Raza", raza)
'pluto:2.18------------------
Dim hog As String
Select Case UCase$(raza)
Case "HUMANO"
hog = "ALDEA DE HUMANOS"
Case "ENANO"
hog = "POBLADO ENANO"
Case "VAMPIRO"
hog = "ALDEA DE VAMPIROS"
Case "GNOMO"
hog = "ALDEA DE GNOMOS"
Case "ORCO"
hog = "POBLADO ORCO"
Case "ELFO"
hog = "ALDEA ÉLFICA"
Case "ELFO OSCURO"
hog = "ALDEA ÉLFICA"
End Select
'-----------------------------
Call WriteVar(userfile, "INIT", "Hogar", hog)
Call WriteVar(userfile, "INIT", "Clase", "Niño")
Call WriteVar(userfile, "INIT", "Desc", "Soy un bebe")
Call WriteVar(userfile, "INIT", "Heading", 3)
Call WriteVar(userfile, "INIT", "Head", 0)

If raza = "Elfo Oscuro" Or raza = "Vampiro" Then
Call WriteVar(userfile, "INIT", "Body", 342)
ElseIf raza = "Orco" Then
Call WriteVar(userfile, "INIT", "Body", 341)
Else
Call WriteVar(userfile, "INIT", "Body", 340)
End If

Call WriteVar(userfile, "INIT", "Arma", 0)
Call WriteVar(userfile, "INIT", "Escudo", 0)
Call WriteVar(userfile, "INIT", "Casco", 0)
'[GAU]
Call WriteVar(userfile, "INIT", "Botas", 0)
'[GAU]
Call WriteVar(userfile, "INIT", "RAZAREMORT", 0)
Call WriteVar(userfile, "INIT", "LastIP", "")

Call WriteVar(userfile, "INIT", "LastSerie", "")
Call WriteVar(userfile, "INIT", "LastMac", "")
Call WriteVar(userfile, "INIT", "Position", pmap & "-" & px & "-" & py)

Call WriteVar(userfile, "INIT", "Esposa", "")
Call WriteVar(userfile, "INIT", "Nhijos", 0)
For X = 1 To 5
Call WriteVar(userfile, "INIT", "Hijo" & X, "")
Next
Call WriteVar(userfile, "INIT", "Amor", 0)
Call WriteVar(userfile, "INIT", "Embarazada", 0)
Call WriteVar(userfile, "INIT", "Bebe", val(Pote))
Call WriteVar(userfile, "INIT", "NombreDelBebe", "")
Call WriteVar(userfile, "INIT", "Padre", a5)
Call WriteVar(userfile, "INIT", "Madre", a6)


Call WriteVar(userfile, "STATS", "PUNTOS", 0)

Call WriteVar(userfile, "STATS", "GLD", 0)
Call WriteVar(userfile, "STATS", "REMORT", 0)
Call WriteVar(userfile, "STATS", "BANCO", 0)

Call WriteVar(userfile, "STATS", "MET", 1)
Call WriteVar(userfile, "STATS", "MaxHP", 5)
Call WriteVar(userfile, "STATS", "MinHP", 5)

Call WriteVar(userfile, "STATS", "FIT", 10)
Call WriteVar(userfile, "STATS", "MaxSTA", 60)
Call WriteVar(userfile, "STATS", "MinSTA", 60)

Call WriteVar(userfile, "STATS", "MaxMAN", 0)
Call WriteVar(userfile, "STATS", "MinMAN", 0)

Call WriteVar(userfile, "STATS", "MaxHIT", 2)
Call WriteVar(userfile, "STATS", "MinHIT", 1)

Call WriteVar(userfile, "STATS", "MaxAGU", 100)
Call WriteVar(userfile, "STATS", "MinAGU", 100)

Call WriteVar(userfile, "STATS", "MaxHAM", 100)
Call WriteVar(userfile, "STATS", "MinHAM", 100)

Call WriteVar(userfile, "STATS", "SkillPtsLibres", 0)
  
Call WriteVar(userfile, "STATS", "EXP", 0)
Call WriteVar(userfile, "STATS", "ELV", 1)
Call WriteVar(userfile, "STATS", "ELU", 1000)
Call WriteVar(userfile, "MUERTES", "UserMuertes", 0)
Call WriteVar(userfile, "MUERTES", "CrimMuertes", 0)
Call WriteVar(userfile, "MUERTES", "NpcsMuertes", 0)
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
'pluto:7.0 quito esto no hace falta con sistema boveda en cuenta
'Call WriteVar(userfile, "BancoInventory", "CantidadItems", 0)
Dim loopd As Integer
'For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
 '   Call WriteVar(userfile, "BancoInventory", "Obj" & loopd, 0 & "-" & 0)
'Next loopd
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Call WriteVar(userfile, "Inventory", "CantidadItems", 1)
Call WriteVar(userfile, "Inventory", "Obj" & 1, 460 & "-" & 1 & "-" & 0)

For loopc = 2 To MAX_INVENTORY_SLOTS
    Call WriteVar(userfile, "Inventory", "Obj" & loopc, 0 & "-" & 0)
Next

Call WriteVar(userfile, "Inventory", "WeaponEqpSlot", 1)
Call WriteVar(userfile, "Inventory", "ArmourEqpSlot", 0)
Call WriteVar(userfile, "Inventory", "CascoEqpSlot", 0)
Call WriteVar(userfile, "Inventory", "EscudoEqpSlot", 0)
Call WriteVar(userfile, "Inventory", "BarcoSlot", 0)
Call WriteVar(userfile, "Inventory", "MunicionSlot", 0)
'pluto:2.4.1
Call WriteVar(userfile, "Inventory", "AnilloEqpSlot", 0)

'[GAU]
Call WriteVar(userfile, "Inventory", "BotaEqpSlot", 0)
'[GAU]


'Reputacion
Call WriteVar(userfile, "REP", "Asesino", 0)
Call WriteVar(userfile, "REP", "Bandido", 0)
Call WriteVar(userfile, "REP", "Burguesia", 0)
Call WriteVar(userfile, "REP", "Ladrones", 0)
Call WriteVar(userfile, "REP", "Nobles", 100)
Call WriteVar(userfile, "REP", "Plebe", 0)

Call WriteVar(userfile, "REP", "Promedio", 100)

Dim cad As String

For loopc = 1 To MAXUSERHECHIZOS
Call WriteVar(userfile, "HECHIZOS", "H" & loopc, 0)
Next




Call WriteVar(userfile, "MASCOTAS", "NroMascotas", 0)

For loopc = 1 To 3
Call WriteVar(userfile, "MONTURA" & loopc, "NIVEL", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "EXP", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "ELU", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "VIDA", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "GOLPE", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "NOMBRE", "")
Call WriteVar(userfile, "MONTURA" & loopc, "ATCUERPO", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "DEFCUERPO", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "ATFLECHAS", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "DEFFLECHAS", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "ATMAGICO", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "DEFMAGICO", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "EVASION", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "LIBRES", 0)
Call WriteVar(userfile, "MONTURA" & loopc, "TIPO", 0)

Next
Exit Sub


errhandler:
Call LogError("Error en CreaBebe")

End Sub

