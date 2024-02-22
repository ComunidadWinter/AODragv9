Attribute VB_Name = "ES"
Option Explicit

Public Sub CargarSpawnList()
  On Error GoTo fallo
    Dim n As Integer, loopc As Integer
    n = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(n) As tCriaturasEntrenador
    For loopc = 1 To n
        SpawnList(loopc).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & loopc))
        SpawnList(loopc).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & loopc)
    Next loopc
Exit Sub
fallo:
Call LogError("CARGARSPAWNLIST" & Err.number & " D: " & Err.Description)

End Sub

Function EsDios(ByVal Name As String) As Boolean
On Error GoTo fallo
Dim NumWizs As Integer
Dim WizNum As Integer
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    If UCase$(Name) = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum)) Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False

Exit Function
fallo:
Call LogError("ESDIOS" & Err.number & " D: " & Err.Description)


End Function

Function EsSemiDios(ByVal Name As String) As Boolean
On Error GoTo fallo
Dim NumWizs As Integer
Dim WizNum As Integer
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    If UCase$(Name) = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum)) Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False

Exit Function
fallo:
Call LogError("ESSEMIDIOS" & Err.number & " D: " & Err.Description)


End Function

Function EsConsejero(ByVal Name As String) As Boolean
On Error GoTo fallo
Dim NumWizs As Integer
Dim WizNum As Integer
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    If UCase$(Name) = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum)) Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False

Exit Function
fallo:
Call LogError("ESCONSEJERO" & Err.number & " D: " & Err.Description)


End Function
Public Function TxtDimension(ByVal Name As String) As Long
On Error GoTo fallo
Dim n As Integer, cad As String, Tam As Long
n = FreeFile(1)
Open Name For Input As #n
Tam = 0
Do While Not EOF(n)
    Tam = Tam + 1
    Line Input #n, cad
Loop
Close n
TxtDimension = Tam

Exit Function
fallo:
Call LogError("TXTDIMENSION" & Err.number & " D: " & Err.Description)

End Function

Public Sub CargarForbidenWords()
On Error GoTo fallo
ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim n As Integer, i As Integer
n = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #n

For i = 1 To UBound(ForbidenNames)
    Line Input #n, ForbidenNames(i)
Next i

Close n
Exit Sub
fallo:
Call LogError("CAGARFORBIDENWORDS" & Err.number & " D: " & Err.Description)

End Sub
Public Sub CargarHechizos()
On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

Dim Hechizo As Integer

'pluto fusión
Dim Leer As New clsLeerInis
Leer.Abrir DatPath & "Hechizos.dat"

'obtiene el numero de hechizos
NumeroHechizos = val(Leer.DarValor("INIT", "NumeroHechizos"))
'NumeroHechizos = val(GetVar(DatPath & "Hechizos.dat", "INIT", "NumeroHechizos"))
ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos
    frmCargando.Label1(2).Caption = "Hechizo: (" & Hechizo & "/" & NumeroHechizos & ")"
    
    Hechizos(Hechizo).Nombre = Leer.DarValor("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).Desc = Leer.DarValor("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.DarValor("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = Leer.DarValor("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.DarValor("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.DarValor("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = val(Leer.DarValor("Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).WAV = val(Leer.DarValor("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = val(Leer.DarValor("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = val(Leer.DarValor("Hechizo" & Hechizo, "Loops"))
    
    Hechizos(Hechizo).Resis = val(Leer.DarValor("Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = val(Leer.DarValor("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(Leer.DarValor("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(Leer.DarValor("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = val(Leer.DarValor("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = val(Leer.DarValor("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = val(Leer.DarValor("Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = val(Leer.DarValor("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = val(Leer.DarValor("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = val(Leer.DarValor("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).SubeHam = val(Leer.DarValor("Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = val(Leer.DarValor("Hechizo" & Hechizo, "MinHam"))
    Hechizos(Hechizo).MaxHam = val(Leer.DarValor("Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = val(Leer.DarValor("Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = val(Leer.DarValor("Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = val(Leer.DarValor("Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = val(Leer.DarValor("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(Leer.DarValor("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Paralizaarea = val(Leer.DarValor("Hechizo" & Hechizo, "Paralizaarea"))

    'Hechizos(Hechizo).Inmoviliza = val(leer.darvalor("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverParalisis"))
    'Hechizos(Hechizo).RemoverEstupidez = val(leer.darvalor("Hechizo" & Hechizo, "RemoverEstupidez"))
    'Hechizos(Hechizo).RemueveInvisibilidadParcial = val(leer.darvalor("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
    
    Hechizos(Hechizo).CuraVeneno = val(Leer.DarValor("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(Leer.DarValor("Hechizo" & Hechizo, "Envenena"))
    'pluto:2.15
    Hechizos(Hechizo).Protec = val(Leer.DarValor("Hechizo" & Hechizo, "Protec"))

    Hechizos(Hechizo).Maldicion = val(Leer.DarValor("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = val(Leer.DarValor("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = val(Leer.DarValor("Hechizo" & Hechizo, "Revivir"))
    Hechizos(Hechizo).Morph = val(Leer.DarValor("Hechizo" & Hechizo, "Morph"))

    Hechizos(Hechizo).Ceguera = val(Leer.DarValor("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = val(Leer.DarValor("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).invoca = val(Leer.DarValor("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = val(Leer.DarValor("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).Cant = val(Leer.DarValor("Hechizo" & Hechizo, "Cant"))
    'Hechizos(Hechizo).Mimetiza = val(leer.darvalor("hechizo" & Hechizo, "Mimetiza"))
    
    
    Hechizos(Hechizo).MinNivel = val(Leer.DarValor("Hechizo" & Hechizo, "MinNivel"))
    Hechizos(Hechizo).itemIndex = val(Leer.DarValor("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = val(Leer.DarValor("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(Leer.DarValor("Hechizo" & Hechizo, "ManaRequerido"))
    Hechizos(Hechizo).Target = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Target"))
    frmCargando.cargar.value = frmCargando.cargar.value + 1
    'DoEvents
Next

'quitar esto
Exit Sub
'------------------------------------------------------------------------------------
'Esto genera el hechizos.log para meterlo al cliente, el server no usa nada de lo de abajo.
'------------------------------------------------------------------------------------
Dim file As String
Dim n As Byte
Dim Object As Integer
file = DatPath & "Hechizos.dat"
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\Hechizo.log" For Append Shared As #nfile

For Object = 1 To NumeroHechizos
Debug.Print Object
Print #nfile, "hechizos(" & Object & ").nombre=" & Chr(34) & Hechizos(Object).Nombre & Chr(34)
Print #nfile, "hechizos(" & Object & ").desc=" & Chr(34) & Hechizos(Object).Desc & Chr(34)
Print #nfile, "hechizos(" & Object & ").palabrasmagicas=" & Chr(34) & Hechizos(Object).PalabrasMagicas & Chr(34)
Print #nfile, "hechizos(" & Object & ").hechizeromsg=" & Chr(34) & Hechizos(Object).HechizeroMsg & Chr(34)
Print #nfile, "hechizos(" & Object & ").propiomsg=" & Chr(34) & Hechizos(Object).PropioMsg & Chr(34)
Print #nfile, "hechizos(" & Object & ").targetmsg=" & Chr(34) & Hechizos(Object).TargetMsg & Chr(34)
If Hechizos(Object).Bendicion > 0 Then Print #nfile, "hechizos(" & Object & ").bendicion =" & Hechizos(Object).Bendicion
If Hechizos(Object).Cant > 0 Then Print #nfile, "hechizos(" & Object & ").cant =" & Hechizos(Object).Cant
If Hechizos(Object).Ceguera > 0 Then Print #nfile, "hechizos(" & Object & ").ceguera =" & Hechizos(Object).Ceguera
If Hechizos(Object).CuraVeneno > 0 Then Print #nfile, "hechizos(" & Object & ").curaveneno =" & Hechizos(Object).CuraVeneno
If Hechizos(Object).Envenena > 0 Then Print #nfile, "hechizos(" & Object & ").envenena =" & Hechizos(Object).Envenena
If Hechizos(Object).Estupidez > 0 Then Print #nfile, "hechizos(" & Object & ").estupidez =" & Hechizos(Object).Estupidez
If Hechizos(Object).FXgrh > 0 Then Print #nfile, "hechizos(" & Object & ").fxgrh =" & Hechizos(Object).FXgrh
If Hechizos(Object).Invisibilidad > 0 Then Print #nfile, "hechizos(" & Object & ").invisibilidad =" & Hechizos(Object).Invisibilidad
If Hechizos(Object).invoca > 0 Then Print #nfile, "hechizos(" & Object & ").invoca =" & Hechizos(Object).invoca
If Hechizos(Object).itemIndex > 0 Then Print #nfile, "hechizos(" & Object & ").itemindex =" & Hechizos(Object).itemIndex
If Hechizos(Object).loops > 0 Then Print #nfile, "hechizos(" & Object & ").loops =" & Hechizos(Object).loops
If Hechizos(Object).Maldicion > 0 Then Print #nfile, "hechizos(" & Object & ").maldicion  =" & Hechizos(Object).Maldicion
If Hechizos(Object).MaMana > 0 Then Print #nfile, "hechizos(" & Object & ").mamana=" & Hechizos(Object).MaMana
If Hechizos(Object).ManaRequerido > 0 Then Print #nfile, "hechizos(" & Object & ").ManaRequerido =" & Hechizos(Object).ManaRequerido
If Hechizos(Object).MaxAgilidad > 0 Then Print #nfile, "hechizos(" & Object & ").maxagilidad =" & Hechizos(Object).MaxAgilidad
If Hechizos(Object).MaxCarisma > 0 Then Print #nfile, "hechizos(" & Object & ").Maxcarisma =" & Hechizos(Object).MaxCarisma
If Hechizos(Object).MaxFuerza > 0 Then Print #nfile, "hechizos(" & Object & ").Maxfuerza =" & Hechizos(Object).MaxFuerza
If Hechizos(Object).MaxHam > 0 Then Print #nfile, "hechizos(" & Object & ").maxham =" & Hechizos(Object).MaxHam
If Hechizos(Object).MaxHP > 0 Then Print #nfile, "hechizos(" & Object & ").Maxhp =" & Hechizos(Object).MaxHP
If Hechizos(Object).MaxSed > 0 Then Print #nfile, "hechizos(" & Object & ").Maxsed =" & Hechizos(Object).MaxSed
If Hechizos(Object).MaxSta > 0 Then Print #nfile, "hechizos(" & Object & ").Maxsta =" & Hechizos(Object).MaxSta
If Hechizos(Object).MiMana > 0 Then Print #nfile, "hechizos(" & Object & ").Mimana =" & Hechizos(Object).MiMana
If Hechizos(Object).MinAgilidad > 0 Then Print #nfile, "hechizos(" & Object & ").minagilidad =" & Hechizos(Object).MinAgilidad
If Hechizos(Object).MinCarisma > 0 Then Print #nfile, "hechizos(" & Object & ").mincarisma =" & Hechizos(Object).MinCarisma
If Hechizos(Object).MinFuerza > 0 Then Print #nfile, "hechizos(" & Object & ").Minfuerza =" & Hechizos(Object).MinFuerza
If Hechizos(Object).MinHam > 0 Then Print #nfile, "hechizos(" & Object & ").Minham =" & Hechizos(Object).MinHam
If Hechizos(Object).MinHP > 0 Then Print #nfile, "hechizos(" & Object & ").Minhp =" & Hechizos(Object).MinHP
If Hechizos(Object).MinSed > 0 Then Print #nfile, "hechizos(" & Object & ").minsed =" & Hechizos(Object).MinSed
If Hechizos(Object).MinSkill > 0 Then Print #nfile, "hechizos(" & Object & ").minskill =" & Hechizos(Object).MinSkill
If Hechizos(Object).MinSta > 0 Then Print #nfile, "hechizos(" & Object & ").Minsta =" & Hechizos(Object).MinSta
If Hechizos(Object).Morph > 0 Then Print #nfile, "hechizos(" & Object & ").morph =" & Hechizos(Object).Morph
If Hechizos(Object).MinNivel > 0 Then Print #nfile, "hechizos(" & Object & ").MinNivel =" & Hechizos(Object).MinNivel
If Hechizos(Object).NumNpc > 0 Then Print #nfile, "hechizos(" & Object & ").numnpc =" & Hechizos(Object).NumNpc
If Hechizos(Object).Paraliza > 0 Then Print #nfile, "hechizos(" & Object & ").paraliza =" & Hechizos(Object).Paraliza
If Hechizos(Object).Paralizaarea > 0 Then Print #nfile, "hechizos(" & Object & ").paralizaarea =" & Hechizos(Object).Paralizaarea
If Hechizos(Object).Protec > 0 Then Print #nfile, "hechizos(" & Object & ").protec =" & Hechizos(Object).Protec
If Hechizos(Object).RemoverMaldicion > 0 Then Print #nfile, "hechizos(" & Object & ").removermaldicion =" & Hechizos(Object).RemoverMaldicion
If Hechizos(Object).RemoverParalisis > 0 Then Print #nfile, "hechizos(" & Object & ").removerparalisis =" & Hechizos(Object).RemoverParalisis
If Hechizos(Object).Resis > 0 Then Print #nfile, "hechizos(" & Object & ").resis =" & Hechizos(Object).Resis
If Hechizos(Object).Revivir > 0 Then Print #nfile, "hechizos(" & Object & ").revivir =" & Hechizos(Object).Revivir
If Hechizos(Object).SubeAgilidad > 0 Then Print #nfile, "hechizos(" & Object & ").subeagilidad =" & Hechizos(Object).SubeAgilidad
If Hechizos(Object).SubeCarisma > 0 Then Print #nfile, "hechizos(" & Object & ").subecarisma =" & Hechizos(Object).SubeCarisma
If Hechizos(Object).SubeFuerza > 0 Then Print #nfile, "hechizos(" & Object & ").subefuerza=" & Hechizos(Object).SubeFuerza
If Hechizos(Object).SubeHam > 0 Then Print #nfile, "hechizos(" & Object & ").subeham =" & Hechizos(Object).SubeHam
If Hechizos(Object).SubeHP > 0 Then Print #nfile, "hechizos(" & Object & ").subehp =" & Hechizos(Object).SubeHP
If Hechizos(Object).SubeMana > 0 Then Print #nfile, "hechizos(" & Object & ").subemana =" & Hechizos(Object).SubeMana
If Hechizos(Object).SubeSed > 0 Then Print #nfile, "hechizos(" & Object & ").subesed =" & Hechizos(Object).SubeSed
If Hechizos(Object).SubeSta > 0 Then Print #nfile, "hechizos(" & Object & ").subesta =" & Hechizos(Object).SubeSta
If Hechizos(Object).Target > 0 Then Print #nfile, "hechizos(" & Object & ").target =" & Hechizos(Object).Target
If Hechizos(Object).Tipo > 0 Then Print #nfile, "hechizos(" & Object & ").tipo =" & Hechizos(Object).Tipo
If Hechizos(Object).WAV > 0 Then Print #nfile, "hechizos(" & Object & ").wav =" & Hechizos(Object).WAV

Next
Close #nfile
Exit Sub
errhandler:
 MsgBox "Error cargando hechizos.dat"
End Sub

Sub LoadMotd()
On Error GoTo fallo
Dim i As Integer
MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))
ReDim MOTD(1 To MaxLines) As String
For i = 1 To MaxLines
    MOTD(i) = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
Next i
Exit Sub
fallo:
Call LogError("LOADMOTD" & Err.number & " D: " & Err.Description)


End Sub
Public Sub DoBackUp()
'Call LogTarea("Sub DoBackUp")
On Error GoTo fallo
haciendoBK = True
Call SendData2(ToAll, 0, 0, 19)

Call SaveGuildsDB
Call LimpiarMundo
Call WorldSave

Call SendData2(ToAll, 0, 0, 19)

haciendoBK = False

'Log

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time
Close #nfile

Exit Sub
fallo:
Call LogError("DOBACKUP" & Err.number & " D: " & Err.Description)


End Sub

Public Sub grabaPJ()
On Error GoTo fallo
Dim Pj As Integer
Dim Name As String
haciendoBKPJ = True
Call SendData(ToAll, 0, 0, "||%%%% POR FAVOR ESPERE, GRABANDO FICHAS DE PJS...%%%%" & "´" & FontTypeNames.FONTTYPE_info)
Call SendData2(ToAll, 0, 0, 19)
For Pj = 1 To LastUser
Call SaveUser(Pj, CharPath & Left$(UCase$(UserList(Pj).Name), 1) & "\" & UCase$(UserList(Pj).Name) & ".chr")
Next Pj
Call SendData2(ToAll, 0, 0, 19)
Call SendData(ToAll, 0, 0, "||%%%% FICHAS GRABADAS, PUEDEN CONTINUAR.GRACIAS. %%%%" & "´" & FontTypeNames.FONTTYPE_info)

haciendoBKPJ = False

'Log

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackupPJ.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time
Close #nfile

Exit Sub
fallo:
Call LogError("GRABAPJ" & Err.number & " D: " & Err.Description)


End Sub
Public Sub SaveMapData(ByVal n As Integer)

'Call LogTarea("Sub SaveMapData N:" & n)
On Error GoTo fallo
Dim loopc As Integer
Dim TempInt As Integer
Dim Y As Integer
Dim X As Integer
Dim SaveAs As String

SaveAs = App.Path & "\WorldBackUP\Map" & n & ".map"

If FileExist(SaveAs, vbNormal) Then
    Kill SaveAs
End If

If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) Then
    Kill Left$(SaveAs, Len(SaveAs) - 4) & ".inf"
End If

'Open .map file
Open SaveAs For Binary As #1
Seek #1, 1
SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
SaveAs = SaveAs & ".inf"
'Open .inf file
Open SaveAs For Binary As #2
Seek #2, 1
'map Header
        
Put #1, , MapInfo(n).MapVersion
Put #1, , MiCabecera
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt

'inf Header
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt

'Write .map file
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        
        '.map file
        Put #1, , MapData(n, X, Y).Blocked
        
        For loopc = 1 To 4
            Put #1, , MapData(n, X, Y).Graphic(loopc)
        Next loopc
        
        'Lugar vacio para futuras expansiones
        Put #1, , MapData(n, X, Y).trigger
        
        Put #1, , TempInt
        
        '.inf file
        'Tile exit
        Put #2, , MapData(n, X, Y).TileExit.Map
        Put #2, , MapData(n, X, Y).TileExit.X
        Put #2, , MapData(n, X, Y).TileExit.Y
        
        'NPC
        If MapData(n, X, Y).NpcIndex > 0 Then
            Put #2, , Npclist(MapData(n, X, Y).NpcIndex).numero
        Else
            Put #2, , 0
        End If
        'Object
        
        If MapData(n, X, Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(n, X, Y).OBJInfo.ObjIndex).OBJType = OBJTYPE_FOGATA Then
                MapData(n, X, Y).OBJInfo.ObjIndex = 0
                MapData(n, X, Y).OBJInfo.Amount = 0
            End If
'            If ObjData(MapData(n, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_MANCHAS Then
'                MapData(n, X, Y).OBJInfo.ObjIndex = 0
'                MapData(n, X, Y).OBJInfo.Amount = 0
'            End If
        End If
        
        Put #2, , MapData(n, X, Y).OBJInfo.ObjIndex
        Put #2, , MapData(n, X, Y).OBJInfo.Amount
        
        'Empty place holders for future expansion
        Put #2, , TempInt
        Put #2, , TempInt
        
    Next X
Next Y

'Close .map file
Close #1

'Close .inf file
Close #2

'write .dat file
SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
Call WriteVar(SaveAs, "Mapa" & n, "Name", MapInfo(n).Name)
Call WriteVar(SaveAs, "Mapa" & n, "MusicNum", MapInfo(n).Music)
Call WriteVar(SaveAs, "Mapa" & n, "StartPos", MapInfo(n).StartPos.Map & "-" & MapInfo(n).StartPos.X & "-" & MapInfo(n).StartPos.Y)

Call WriteVar(SaveAs, "Mapa" & n, "Terreno", MapInfo(n).Terreno)
Call WriteVar(SaveAs, "Mapa" & n, "Zona", MapInfo(n).Zona)
Call WriteVar(SaveAs, "Mapa" & n, "Restringir", MapInfo(n).Restringir)
Call WriteVar(SaveAs, "Mapa" & n, "BackUp", str(MapInfo(n).BackUp))
Call WriteVar(SaveAs, "Mapa" & n, "Dueño", str(MapInfo(n).Dueño))
Call WriteVar(SaveAs, "Mapa" & n, "Aldea", str(MapInfo(n).Aldea))
'pluto:6.0A
Call WriteVar(SaveAs, "Mapa" & n, "Invisible", str(MapInfo(n).Invisible))
Call WriteVar(SaveAs, "Mapa" & n, "Resucitar", str(MapInfo(n).Resucitar))
Call WriteVar(SaveAs, "Mapa" & n, "Mascotas", str(MapInfo(n).Mascotas))
Call WriteVar(SaveAs, "Mapa" & n, "Insegura", str(MapInfo(n).Insegura))
Call WriteVar(SaveAs, "Mapa" & n, "Lluvia", str(MapInfo(n).Lluvia))
Call WriteVar(SaveAs, "Mapa" & n, "Domar", str(MapInfo(n).Domar))
Call WriteVar(SaveAs, "Mapa" & n, "Monturas", str(MapInfo(n).Monturas))
If MapInfo(n).Pk Then
    Call WriteVar(SaveAs, "Mapa" & n, "pk", "0")
Else
    Call WriteVar(SaveAs, "Mapa" & n, "pk", "1")
End If

Exit Sub
fallo:
Call LogError("SAVEMAPDATA" & Err.number & " D: " & Err.Description)


End Sub

Sub LoadArmasHerreria()
On Error GoTo fallo
Dim n As Integer, LC As Integer

n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To n) As Integer

For LC = 1 To n
    ArmasHerrero(LC) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & LC, "Index"))
'pluto:6.0a
ObjData(ArmasHerrero(LC)).ParaHerre = 1
Next LC
Exit Sub
fallo:
Call LogError("LOADARMASHERRERIA" & Err.number & " D: " & Err.Description)


End Sub

Sub LoadArmadurasHerreria()
On Error GoTo fallo
Dim n As Integer, LC As Integer

n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To n) As Integer

For LC = 1 To n
    ArmadurasHerrero(LC) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & LC, "Index"))
'pluto:6.0a
ObjData(ArmadurasHerrero(LC)).ParaHerre = 1
Next LC
Exit Sub
fallo:
Call LogError("LOADARMADURASHERRERIA" & Err.number & " D: " & Err.Description)

End Sub
Sub LoadPorcentajesMascotas()


PMascotas(1).Tipo = "Unicornio"
PMascotas(2).Tipo = "Caballo Negro"
PMascotas(3).Tipo = "Tigre"
PMascotas(4).Tipo = "Elefante"
PMascotas(5).Tipo = "Dragón"
PMascotas(6).Tipo = "Jabato"
PMascotas(7).Tipo = "Jabalí"
PMascotas(8).Tipo = "Escarabajo"
PMascotas(9).Tipo = "Rinosaurio"
PMascotas(10).Tipo = "Cerbero"
PMascotas(11).Tipo = "Wyvern"
PMascotas(12).Tipo = "Avestruz"

 'unicornio
PMascotas(1).AumentoMagia = 15
PMascotas(1).ReduceMagia = 9
PMascotas(1).AumentoEvasion = 6
PMascotas(1).VidaporLevel = 35
PMascotas(1).GolpeporLevel = 6
PMascotas(1).TopeAtMagico = 15
PMascotas(1).TopeDefMagico = 9
PMascotas(1).TopeEvasion = 6
'negro
PMascotas(2).AumentoMagia = 2
PMascotas(2).ReduceMagia = 4
PMascotas(2).AumentoEvasion = 1
PMascotas(2).VidaporLevel = 30
PMascotas(2).GolpeporLevel = 8
PMascotas(2).TopeAtMagico = 9
PMascotas(2).TopeDefMagico = 15
PMascotas(2).TopeEvasion = 6
'tigre
PMascotas(3).ReduceCuerpo = 2
PMascotas(3).AumentoEvasion = 4
PMascotas(3).AumentoFlecha = 1
PMascotas(3).VidaporLevel = 35
PMascotas(3).GolpeporLevel = 10
PMascotas(3).TopeAtFlechas = 9
PMascotas(3).TopeDefMagico = 9
PMascotas(3).TopeEvasion = 12
'elefante
PMascotas(4).AumentoCuerpo = 4
PMascotas(4).ReduceCuerpo = 1
PMascotas(4).ReduceFlecha = 1
PMascotas(4).VidaporLevel = 50
PMascotas(4).GolpeporLevel = 12
PMascotas(4).TopeAtCuerpo = 15
PMascotas(4).TopeDefCuerpo = 9
PMascotas(4).TopeEvasion = 6
'dragon
PMascotas(5).AumentoCuerpo = 4
PMascotas(5).ReduceCuerpo = 4
PMascotas(5).AumentoMagia = 4
PMascotas(5).ReduceMagia = 4
PMascotas(5).AumentoFlecha = 4
PMascotas(5).ReduceFlecha = 4
PMascotas(5).AumentoEvasion = 4
PMascotas(5).VidaporLevel = 80
PMascotas(5).GolpeporLevel = 28
PMascotas(5).TopeAtMagico = 9
PMascotas(5).TopeDefMagico = 9
PMascotas(5).TopeEvasion = 9
PMascotas(5).TopeAtCuerpo = 9
PMascotas(5).TopeDefCuerpo = 9
PMascotas(5).TopeAtFlechas = 9
PMascotas(5).TopeDefFlechas = 9
'jabalí pequeño
PMascotas(6).AumentoCuerpo = 1
PMascotas(6).ReduceCuerpo = 1
PMascotas(6).ReduceFlecha = 0
PMascotas(6).VidaporLevel = 7
PMascotas(6).GolpeporLevel = 6
PMascotas(6).TopeAtMagico = 16
PMascotas(6).TopeDefMagico = 16
PMascotas(6).TopeEvasion = 16
PMascotas(6).TopeAtCuerpo = 16
PMascotas(6).TopeDefCuerpo = 16
PMascotas(6).TopeAtFlechas = 16
PMascotas(6).TopeDefFlechas = 16
'jabalí gigante
PMascotas(7).AumentoCuerpo = 2
PMascotas(7).ReduceCuerpo = 2
PMascotas(7).ReduceFlecha = 3
PMascotas(7).VidaporLevel = 35
PMascotas(7).GolpeporLevel = 8
PMascotas(7).TopeDefCuerpo = 12
PMascotas(7).TopeAtCuerpo = 9
PMascotas(7).TopeDefFlechas = 9
'escarabajo
PMascotas(8).AumentoMagia = 3
PMascotas(8).ReduceMagia = 3
PMascotas(8).AumentoEvasion = 1
PMascotas(8).VidaporLevel = 40
PMascotas(8).GolpeporLevel = 7
PMascotas(8).TopeDefCuerpo = 12
PMascotas(8).TopeDefMagico = 12
PMascotas(8).TopeAtMagico = 6
'rinosaurio
PMascotas(9).AumentoCuerpo = 1
PMascotas(9).ReduceCuerpo = 4
PMascotas(9).ReduceFlecha = 1
PMascotas(9).VidaporLevel = 55
PMascotas(9).GolpeporLevel = 12
PMascotas(9).TopeEvasion = 9
PMascotas(9).TopeDefMagico = 15
PMascotas(9).TopeAtCuerpo = 6
'cerbero
PMascotas(10).ReduceCuerpo = 4
PMascotas(10).AumentoEvasion = 2
PMascotas(10).AumentoFlecha = 1
PMascotas(10).VidaporLevel = 45
PMascotas(10).GolpeporLevel = 10
PMascotas(10).TopeAtFlechas = 6
PMascotas(10).TopeDefMagico = 12
PMascotas(10).TopeDefCuerpo = 12
'wyvern
PMascotas(11).AumentoMagia = 2
PMascotas(11).ReduceMagia = 2
PMascotas(11).AumentoEvasion = 3
PMascotas(11).VidaporLevel = 40
PMascotas(11).GolpeporLevel = 10
PMascotas(11).TopeDefFlechas = 9
PMascotas(11).TopeAtMagico = 12
PMascotas(11).TopeDefMagico = 9
'avestruz
PMascotas(12).ReduceCuerpo = 1
PMascotas(12).AumentoEvasion = 2
PMascotas(12).AumentoFlecha = 4
PMascotas(12).VidaporLevel = 35
PMascotas(12).GolpeporLevel = 8
PMascotas(12).TopeAtFlechas = 15
PMascotas(12).TopeDefFlechas = 9
PMascotas(12).TopeEvasion = 6
'tope niveles
PMascotas(1).TopeLevel = 30
PMascotas(2).TopeLevel = 30
PMascotas(3).TopeLevel = 30
PMascotas(4).TopeLevel = 30
PMascotas(5).TopeLevel = 16
PMascotas(6).TopeLevel = 16
PMascotas(7).TopeLevel = 30
PMascotas(8).TopeLevel = 30
PMascotas(9).TopeLevel = 30
PMascotas(10).TopeLevel = 30
PMascotas(11).TopeLevel = 30
PMascotas(12).TopeLevel = 30

'pluto:6.0A cargamos exp mascotas
Dim n As Byte
Dim nn As Byte
Dim aa As Integer
Dim bb As Long
Dim cc As Long
For n = 1 To 30
aa = aa + 400
bb = bb + 1800
cc = cc + 20
For nn = 1 To 12

If nn = 5 Then
PMascotas(nn).exp(n) = PMascotas(nn).exp(n) + bb
ElseIf nn = 6 Then
PMascotas(nn).exp(n) = PMascotas(nn).exp(n) + cc
Else
PMascotas(nn).exp(n) = PMascotas(nn).exp(n) + aa
End If

Next nn
Next n

End Sub
Sub LoadObjCarpintero()
On Error GoTo fallo
Dim n As Integer, LC As Integer

n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To n) As Integer

For LC = 1 To n
    ObjCarpintero(LC) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & LC, "Index"))
'pluto:6.0a
ObjData(ObjCarpintero(LC)).ParaCarpin = 1
Next LC
Exit Sub
fallo:
Call LogError("LOADOBJCARPINTERO" & Err.number & " D: " & Err.Description)

End Sub

'[MerLiNz:6]
Sub LoadObjMagicosermitano()
On Error GoTo fallo
Dim n As Integer, LC As Integer

n = val(GetVar(DatPath & "Objermitano.dat", "INIT", "NumObjs"))

ReDim Preserve Objermitano(1 To n) As Integer

For LC = 1 To n
    Objermitano(LC) = val(GetVar(DatPath & "Objermitano.dat", "Obj" & LC, "Index"))
'pluto:6.0a
ObjData(Objermitano(LC)).ParaErmi = 1
Next LC

Exit Sub
fallo:
Call LogError("LOADOBJMAGICOERMITAÑO" & Err.number & " D: " & Err.Description)

'[\END]
End Sub
'Pluto:hoy
Sub Loadtrivial()
On Error GoTo perro
Dim n As Integer
Dim numtrivial As Integer
Dim Leer As New clsLeerInis
Dim obj As ObjData


Leer.Abrir DatPath & "Trivial.txt"



'numtrivial = val(GetVar(DatPath & "Trivial.txt", "INIT", "NumTrivial"))
numtrivial = val(Leer.DarValor("INIT", "NumTrivial"))

n = RandomNumber(1, numtrivial)
'PreTrivial = GetVar(DatPath & "TRIVIAL.TXT", "T" & n, "tx")
PreTrivial = Leer.DarValor("T" & n, "tx")

'ResTrivial = GetVar(DatPath & "TRIVIAL.TXT", "T" & n, "RES")
ResTrivial = Leer.DarValor("T" & n, "RES")

Exit Sub

perro:
LogError ("Trivial: Error en la pregunta numero: " & n & " : " & Err.Description)
End Sub

'Pluto:2.4
Sub Loadrecord()
On Error GoTo perro
NivCrimi = val(GetVar(IniPath & "RECORD.TXT", "INIT", "NivCrimi"))
NivCiu = val(GetVar(IniPath & "RECORD.TXT", "INIT", "NivCiu"))
MaxTorneo = val(GetVar(IniPath & "RECORD.TXT", "INIT", "MaxTorneo"))
Moro = val(GetVar(IniPath & "RECORD.TXT", "INIT", "Moro"))
NNivCrimi = GetVar(IniPath & "RECORD.TXT", "INIT", "NNivCrimi")
NNivCiu = GetVar(IniPath & "RECORD.TXT", "INIT", "NNivCiu")
NMaxTorneo = GetVar(IniPath & "RECORD.TXT", "INIT", "NMaxTorneo")
NMoro = GetVar(IniPath & "RECORD.TXT", "INIT", "NMoro")
'pluto:6.9
'Clan1Torneo = GetVar(IniPath & "RECORD.TXT", "INIT", "Clan1Torneo")
'Clan2Torneo = GetVar(IniPath & "RECORD.TXT", "INIT", "Clan2Torneo")
'PClan1Torneo = val(GetVar(IniPath & "RECORD.TXT", "INIT", "PClan1Torneo"))
'PClan2Torneo = val(GetVar(IniPath & "RECORD.TXT", "INIT", "PClan2Torneo"))
Exit Sub
perro:
LogError ("Records: Error en cargando Records: " & Err.Description)
End Sub
'Pluto:hoy
Sub LoadEgipto()
On Error GoTo perro
Dim n As Integer
Dim numegipto As Integer
numegipto = val(GetVar(DatPath & "egipto.txt", "INIT", "NumEgipto"))
n = RandomNumber(1, numegipto)
PreEgipto = GetVar(DatPath & "EGIPTO.TXT", "T" & n, "tx")
ResEgipto = GetVar(DatPath & "EGIPTO.TXT", "T" & n, "RES")
Exit Sub
perro:
LogError ("Egipto: Error en la pregunta numero: " & n & " : " & Err.Description)
End Sub
Sub LoadOBJData()
On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
'pluto fusion
Dim Leer As New clsLeerInis
Leer.Abrir DatPath & "Obj.dat"

'obtiene el numero de obj
NumObjDatas = val(Leer.DarValor("INIT", "NumObjs"))

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
Dim Calcu As Double
  
'Llena la lista
For Object = 1 To NumObjDatas
    Calcu = Object
    Calcu = Calcu * 100
    Calcu = Calcu / NumObjDatas
    frmCargando.Label1(2).Caption = "Objeto: (" & Object & "/" & NumObjDatas & ") " & Round(Calcu, 1) & "%"

    ObjData(Object).Name = Leer.DarValor("OBJ" & Object, "Name")
    'ObjData(Object).Name = Leer.DarValor("OBJ" & Object, "Name")
'pluto 2.17
ObjData(Object).Magia = val(Leer.DarValor("OBJ" & Object, "Magia"))

    'pluto:2.8.0
    ObjData(Object).Vendible = val(Leer.DarValor("OBJ" & Object, "Vendible"))

    
    ObjData(Object).GrhIndex = val(Leer.DarValor("OBJ" & Object, "GrhIndex"))
    
    ObjData(Object).OBJType = val(Leer.DarValor("OBJ" & Object, "ObjType"))
    ObjData(Object).SubTipo = val(Leer.DarValor("OBJ" & Object, "Subtipo"))
    'pluto:6.0A
    ObjData(Object).ArmaNpc = val(Leer.DarValor("OBJ" & Object, "ArmaNpc"))
    
    ObjData(Object).Newbie = val(Leer.DarValor("OBJ" & Object, "Newbie"))
    'pluto:2.3
    ObjData(Object).Peso = 0 ' val(Leer.DarValor("OBJ" & Object, "Peso"))

        If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
        ObjData(Object).ShieldAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
        ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))  ' * 2
        '[MerLiNz:6]
        ObjData(Object).Gemas = val(Leer.DarValor("OBJ" & Object, "Gemas"))
        ObjData(Object).Diamantes = val(Leer.DarValor("OBJ" & Object, "Diamantes"))
        '[\END]
    End If
    'pluto:6.2----------
    If ObjData(Object).OBJType = OBJTYPE_Anillo Then
    ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
    ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
    ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
    ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))  ' * 2
    ObjData(Object).Gemas = val(Leer.DarValor("OBJ" & Object, "Gemas"))
    ObjData(Object).Diamantes = val(Leer.DarValor("OBJ" & Object, "Diamantes"))
    End If
    '--------------------
    
    If ObjData(Object).SubTipo = OBJTYPE_CASCO Then
        
        ObjData(Object).CascoAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
        ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
        '[MerLiNz:6]
        ObjData(Object).Gemas = val(Leer.DarValor("OBJ" & Object, "Gemas"))
        ObjData(Object).Diamantes = val(Leer.DarValor("OBJ" & Object, "Diamantes"))
        '[\END]
        ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))  '* 2
    
    End If
     '[GAU]
    If ObjData(Object).SubTipo = OBJTYPE_BOTA Then
        ObjData(Object).Botas = val(Leer.DarValor("OBJ" & Object, "Anim"))
        ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))  ' * 2
    End If
    '[GAU]
    ObjData(Object).Ropaje = val(Leer.DarValor("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(Leer.DarValor("OBJ" & Object, "HechizoIndex"))
    
    If ObjData(Object).OBJType = OBJTYPE_WEAPON Then
             ObjData(Object).WeaponAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(Leer.DarValor("OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = val(Leer.DarValor("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))  ' * 2
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
            ObjData(Object).proyectil = val(Leer.DarValor("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(Leer.DarValor("OBJ" & Object, "Municiones"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = val(Leer.DarValor("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(Leer.DarValor("OBJ" & Object, "Diamantes"))
            '[\END]
    ObjData(Object).SkArma = val(Leer.DarValor("OBJ" & Object, "SKARMA"))
    ObjData(Object).SkArco = val(Leer.DarValor("OBJ" & Object, "SKARCO"))

    End If
    
    If ObjData(Object).OBJType = OBJTYPE_ARMOUR Then
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))  ' * 2
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = val(Leer.DarValor("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(Leer.DarValor("OBJ" & Object, "Diamantes"))
            'pluto:2.10
            ObjData(Object).ObjetoClan = Leer.DarValor("OBJ" & Object, "ObjetoClan")

            '[\END]
    End If
    
    If ObjData(Object).OBJType = OBJTYPE_HERRAMIENTAS Then
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))  '* 2
            '[MerLiNz:6]
            ObjData(Object).Gemas = val(Leer.DarValor("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(Leer.DarValor("OBJ" & Object, "Diamantes"))
            '[\END]
    End If
    
    If ObjData(Object).OBJType = OBJTYPE_INSTRUMENTOS Then
        ObjData(Object).Snd1 = val(Leer.DarValor("OBJ" & Object, "SND1"))
        ObjData(Object).Snd2 = val(Leer.DarValor("OBJ" & Object, "SND2"))
        ObjData(Object).Snd3 = val(Leer.DarValor("OBJ" & Object, "SND3"))
        ObjData(Object).MinInt = val(Leer.DarValor("OBJ" & Object, "MinInt"))
    End If
    
    ObjData(Object).LingoteIndex = val(Leer.DarValor("OBJ" & Object, "LingoteIndex"))
    
    If ObjData(Object).OBJType = 31 Or ObjData(Object).OBJType = 23 Then
        ObjData(Object).MinSkill = val(Leer.DarValor("OBJ" & Object, "MinSkill"))
    End If
    
    ObjData(Object).MineralIndex = val(Leer.DarValor("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = val(Leer.DarValor("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(Leer.DarValor("OBJ" & Object, "MinHP"))
  
    
    ObjData(Object).Mujer = val(Leer.DarValor("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(Leer.DarValor("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = val(Leer.DarValor("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(Leer.DarValor("OBJ" & Object, "MinAgu"))
    
    'pluto:7.0
    ObjData(Object).MinDef = val(Leer.DarValor("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(Leer.DarValor("OBJ" & Object, "MAXDEF"))
    ObjData(Object).Defmagica = val(Leer.DarValor("OBJ" & Object, "DEFMAGICA"))
    'nati:agrego DefCuerpo
    ObjData(Object).Defcuerpo = val(Leer.DarValor("OBJ" & Object, "DEFCUERPO"))
    ObjData(Object).Drop = val(Leer.DarValor("OBJ" & Object, "DROP"))

    'ObjData(Object).Defproyectil = val(Leer.DarValor("OBJ" & Object, "DEFPROYECTIL"))

    ObjData(Object).Respawn = val(Leer.DarValor("OBJ" & Object, "ReSpawn"))
    
    ObjData(Object).RazaEnana = val(Leer.DarValor("OBJ" & Object, "RazaEnana"))
    ObjData(Object).razaelfa = val(Leer.DarValor("OBJ" & Object, "RazaElfa"))
    ObjData(Object).razavampiro = val(Leer.DarValor("OBJ" & Object, "Razavampiro"))
    ObjData(Object).razaorca = val(Leer.DarValor("OBJ" & Object, "Razaorca"))
    ObjData(Object).razahumana = val(Leer.DarValor("OBJ" & Object, "Razahumana"))

    ObjData(Object).Valor = val(Leer.DarValor("OBJ" & Object, "Valor"))
    ObjData(Object).nocaer = val(Leer.DarValor("OBJ" & Object, "nocaer"))
    ObjData(Object).objetoespecial = val(Leer.DarValor("OBJ" & Object, "objetoespecial"))

    ObjData(Object).Crucial = val(Leer.DarValor("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = val(Leer.DarValor("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = val(Leer.DarValor("OBJ" & Object, "Llave"))
            ObjData(Object).Clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
    End If
    
    
    If ObjData(Object).OBJType = OBJTYPE_PUERTAS Or ObjData(Object).OBJType = OBJTYPE_BOTELLAVACIA Or ObjData(Object).OBJType = OBJTYPE_BOTELLALLENA Then
        ObjData(Object).IndexAbierta = val(Leer.DarValor("OBJ" & Object, "IndexAbierta"))
        ObjData(Object).IndexCerrada = val(Leer.DarValor("OBJ" & Object, "IndexCerrada"))
        ObjData(Object).IndexCerradaLlave = val(Leer.DarValor("OBJ" & Object, "IndexCerradaLlave"))
    End If
    
    
    'Puertas y llaves
    ObjData(Object).Clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = Leer.DarValor("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(Leer.DarValor("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(Leer.DarValor("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.DarValor("OBJ" & Object, "ID")
    
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        ObjData(Object).ClaseProhibida(i) = Leer.DarValor("OBJ" & Object, "CP" & i)
    Next
            
    ObjData(Object).Resistencia = val(Leer.DarValor("OBJ" & Object, "Resistencia"))
    
    'Pociones
    If ObjData(Object).OBJType = 11 Then
        ObjData(Object).TipoPocion = val(Leer.DarValor("OBJ" & Object, "TipoPocion"))
        ObjData(Object).MaxModificador = val(Leer.DarValor("OBJ" & Object, "MaxModificador"))
        ObjData(Object).MinModificador = val(Leer.DarValor("OBJ" & Object, "MinModificador"))
        ObjData(Object).DuracionEfecto = val(Leer.DarValor("OBJ" & Object, "DuracionEfecto"))
    End If

    ObjData(Object).SkCarpinteria = val(Leer.DarValor("OBJ" & Object, "SkCarpinteria"))  '* 2
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = val(Leer.DarValor("OBJ" & Object, "Madera"))
    
    If ObjData(Object).OBJType = OBJTYPE_BARCOS Then
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
    End If
    
    If ObjData(Object).OBJType = OBJTYPE_FLECHAS Then
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
    End If
    
    'Bebidas
    ObjData(Object).MinSta = val(Leer.DarValor("OBJ" & Object, "MinST"))
    ObjData(Object).razavampiro = val(Leer.DarValor("OBJ" & Object, "razavampiro"))
    'pluto:6.0A----
    ObjData(Object).Cregalos = val(Leer.DarValor("OBJ" & Object, "Cregalos"))
    ObjData(Object).Pregalo = val(Leer.DarValor("OBJ" & Object, "Pregalo"))
    '--------------
    frmCargando.cargar.value = frmCargando.cargar.value + 1
'pluto:6.0A
    If ObjData(Object).Pregalo > 0 Then
Select Case ObjData(Object).Pregalo
Case 1
Reo1 = Reo1 + 1
ObjRegalo1(Reo1) = Object
Case 2
Reo2 = Reo2 + 1
ObjRegalo2(Reo2) = Object
Case 3
Reo3 = Reo3 + 1
ObjRegalo3(Reo3) = Object
End Select
    End If
    
Next Object
'quitar esto
Exit Sub
'------------------------------------------------------------------------------------
'Esto genera el obj.log para meterlo al cliente, el server no usa nada de lo de abajo.
'------------------------------------------------------------------------------------
Dim file As String
Dim n As Byte
file = DatPath & "Obj.dat"
Dim nfile As Integer
Dim vec As Byte
Dim vec2 As Integer
vec = 1
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\Objeto.log" For Append Shared As #nfile
Print #nfile, "Sub CargamosObjetos" & vec & "()"
For Object = 1 To NumObjDatas
vec2 = vec2 + 1
Debug.Print Object
If vec2 > 100 Then
vec = vec + 1
vec2 = 0
Print #nfile, "end sub"
Print #nfile, "sub CargamosObjetos" & vec & "()"
End If

Print #nfile, "ObjData(" & Object & ").name=" & Chr(34) & ObjData(Object).Name & Chr(34)
If ObjData(Object).Agarrable > 0 Then Print #nfile, "ObjData(" & Object & ").agarrable =" & ObjData(Object).Agarrable
If ObjData(Object).Apuñala > 0 Then Print #nfile, "ObjData(" & Object & ").apuñala=" & ObjData(Object).Apuñala
If ObjData(Object).ArmaNpc > 0 Then Print #nfile, "ObjData(" & Object & ").armanpc=" & ObjData(Object).ArmaNpc
If ObjData(Object).Botas > 0 Then Print #nfile, "ObjData(" & Object & ").botas=" & ObjData(Object).Botas
If ObjData(Object).Caos > 0 Then Print #nfile, "ObjData(" & Object & ").caos=" & ObjData(Object).Caos
If ObjData(Object).CascoAnim > 0 Then Print #nfile, "ObjData(" & Object & ").cascoanim=" & ObjData(Object).CascoAnim
If ObjData(Object).Cerrada > 0 Then Print #nfile, "ObjData(" & Object & ").cerrada=" & ObjData(Object).Cerrada
For n = 1 To 21
If ObjData(Object).ClaseProhibida(n) <> "" Then Print #nfile, "ObjData(" & Object & ").claseprohibida(" & n & ")=" & Chr(34) & ObjData(Object).ClaseProhibida(n) & Chr(34)
Next
If ObjData(Object).Clave > 0 Then Print #nfile, "ObjData(" & Object & ").clave=" & ObjData(Object).Clave
If ObjData(Object).Crucial > 0 Then Print #nfile, "ObjData(" & Object & ").crucial=" & ObjData(Object).Crucial
If ObjData(Object).Def > 0 Then Print #nfile, "ObjData(" & Object & ").def=" & ObjData(Object).Def
If ObjData(Object).Diamantes > 0 Then Print #nfile, "ObjData(" & Object & ").diamantes=" & ObjData(Object).Diamantes
If ObjData(Object).DuracionEfecto > 0 Then Print #nfile, "ObjData(" & Object & ").duracionefecto=" & ObjData(Object).DuracionEfecto
If ObjData(Object).Envenena > 0 Then Print #nfile, "ObjData(" & Object & ").envenena=" & ObjData(Object).Envenena
If ObjData(Object).ForoID <> "" Then Print #nfile, "ObjData(" & Object & ").foroid=" & Chr(34) & ObjData(Object).ForoID & Chr(34)
If ObjData(Object).Gemas > 0 Then Print #nfile, "ObjData(" & Object & ").gemas=" & ObjData(Object).Gemas
If ObjData(Object).GrhIndex > 0 Then Print #nfile, "ObjData(" & Object & ").grhindex=" & ObjData(Object).GrhIndex
If ObjData(Object).GrhSecundario > 0 Then Print #nfile, "ObjData(" & Object & ").grhsecundario=" & ObjData(Object).GrhSecundario
If ObjData(Object).HechizoIndex > 0 Then Print #nfile, "ObjData(" & Object & ").hechizoindex=" & ObjData(Object).HechizoIndex
If ObjData(Object).Hombre > 0 Then Print #nfile, "ObjData(" & Object & ").hombre=" & ObjData(Object).Hombre
If ObjData(Object).IndexAbierta > 0 Then Print #nfile, "ObjData(" & Object & ").indexabierta=" & ObjData(Object).IndexAbierta
If ObjData(Object).IndexCerrada > 0 Then Print #nfile, "ObjData(" & Object & ").indexcerrada=" & ObjData(Object).IndexCerrada
If ObjData(Object).IndexCerradaLlave > 0 Then Print #nfile, "ObjData(" & Object & ").indexcerradallave=" & ObjData(Object).IndexCerradaLlave
If ObjData(Object).LingH > 0 Then Print #nfile, "ObjData(" & Object & ").lingh=" & ObjData(Object).LingH
If ObjData(Object).LingO > 0 Then Print #nfile, "ObjData(" & Object & ").lingo=" & ObjData(Object).LingO
If ObjData(Object).LingoteIndex > 0 Then Print #nfile, "ObjData(" & Object & ").lingoteindex=" & ObjData(Object).LingoteIndex
If ObjData(Object).LingP > 0 Then Print #nfile, "ObjData(" & Object & ").lingp=" & ObjData(Object).LingP
If ObjData(Object).Llave > 0 Then Print #nfile, "ObjData(" & Object & ").llave=" & ObjData(Object).Llave
If ObjData(Object).Madera > 0 Then Print #nfile, "ObjData(" & Object & ").madera=" & ObjData(Object).Madera
If ObjData(Object).Magia > 0 Then Print #nfile, "ObjData(" & Object & ").magia=" & ObjData(Object).Magia
If ObjData(Object).MaxDef > 0 Then Print #nfile, "ObjData(" & Object & ").maxdef=" & ObjData(Object).MaxDef
If ObjData(Object).MaxHIT > 0 Then Print #nfile, "ObjData(" & Object & ").maxhit=" & ObjData(Object).MaxHIT
If ObjData(Object).MaxHP > 0 Then Print #nfile, "ObjData(" & Object & ").maxhp=" & ObjData(Object).MaxHP
If ObjData(Object).MaxItems > 0 Then Print #nfile, "ObjData(" & Object & ").maxitems=" & ObjData(Object).MaxItems
If ObjData(Object).MaxModificador > 0 Then Print #nfile, "ObjData(" & Object & ").maxmodificador=" & ObjData(Object).MaxModificador
If ObjData(Object).MinDef > 0 Then Print #nfile, "ObjData(" & Object & ").mindef=" & ObjData(Object).MinDef
'pluto:7.0
If ObjData(Object).Defmagica > 0 Then Print #nfile, "ObjData(" & Object & ").defmagica =" & ObjData(Object).Defmagica
'nati: Agrego defCuerpo
If ObjData(Object).Defcuerpo > 0 Then Print #nfile, "ObjData(" & Object & ").defcuerpo =" & ObjData(Object).Defcuerpo
'If ObjData(Object).Defproyectil > 0 Then Print #nfile, "ObjData(" & Object & ").defproyectil =" & ObjData(Object).Defproyectil

If ObjData(Object).MineralIndex > 0 Then Print #nfile, "ObjData(" & Object & ").mineralindex=" & ObjData(Object).MineralIndex
If ObjData(Object).MinHam > 0 Then Print #nfile, "ObjData(" & Object & ").minham=" & ObjData(Object).MinHam
If ObjData(Object).MinHIT > 0 Then Print #nfile, "ObjData(" & Object & ").minhit=" & ObjData(Object).MinHIT
If ObjData(Object).MinHP > 0 Then Print #nfile, "ObjData(" & Object & ").minhp=" & ObjData(Object).MinHP
If ObjData(Object).MinInt > 0 Then Print #nfile, "ObjData(" & Object & ").minint=" & ObjData(Object).MinInt
If ObjData(Object).MinModificador > 0 Then Print #nfile, "ObjData(" & Object & ").minmodificador=" & ObjData(Object).MinModificador
If ObjData(Object).MinSed > 0 Then Print #nfile, "ObjData(" & Object & ").minsed=" & ObjData(Object).MinSed
If ObjData(Object).MinSkill > 0 Then Print #nfile, "ObjData(" & Object & ").minskill=" & ObjData(Object).MinSkill
If ObjData(Object).MinSta > 0 Then Print #nfile, "ObjData(" & Object & ").minsta=" & ObjData(Object).MinSta
If ObjData(Object).Mujer > 0 Then Print #nfile, "ObjData(" & Object & ").mujer=" & ObjData(Object).Mujer
If ObjData(Object).Municion > 0 Then Print #nfile, "ObjData(" & Object & ").municion=" & ObjData(Object).Municion
If ObjData(Object).Newbie > 0 Then Print #nfile, "ObjData(" & Object & ").Newbie=" & ObjData(Object).Newbie
If ObjData(Object).nocaer > 0 Then Print #nfile, "ObjData(" & Object & ").nocaer=" & ObjData(Object).nocaer
If ObjData(Object).ObjetoClan <> "" Then Print #nfile, "ObjData(" & Object & ").objetoclan=" & Chr(34) & ObjData(Object).ObjetoClan & Chr(34)
If ObjData(Object).objetoespecial > 0 Then Print #nfile, "ObjData(" & Object & ").objetoespecial=" & ObjData(Object).objetoespecial
If ObjData(Object).OBJType > 0 Then Print #nfile, "ObjData(" & Object & ").objtype=" & ObjData(Object).OBJType
If ObjData(Object).Peso > 0 Then Print #nfile, "ObjData(" & Object & ").peso=" & ObjData(Object).Peso
If ObjData(Object).proyectil > 0 Then Print #nfile, "ObjData(" & Object & ").proyectil=" & ObjData(Object).proyectil
If ObjData(Object).razaelfa > 0 Then Print #nfile, "ObjData(" & Object & ").razaelfa=" & ObjData(Object).razaelfa
If ObjData(Object).RazaEnana > 0 Then Print #nfile, "ObjData(" & Object & ").razaenana=" & ObjData(Object).RazaEnana
If ObjData(Object).razahumana > 0 Then Print #nfile, "ObjData(" & Object & ").razahumana=" & ObjData(Object).razahumana
If ObjData(Object).razaorca > 0 Then Print #nfile, "ObjData(" & Object & ").razaorca=" & ObjData(Object).razaorca
If ObjData(Object).razavampiro > 0 Then Print #nfile, "ObjData(" & Object & ").razavampiro=" & ObjData(Object).razavampiro
If ObjData(Object).Real > 0 Then Print #nfile, "ObjData(" & Object & ").real=" & ObjData(Object).Real
If ObjData(Object).Resistencia > 0 Then Print #nfile, "ObjData(" & Object & ").resistencia=" & ObjData(Object).Resistencia
If ObjData(Object).Respawn > 0 Then Print #nfile, "ObjData(" & Object & ").respawn=" & ObjData(Object).Respawn
If ObjData(Object).Ropaje > 0 Then Print #nfile, "ObjData(" & Object & ").ropaje=" & ObjData(Object).Ropaje
If ObjData(Object).ShieldAnim > 0 Then Print #nfile, "ObjData(" & Object & ").shieldanim=" & ObjData(Object).ShieldAnim
If ObjData(Object).SkArco > 0 Then Print #nfile, "ObjData(" & Object & ").skarco=" & ObjData(Object).SkArco
If ObjData(Object).SkArma > 0 Then Print #nfile, "ObjData(" & Object & ").skarma=" & ObjData(Object).SkArma
If ObjData(Object).SkCarpinteria > 0 Then Print #nfile, "ObjData(" & Object & ").skcarpinteria=" & ObjData(Object).SkCarpinteria
If ObjData(Object).SkHerreria > 0 Then Print #nfile, "ObjData(" & Object & ").skherreria=" & ObjData(Object).SkHerreria
If ObjData(Object).Snd1 > 0 Then Print #nfile, "ObjData(" & Object & ").snd1=" & ObjData(Object).Snd1
If ObjData(Object).Snd2 > 0 Then Print #nfile, "ObjData(" & Object & ").snd2=" & ObjData(Object).Snd2
If ObjData(Object).Snd3 > 0 Then Print #nfile, "ObjData(" & Object & ").snd3=" & ObjData(Object).Snd3
If ObjData(Object).SubTipo > 0 Then Print #nfile, "ObjData(" & Object & ").subtipo=" & ObjData(Object).SubTipo
If ObjData(Object).texto <> "" Then Print #nfile, "ObjData(" & Object & ").texto=" & Chr(34) & ObjData(Object).texto & Chr(34)
If ObjData(Object).TipoPocion > 0 Then Print #nfile, "ObjData(" & Object & ").tipopocion=" & ObjData(Object).TipoPocion
If ObjData(Object).Valor > 0 Then Print #nfile, "ObjData(" & Object & ").valor=" & ObjData(Object).Valor
If ObjData(Object).Vendible > 0 Then Print #nfile, "ObjData(" & Object & ").vendible=" & ObjData(Object).Vendible
If ObjData(Object).WeaponAnim > 0 Then Print #nfile, "ObjData(" & Object & ").weaponanim=" & ObjData(Object).WeaponAnim
If ObjData(Object).Pregalo > 0 Then Print #nfile, "ObjData(" & Object & ").pregalo=" & ObjData(Object).Pregalo
If ObjData(Object).Cregalos > 0 Then Print #nfile, "ObjData(" & Object & ").cregalos=" & ObjData(Object).Cregalos
'pluto:7.0
If ObjData(Object).Drop > 0 Then Print #nfile, "ObjData(" & Object & ").drop=" & ObjData(Object).Drop

DoEvents
Next
Close #nfile



Exit Sub

errhandler:
    MsgBox "error cargando objetos"


End Sub
'pluto:2.3
Sub LoadUserMontura(UserIndex As Integer, userfile As String)
'on error GoTo fallo
'Dim LoopC As Integer
'Dim Leer As New clsLeerInis
'Leer.Abrir userfile
'For LoopC = 1 To MAXMONTURA
'UserList(UserIndex).Montura.Nivel(LoopC) = val(leer.darvalor("MONTURA", "NIVEL" & LoopC))
'UserList(UserIndex).Montura.exp(LoopC) = val(leer.darvalor("MONTURA", "EXP" & LoopC))
'UserList(UserIndex).Montura.Elu(LoopC) = val(leer.darvalor("MONTURA", "ELU" & LoopC))
'UserList(UserIndex).Montura.Vida(LoopC) = val(leer.darvalor("MONTURA", "VIDA" & LoopC))
'UserList(UserIndex).Montura.Golpe(LoopC) = val(leer.darvalor("MONTURA", "GOLPE" & LoopC))
'UserList(UserIndex).Montura.Nombre(LoopC) = leer.darvalor("MONTURA", "NOMBRE" & LoopC)

'Next

'Exit Sub
'fallo:
'Call LogError("LOADUSERMONTURA" & Err.Number & " D: " & Err.Description)


End Sub

Sub LoadUserStats(UserIndex As Integer, userfile As String)
'on error GoTo fallo
'Dim LoopC As Integer

'For LoopC = 1 To NUMATRIBUTOS
 ' UserList(UserIndex).Stats.UserAtributos(LoopC) = leer.darvalor( "ATRIBUTOS", "AT" & LoopC)
  'UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
'Next

'For LoopC = 1 To NUMSKILLS
 ' UserList(UserIndex).Stats.UserSkills(LoopC) = val(leer.darvalor( "SKILLS", "SK" & LoopC))
'Next

'For LoopC = 1 To MAXUSERHECHIZOS
 ' UserList(UserIndex).Stats.UserHechizos(LoopC) = val(leer.darvalor( "Hechizos", "H" & LoopC))
'Next
'pluto:2-3-04
'UserList(UserIndex).Stats.Puntos = val(leer.darvalor( "STATS", "PUNTOS"))

'UserList(UserIndex).Stats.GLD = val(leer.darvalor( "STATS", "GLD"))
'UserList(UserIndex).Remort = val(leer.darvalor( "STATS", "REMORT"))
'UserList(UserIndex).Stats.Banco = val(leer.darvalor( "STATS", "BANCO"))

'UserList(UserIndex).Stats.MET = val(leer.darvalor( "STATS", "MET"))
'UserList(UserIndex).Stats.MaxHP = val(leer.darvalor( "STATS", "MaxHP"))
'UserList(UserIndex).Stats.MinHP = val(leer.darvalor( "STATS", "MinHP"))

'UserList(UserIndex).Stats.FIT = val(leer.darvalor( "STATS", "FIT"))
'UserList(UserIndex).Stats.MinSta = val(leer.darvalor( "STATS", "MinSTA"))
'UserList(UserIndex).Stats.MaxSta = val(leer.darvalor( "STATS", "MaxSTA"))

'UserList(UserIndex).Stats.MaxMAN = val(leer.darvalor( "STATS", "MaxMAN"))
'UserList(UserIndex).Stats.MinMAN = val(leer.darvalor( "STATS", "MinMAN"))

'UserList(UserIndex).Stats.MaxHIT = val(leer.darvalor( "STATS", "MaxHIT"))
'UserList(UserIndex).Stats.MinHIT = val(leer.darvalor( "STATS", "MinHIT"))

'UserList(UserIndex).Stats.MaxAGU = val(leer.darvalor( "STATS", "MaxAGU"))
'UserList(UserIndex).Stats.MinAGU = val(leer.darvalor( "STATS", "MinAGU"))

'UserList(UserIndex).Stats.MaxHam = val(leer.darvalor( "STATS", "MaxHAM"))
'UserList(UserIndex).Stats.MinHam = val(leer.darvalor( "STATS", "MinHAM"))

'UserList(UserIndex).Stats.SkillPts = val(leer.darvalor( "STATS", "SkillPtsLibres"))

'UserList(UserIndex).Stats.exp = val(leer.darvalor( "STATS", "EXP"))
'UserList(UserIndex).Stats.Elu = val(leer.darvalor( "STATS", "ELU"))
'UserList(UserIndex).Stats.ELV = val(leer.darvalor( "STATS", "ELV"))
'pluto:2.4.5
'UserList(UserIndex).Stats.PClan = val(leer.darvalor( "STATS", "PCLAN"))
'UserList(UserIndex).Stats.GTorneo = val(leer.darvalor( "STATS", "GTORNEO"))



'UserList(UserIndex).Stats.UsuariosMatados = val(leer.darvalor( "MUERTES", "UserMuertes"))
'UserList(UserIndex).Stats.CriminalesMatados = val(leer.darvalor( "MUERTES", "CrimMuertes"))
'UserList(UserIndex).Stats.NPCsMuertos = val(leer.darvalor( "MUERTES", "NpcsMuertes"))
'Exit Sub
'fallo:
'Call LogError("LOADUSERSTATS" & Err.Number & " D: " & Err.Description)

End Sub

Sub LoadUserReputacion(UserIndex As Integer, userfile As String)
'on error GoTo fallo
'UserList(UserIndex).Reputacion.AsesinoRep = val(leer.darvalor( "REP", "Asesino"))
'UserList(UserIndex).Reputacion.BandidoRep = val(leer.darvalor( "REP", "Dandido"))
'UserList(UserIndex).Reputacion.BurguesRep = val(leer.darvalor( "REP", "Burguesia"))
'UserList(UserIndex).Reputacion.LadronesRep = val(leer.darvalor( "REP", "Ladrones"))
'UserList(UserIndex).Reputacion.NobleRep = val(leer.darvalor( "REP", "Nobles"))
'UserList(UserIndex).Reputacion.PlebeRep = val(leer.darvalor( "REP", "Plebe"))
'UserList(UserIndex).Reputacion.Promedio = val(leer.darvalor( "REP", "Promedio"))
'pluto:2-3-04
'If UserList(UserIndex).Faccion.FuerzasCaos > 0 And UserList(UserIndex).Reputacion.Promedio >= 0 Then Call ExpulsarCaos(UserIndex)
'Exit Sub
'fallo:
'Call LogError("LOADUSERREPUTACION" & Err.Number & " D: " & Err.Description)


End Sub


Sub LoadUserInit(UserIndex As Integer, userfile As String, Name As String)

On Error GoTo fallo
Dim loopc As Integer
Dim ln As String
Dim Ln2 As String
'pluto:2.24

Dim Leer As New clsLeerInis
Leer.Abrir userfile

UserList(UserIndex).Faccion.ArmadaReal = val(Leer.DarValor("FACCIONES", "EjercitoReal"))
UserList(UserIndex).Faccion.FuerzasCaos = val(Leer.DarValor("FACCIONES", "EjercitoCaos"))
UserList(UserIndex).Faccion.CiudadanosMatados = val(Leer.DarValor("FACCIONES", "CiudMatados"))
UserList(UserIndex).Faccion.CriminalesMatados = val(Leer.DarValor("FACCIONES", "CrimMatados"))
UserList(UserIndex).Faccion.RecibioArmaduraCaos = val(Leer.DarValor("FACCIONES", "rArCaos"))
UserList(UserIndex).Faccion.RecibioArmaduraReal = val(Leer.DarValor("FACCIONES", "rArReal"))
UserList(UserIndex).Faccion.RecibioArmaduraLegion = val(Leer.DarValor("FACCIONES", "rArLegion"))
UserList(UserIndex).Faccion.RecibioExpInicialCaos = val(Leer.DarValor("FACCIONES", "rExCaos"))
UserList(UserIndex).Faccion.RecibioExpInicialReal = val(Leer.DarValor("FACCIONES", "rExReal"))
UserList(UserIndex).Faccion.RecompensasCaos = val(Leer.DarValor("FACCIONES", "recCaos"))
UserList(UserIndex).Faccion.RecompensasReal = val(Leer.DarValor("FACCIONES", "recReal"))
UserList(UserIndex).flags.Muerto = val(Leer.DarValor("FLAGS", "Muerto"))
UserList(UserIndex).flags.Escondido = val(Leer.DarValor("FLAGS", "Escondido"))
UserList(UserIndex).flags.Hambre = val(Leer.DarValor("FLAGS", "Hambre"))
UserList(UserIndex).flags.Sed = val(Leer.DarValor("FLAGS", "Sed"))
UserList(UserIndex).flags.Desnudo = val(Leer.DarValor("FLAGS", "Desnudo"))
UserList(UserIndex).Mision.estado = val(Leer.DarValor("QUEST", "Estado"))
UserList(UserIndex).Mision.TimeComienzo = Leer.DarValor("QUEST", "TimeC")
UserList(UserIndex).Mision.numero = val(Leer.DarValor("QUEST", "Numero"))
'pluto:7.0----------------------------------------------------------------
UserList(UserIndex).Mision.Actual1 = val(Leer.DarValor("QUEST", "Actual1"))
UserList(UserIndex).Mision.Actual2 = val(Leer.DarValor("QUEST", "Actual2"))
UserList(UserIndex).Mision.Actual3 = val(Leer.DarValor("QUEST", "Actual3"))
UserList(UserIndex).Mision.Actual4 = val(Leer.DarValor("QUEST", "Actual4"))
UserList(UserIndex).Mision.Actual5 = val(Leer.DarValor("QUEST", "Actual5"))
UserList(UserIndex).Mision.Actual6 = val(Leer.DarValor("QUEST", "Actual6"))
UserList(UserIndex).Mision.Actual7 = val(Leer.DarValor("QUEST", "Actual7"))
UserList(UserIndex).Mision.Actual8 = val(Leer.DarValor("QUEST", "Actual8"))
UserList(UserIndex).Mision.Actual9 = val(Leer.DarValor("QUEST", "Actual9"))
UserList(UserIndex).Mision.Actual10 = val(Leer.DarValor("QUEST", "Actual10"))
UserList(UserIndex).Mision.Actual11 = val(Leer.DarValor("QUEST", "Actual11"))
UserList(UserIndex).Mision.Actual12 = val(Leer.DarValor("QUEST", "Actual12"))
UserList(UserIndex).Mision.NpcQuest = val(Leer.DarValor("QUEST", "NpcQuest"))

For loopc = 1 To 5
UserList(UserIndex).Mision.NEnemigosConseguidos(loopc) = val(Leer.DarValor("QUEST", "EC" & loopc))
Next
UserList(UserIndex).Mision.PjConseguidos = val(Leer.DarValor("QUEST", "PJC"))
'------------------------------------------------------------------------
'UserList(UserIndex).Mision.Level = val(Leer.DarValor("QUEST", "Level"))
'UserList(UserIndex).Mision.Entrega = val(Leer.DarValor("QUEST", "Entrega"))
'UserList(UserIndex).Mision.Cantidad = val(Leer.DarValor("QUEST", "Cantidad"))
'UserList(UserIndex).Mision.Objeto = val(Leer.DarValor("QUEST", "Objeto"))
'UserList(UserIndex).Mision.Enemigo = val(Leer.DarValor("QUEST", "Enemigo"))
'UserList(UserIndex).Mision.clase = Leer.DarValor("QUEST", "Clase")
UserList(UserIndex).flags.Envenenado = val(Leer.DarValor("FLAGS", "Envenenado"))
UserList(UserIndex).flags.Morph = val(Leer.DarValor("FLAGS", "Morph"))
UserList(UserIndex).flags.Paralizado = val(Leer.DarValor("FLAGS", "Paralizado"))
UserList(UserIndex).flags.Angel = val(Leer.DarValor("FLAGS", "Angel"))
UserList(UserIndex).flags.Demonio = val(Leer.DarValor("FLAGS", "Demonio"))
'pluto:6.5
UserList(UserIndex).flags.Minotauro = val(Leer.DarValor("FLAGS", "Minotauro"))
UserList(UserIndex).flags.MinutosOnline = val(Leer.DarValor("FLAGS", "MinOn"))
'pluto:7.0
UserList(UserIndex).flags.Creditos = val(Leer.DarValor("FLAGS", "Creditos"))
UserList(UserIndex).flags.DragCredito1 = val(Leer.DarValor("FLAGS", "DragC1"))
UserList(UserIndex).flags.DragCredito2 = val(Leer.DarValor("FLAGS", "DragC2"))
UserList(UserIndex).flags.DragCredito3 = val(Leer.DarValor("FLAGS", "DragC3"))
UserList(UserIndex).flags.DragCredito4 = val(Leer.DarValor("FLAGS", "DragC4"))
UserList(UserIndex).flags.DragCredito5 = val(Leer.DarValor("FLAGS", "DragC5"))
'pluto:6.9
UserList(UserIndex).flags.DragCredito6 = val(Leer.DarValor("FLAGS", "DragC6"))

UserList(UserIndex).flags.Elixir = val(Leer.DarValor("FLAGS", "Elixir"))
'---------------------

UserList(UserIndex).flags.Navegando = val(Leer.DarValor("FLAGS", "Navegando"))
UserList(UserIndex).flags.Montura = val(Leer.DarValor("FLAGS", "Montura"))
UserList(UserIndex).flags.ClaseMontura = val(Leer.DarValor("FLAGS", "ClaseMontura"))
UserList(UserIndex).Counters.Pena = val(Leer.DarValor("COUNTERS", "Pena"))
UserList(UserIndex).EmailActual = Leer.DarValor("CONTACTO", "EmailActual")
UserList(UserIndex).Email = Leer.DarValor("CONTACTO", "Email")
UserList(UserIndex).Remorted = Leer.DarValor("INIT", "RAZAREMORT")
'pluto:6.0A
UserList(UserIndex).BD = val(Leer.DarValor("INIT", "BD"))

UserList(UserIndex).Genero = Leer.DarValor("INIT", "Genero")
UserList(UserIndex).clase = Leer.DarValor("INIT", "Clase")
UserList(UserIndex).raza = Leer.DarValor("INIT", "Raza")
UserList(UserIndex).Hogar = Leer.DarValor("INIT", "Hogar")
UserList(UserIndex).Char.Heading = val(Leer.DarValor("INIT", "Heading"))
UserList(UserIndex).Esposa = Trim$(Leer.DarValor("INIT", "Esposa"))
UserList(UserIndex).Paquete = 0
'pluto:2.24-------------------------------
'Dim filexx As String

'If UserList(UserIndex).Esposa = "0" Then
'filexx = "C:\Esposas\Charfile\" & Left$(UCase$(Name), 1) & "\" & UCase$(Name) & ".chr"
'UserList(UserIndex).Esposa = GetVar(filexx, "INIT", "Esposa")
'End If
'-----------------------------------------

UserList(UserIndex).Nhijos = val(Leer.DarValor("INIT", "Nhijos"))

For loopc = 1 To 5
UserList(UserIndex).Hijo(loopc) = Trim$(Leer.DarValor("INIT", "Hijo" & loopc))
Next

UserList(UserIndex).Amor = val(Leer.DarValor("INIT", "Amor"))
UserList(UserIndex).Embarazada = val(Leer.DarValor("INIT", "Embarazada"))
UserList(UserIndex).Bebe = val(Leer.DarValor("INIT", "Bebe"))
UserList(UserIndex).NombreDelBebe = Trim$(Leer.DarValor("INIT", "NombreDelBebe"))
UserList(UserIndex).Padre = Trim$(Leer.DarValor("INIT", "Padre"))
UserList(UserIndex).Madre = Trim$(Leer.DarValor("INIT", "Madre"))
UserList(UserIndex).OrigChar.Head = val(Leer.DarValor("INIT", "Head"))
UserList(UserIndex).OrigChar.Body = val(Leer.DarValor("INIT", "Body"))
UserList(UserIndex).OrigChar.WeaponAnim = val(Leer.DarValor("INIT", "Arma"))
UserList(UserIndex).OrigChar.ShieldAnim = val(Leer.DarValor("INIT", "Escudo"))
UserList(UserIndex).OrigChar.CascoAnim = val(Leer.DarValor("INIT", "Casco"))
UserList(UserIndex).OrigChar.Botas = val(Leer.DarValor("INIT", "Botas"))

'UserList(UserIndex).Faccion.ArmadaReal = val(leer.darvalor( "FACCIONES", "EjercitoReal"))
'UserList(UserIndex).Faccion.FuerzasCaos = val(leer.darvalor( "FACCIONES", "EjercitoCaos"))
'UserList(UserIndex).Faccion.CiudadanosMatados = val(leer.darvalor( "FACCIONES", "CiudMatados"))
'UserList(UserIndex).Faccion.CriminalesMatados = val(leer.darvalor( "FACCIONES", "CrimMatados"))
'UserList(UserIndex).Faccion.RecibioArmaduraCaos = val(leer.darvalor( "FACCIONES", "rArCaos"))
'UserList(UserIndex).Faccion.RecibioArmaduraReal = val(leer.darvalor( "FACCIONES", "rArReal"))
'pluto:2.3
'UserList(UserIndex).Faccion.RecibioArmaduraLegion = val(leer.darvalor( "FACCIONES", "rArLegion"))
'UserList(UserIndex).Faccion.RecibioExpInicialCaos = val(leer.darvalor( "FACCIONES", "rExCaos"))
'UserList(UserIndex).Faccion.RecibioExpInicialReal = val(leer.darvalor( "FACCIONES", "rExReal"))

'UserList(UserIndex).Faccion.RecompensasCaos = val(leer.darvalor( "FACCIONES", "recCaos"))
'UserList(UserIndex).Faccion.RecompensasReal = val(leer.darvalor( "FACCIONES", "recReal"))

'UserList(UserIndex).flags.Muerto = val(leer.darvalor( "FLAGS", "Muerto"))
'UserList(UserIndex).flags.Escondido = val(leer.darvalor( "FLAGS", "Escondido"))

'UserList(UserIndex).flags.Hambre = val(leer.darvalor( "FLAGS", "Hambre"))
'UserList(UserIndex).flags.Sed = val(leer.darvalor( "FLAGS", "Sed"))
'UserList(UserIndex).flags.Desnudo = val(leer.darvalor( "FLAGS", "Desnudo"))

'pluto:hoy
'UserList(UserIndex).Mision.estado = val(leer.darvalor( "QUEST", "Estado"))
'UserList(UserIndex).Mision.Numero = val(leer.darvalor( "QUEST", "Numero"))
'UserList(UserIndex).Mision.level = val(leer.darvalor( "QUEST", "Level"))
'UserList(UserIndex).Mision.Entrega = val(leer.darvalor( "QUEST", "Entrega"))
'UserList(UserIndex).Mision.Cantidad = val(leer.darvalor( "QUEST", "Cantidad"))
'UserList(UserIndex).Mision.Objeto = val(leer.darvalor( "QUEST", "Objeto"))
'UserList(UserIndex).Mision.Enemigo = val(leer.darvalor( "QUEST", "Enemigo"))
'UserList(UserIndex).Mision.clase = leer.darvalor( "QUEST", "Clase")
'[Tite]Party
UserList(UserIndex).flags.party = False
UserList(UserIndex).flags.partyNum = 0
UserList(UserIndex).flags.invitado = ""
'[\Tite]
'UserList(UserIndex).flags.Envenenado = val(leer.darvalor( "FLAGS", "Envenenado"))
'UserList(UserIndex).flags.Morph = val(leer.darvalor( "FLAGS", "Morph"))
'UserList(UserIndex).flags.Paralizado = val(leer.darvalor( "FLAGS", "Paralizado"))
'UserList(UserIndex).flags.Angel = val(leer.darvalor( "FLAGS", "Angel"))
'UserList(UserIndex).flags.Demonio = val(leer.darvalor( "FLAGS", "Demonio"))

'UserList(UserIndex).flags.Navegando = val(leer.darvalor( "FLAGS", "Navegando"))
'pluto:2.3
'UserList(UserIndex).flags.Montura = val(leer.darvalor( "FLAGS", "Montura"))
'UserList(UserIndex).flags.ClaseMontura = val(leer.darvalor( "FLAGS", "ClaseMontura"))


'UserList(UserIndex).Counters.Pena = val(leer.darvalor( "COUNTERS", "Pena"))
'pluto:2.10
'UserList(UserIndex).EmailActual = leer.darvalor( "CONTACTO", "EmailActual")

'UserList(UserIndex).Email = leer.darvalor( "CONTACTO", "Email")
'UserList(UserIndex).Remorted = leer.darvalor( "INIT", "RAZAREMORT")
'UserList(UserIndex).Genero = leer.darvalor( "INIT", "Genero")
'UserList(UserIndex).clase = leer.darvalor( "INIT", "Clase")
'UserList(UserIndex).raza = leer.darvalor( "INIT", "Raza")
'UserList(UserIndex).Hogar = leer.darvalor( "INIT", "Hogar")
'UserList(UserIndex).Char.Heading = val(leer.darvalor( "INIT", "Heading"))
'pluto:2.14--------
'UserList(UserIndex).Esposa = leer.darvalor( "INIT", "Esposa")
'UserList(UserIndex).Nhijos = val(leer.darvalor( "INIT", "Nhijos"))
'pluto:2.15
'Dim X As Byte
'For X = 1 To 5
'UserList(UserIndex).Hijo(X) = leer.darvalor( "INIT", "Hijo" & X)
'Next
'UserList(UserIndex).Amor = val(leer.darvalor( "INIT", "Amor"))
'UserList(UserIndex).Embarazada = val(leer.darvalor( "INIT", "Embarazada"))
'UserList(UserIndex).Bebe = val(leer.darvalor( "INIT", "Bebe"))
'UserList(UserIndex).NombreDelBebe = leer.darvalor( "INIT", "NombreDelBebe")
'UserList(UserIndex).Padre = leer.darvalor( "INIT", "Padre")
'UserList(UserIndex).Madre = leer.darvalor( "INIT", "Madre")
'-------------------

'UserList(UserIndex).OrigChar.Head = val(leer.darvalor( "INIT", "Head"))
'UserList(UserIndex).OrigChar.Body = val(leer.darvalor( "INIT", "Body"))
'UserList(UserIndex).OrigChar.WeaponAnim = val(leer.darvalor( "INIT", "Arma"))
'UserList(UserIndex).OrigChar.ShieldAnim = val(leer.darvalor( "INIT", "Escudo"))
'UserList(UserIndex).OrigChar.CascoAnim = val(leer.darvalor( "INIT", "Casco"))
'[GAU]
'UserList(UserIndex).OrigChar.Botas = val(leer.darvalor( "INIT", "Botas"))
'[GAU]
UserList(UserIndex).OrigChar.Heading = SOUTH


If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char = UserList(UserIndex).OrigChar
Else
        If Not Criminal(UserIndex) Then UserList(UserIndex).Char.Body = iCuerpoMuerto Else UserList(UserIndex).Char.Body = iCuerpoMuerto2
        If Not Criminal(UserIndex) Then UserList(UserIndex).Char.Head = iCabezaMuerto Else UserList(UserIndex).Char.Head = iCabezaMuerto2
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.CascoAnim = NingunCasco
'[GAU]
        UserList(UserIndex).Char.Botas = NingunBota
        '[GAU]
End If


UserList(UserIndex).Desc = Trim$(Leer.DarValor("INIT", "Desc"))
'UserList(UserIndex).Desc = Leer.DarValor("INIT", "Desc")



UserList(UserIndex).Pos.Map = val(ReadField(1, Leer.DarValor("INIT", "Position"), 45))
UserList(UserIndex).Pos.X = val(ReadField(2, Leer.DarValor("INIT", "Position"), 45))
UserList(UserIndex).Pos.Y = val(ReadField(3, Leer.DarValor("INIT", "Position"), 45))

'Delzak
'If UserList(UserIndex).Pos.Map <> 0 Then Call BuscaPosicionValida(UserIndex)

'UserList(UserIndex).Invent.NroItems = leer.darvalor( "Inventory", "CantidadItems")
UserList(UserIndex).Invent.NroItems = Leer.DarValor("Inventory", "CantidadItems")
Dim loopd As Integer

'[KEVIN]--------------------------------------------------------------------




'***********************************************************************************
'pluto:7.0 quito todo esto lo paso a cuentas

'UserList(UserIndex).BancoInvent.NroItems = val(Leer.DarValor("BancoInventory", "CantidadItems"))

'Lista de objetos del banco
'For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
   
 '   ln2 = Leer.DarValor("BancoInventory", "Obj" & loopd)
    
  '  UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex = val(ReadField(1, ln2, 45))
   ' UserList(UserIndex).BancoInvent.Object(loopd).Amount = val(ReadField(2, ln2, 45))
'Next loopd
'------------------------------------------------------------------------------------








'[/KEVIN]*****************************************************************************


'Lista de objetos
For loopc = 1 To MAX_INVENTORY_SLOTS
    'ln = leer.darvalor( "Inventory", "Obj" & LoopC)
    ln = Leer.DarValor("Inventory", "Obj" & loopc)
    
    UserList(UserIndex).Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
    UserList(UserIndex).Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))
    UserList(UserIndex).Invent.Object(loopc).Equipped = val(ReadField(3, ln, 45))
Next loopc

'Obtiene el indice-objeto del arma
'UserList(UserIndex).Invent.WeaponEqpSlot = val(leer.darvalor( "Inventory", "WeaponEqpSlot"))
UserList(UserIndex).Invent.WeaponEqpSlot = val(Leer.DarValor("Inventory", "WeaponEqpSlot"))

If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
End If
'Obtiene el indice-objeto del anillo
'UserList(UserIndex).Invent.AnilloEqpSlot = val(leer.darvalor( "Inventory", "AnilloEqpSlot"))
UserList(UserIndex).Invent.AnilloEqpSlot = val(Leer.DarValor("Inventory", "AnilloEqpSlot"))

If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
    UserList(UserIndex).Invent.AnilloEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.AnilloEqpSlot).ObjIndex
End If
'Obtiene el indice-objeto del armadura
'UserList(UserIndex).Invent.ArmourEqpSlot = val(leer.darvalor( "Inventory", "ArmourEqpSlot"))
UserList(UserIndex).Invent.ArmourEqpSlot = val(Leer.DarValor("Inventory", "ArmourEqpSlot"))

If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
    UserList(UserIndex).flags.Desnudo = 0
Else
    UserList(UserIndex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
'UserList(UserIndex).Invent.EscudoEqpSlot = val(leer.darvalor( "Inventory", "EscudoEqpSlot"))
UserList(UserIndex).Invent.EscudoEqpSlot = val(Leer.DarValor("Inventory", "EscudoEqpSlot"))

If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
'UserList(UserIndex).Invent.CascoEqpSlot = val(leer.darvalor( "Inventory", "CascoEqpSlot"))
UserList(UserIndex).Invent.CascoEqpSlot = val(Leer.DarValor("Inventory", "CascoEqpSlot"))

If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
End If
'[GAU]
'Obtiene el indice-objeto de las botas
'UserList(UserIndex).Invent.BotaEqpSlot = val(leer.darvalor( "Inventory", "BotaEqpSlot"))
UserList(UserIndex).Invent.BotaEqpSlot = val(Leer.DarValor("Inventory", "BotaEqpSlot"))
If UserList(UserIndex).Invent.BotaEqpSlot > 0 Then
    UserList(UserIndex).Invent.BotaEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BotaEqpSlot).ObjIndex
End If
'[GAU]
'Obtiene el indice-objeto barco
'UserList(UserIndex).Invent.BarcoSlot = val(leer.darvalor( "Inventory", "BarcoSlot"))
UserList(UserIndex).Invent.BarcoSlot = val(Leer.DarValor("Inventory", "BarcoSlot"))

If UserList(UserIndex).Invent.BarcoSlot > 0 Then
    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
'UserList(UserIndex).Invent.MunicionEqpSlot = val(leer.darvalor( "Inventory", "MunicionSlot"))
UserList(UserIndex).Invent.MunicionEqpSlot = val(Leer.DarValor("Inventory", "MunicionSlot"))

If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex
End If

'UserList(UserIndex).NroMacotas = val(leer.darvalor( "Mascotas", "NroMascotas"))
UserList(UserIndex).NroMacotas = val(Leer.DarValor("Mascotas", "NroMascotas"))
If UserList(UserIndex).NroMacotas < 0 Then UserList(UserIndex).NroMacotas = 0

'Lista de objetos
For loopc = 1 To MAXMASCOTAS
   ' UserList(UserIndex).MascotasType(LoopC) = val(leer.darvalor( "Mascotas", "Mas" & LoopC))
   UserList(UserIndex).MascotasType(loopc) = val(Leer.DarValor("Mascotas", "Mas" & loopc))
Next loopc

'UserList(UserIndex).GuildInfo.FundoClan = val(leer.darvalor( "Guild", "FundoClan"))
'UserList(UserIndex).GuildInfo.EsGuildLeader = val(leer.darvalor( "Guild", "EsGuildLeader"))
'UserList(UserIndex).GuildInfo.Echadas = val(leer.darvalor( "Guild", "Echadas"))
'UserList(UserIndex).GuildInfo.Solicitudes = val(leer.darvalor( "Guild", "Solicitudes"))
'UserList(UserIndex).GuildInfo.SolicitudesRechazadas = val(leer.darvalor( "Guild", "SolicitudesRechazadas"))
'UserList(UserIndex).GuildInfo.VecesFueGuildLeader = val(leer.darvalor( "Guild", "VecesFueGuildLeader"))
'UserList(UserIndex).GuildInfo.YaVoto = val(leer.darvalor( "Guild", "YaVoto"))
'UserList(UserIndex).GuildInfo.ClanesParticipo = val(leer.darvalor( "Guild", "ClanesParticipo"))
'UserList(UserIndex).GuildInfo.GuildPoints = val(leer.darvalor( "Guild", "GuildPts"))
'UserList(UserIndex).GuildInfo.ClanFundado = leer.darvalor( "Guild", "ClanFundado")
'UserList(UserIndex).GuildInfo.GuildName = leer.darvalor( "Guild", "GuildName")

UserList(UserIndex).GuildInfo.FundoClan = val(Leer.DarValor("Guild", "FundoClan"))
UserList(UserIndex).GuildInfo.EsGuildLeader = val(Leer.DarValor("Guild", "EsGuildLeader"))
UserList(UserIndex).GuildInfo.Echadas = val(Leer.DarValor("Guild", "Echadas"))
UserList(UserIndex).GuildInfo.Solicitudes = val(Leer.DarValor("Guild", "Solicitudes"))
UserList(UserIndex).GuildInfo.SolicitudesRechazadas = val(Leer.DarValor("Guild", "SolicitudesRechazadas"))
UserList(UserIndex).GuildInfo.VecesFueGuildLeader = val(Leer.DarValor("Guild", "VecesFueGuildLeader"))
UserList(UserIndex).GuildInfo.YaVoto = val(Leer.DarValor("Guild", "Yavoto"))
UserList(UserIndex).GuildInfo.ClanesParticipo = val(Leer.DarValor("Guild", "ClanesParticipo"))
UserList(UserIndex).GuildInfo.GuildPoints = val(Leer.DarValor("Guild", "GuildPts"))
UserList(UserIndex).GuildInfo.ClanFundado = Trim$(Leer.DarValor("Guild", "ClanFundado"))
UserList(UserIndex).GuildInfo.GuildName = Trim$(Leer.DarValor("Guild", "GuildName"))

'loaduserstats-------------------------------
For loopc = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributos(loopc) = Leer.DarValor("ATRIBUTOS", "AT" & loopc)
  UserList(UserIndex).Stats.UserAtributosBackUP(loopc) = UserList(UserIndex).Stats.UserAtributos(loopc)
Next
'pluto:7.0
UserList(UserIndex).UserDañoProyetilesRaza = val(Leer.DarValor("PORC", "P1"))
UserList(UserIndex).UserDañoArmasRaza = val(Leer.DarValor("PORC", "P2"))
UserList(UserIndex).UserDañoMagiasRaza = val(Leer.DarValor("PORC", "P3"))
UserList(UserIndex).UserDefensaMagiasRaza = val(Leer.DarValor("PORC", "P4"))
UserList(UserIndex).UserEvasiónRaza = val(Leer.DarValor("PORC", "P5"))
UserList(UserIndex).UserDefensaEscudos = val(Leer.DarValor("PORC", "P6"))

If UserList(UserIndex).UserDañoProyetilesRaza + UserList(UserIndex).UserDañoArmasRaza + UserList(UserIndex).UserDañoMagiasRaza + UserList(UserIndex).UserDefensaMagiasRaza + UserList(UserIndex).UserEvasiónRaza + UserList(UserIndex).UserDefensaEscudos > 15 Then
    UserList(UserIndex).UserDañoArmasRaza = 5
    UserList(UserIndex).UserDañoMagiasRaza = 5
    UserList(UserIndex).UserDefensaMagiasRaza = 5
End If


For loopc = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(loopc) = val(Leer.DarValor("SKILLS", "SK" & loopc))
Next

For loopc = 1 To MAXUSERHECHIZOS
  UserList(UserIndex).Stats.UserHechizos(loopc) = val(Leer.DarValor("Hechizos", "H" & loopc))
Next
'pluto:2-3-04
UserList(UserIndex).Stats.Puntos = val(Leer.DarValor("STATS", "PUNTOS"))

UserList(UserIndex).Stats.GLD = val(Leer.DarValor("STATS", "GLD"))
UserList(UserIndex).Remort = val(Leer.DarValor("STATS", "REMORT"))
UserList(UserIndex).Stats.Banco = val(Leer.DarValor("STATS", "BANCO"))

UserList(UserIndex).Stats.MET = val(Leer.DarValor("STATS", "MET"))
UserList(UserIndex).Stats.MaxHP = val(Leer.DarValor("STATS", "MaxHP"))
UserList(UserIndex).Stats.MinHP = val(Leer.DarValor("STATS", "MinHP"))

UserList(UserIndex).Stats.FIT = val(Leer.DarValor("STATS", "FIT"))
UserList(UserIndex).Stats.MinSta = val(Leer.DarValor("STATS", "MinSTA"))
UserList(UserIndex).Stats.MaxSta = val(Leer.DarValor("STATS", "MaxSTA"))

UserList(UserIndex).Stats.MaxMAN = val(Leer.DarValor("STATS", "MaxMAN"))
UserList(UserIndex).Stats.MinMAN = val(Leer.DarValor("STATS", "MinMAN"))

UserList(UserIndex).Stats.MaxHIT = val(Leer.DarValor("STATS", "MaxHIT"))
UserList(UserIndex).Stats.MinHIT = val(Leer.DarValor("STATS", "MinHIT"))

UserList(UserIndex).Stats.MaxAGU = val(Leer.DarValor("STATS", "MaxAGU"))
UserList(UserIndex).Stats.MinAGU = val(Leer.DarValor("STATS", "MinAGU"))

UserList(UserIndex).Stats.MaxHam = val(Leer.DarValor("STATS", "MaxHAM"))
UserList(UserIndex).Stats.MinHam = val(Leer.DarValor("STATS", "MinHAM"))

UserList(UserIndex).Stats.SkillPts = val(Leer.DarValor("STATS", "SkillPtsLibres"))

UserList(UserIndex).Stats.exp = val(Leer.DarValor("STATS", "EXP"))
UserList(UserIndex).Stats.Elu = val(Leer.DarValor("STATS", "ELU"))
UserList(UserIndex).Stats.ELV = val(Leer.DarValor("STATS", "ELV"))
UserList(UserIndex).Stats.LibrosUsados = val(Leer.DarValor("STATS", "LIBROSUSADOS"))
UserList(UserIndex).Stats.Fama = val(Leer.DarValor("STATS", "FAMA"))
'pluto:2.4.5
UserList(UserIndex).Stats.PClan = val(Leer.DarValor("STATS", "PCLAN"))
UserList(UserIndex).Stats.GTorneo = val(Leer.DarValor("STATS", "GTORNEO"))



UserList(UserIndex).Stats.UsuariosMatados = val(Leer.DarValor("MUERTES", "UserMuertes"))
UserList(UserIndex).Stats.CriminalesMatados = val(Leer.DarValor("MUERTES", "CrimMuertes"))
UserList(UserIndex).Stats.NPCsMuertos = val(Leer.DarValor("MUERTES", "NpcsMuertes"))
'--------------------------------------------

'Delzak-----------------------------------------
'...............................................
'              SISTEMA PREMIOS
'...............................................
'--Modificado por Pluto:7.0---------------------

'Stats de premios por matar NPCs
For loopc = 1 To 34
UserList(UserIndex).Stats.PremioNPC(loopc) = val(Leer.DarValor("PREMIOS", "L" & loopc))
Next
'--------------------------------------------


'PLUTO 6.0A  loadusermonturas ---------------------------

 UserList(UserIndex).Nmonturas = val(Leer.DarValor("MONTURAS", "NroMonturas"))

 
 Dim n As Byte

    For n = 1 To 3
If val(Leer.DarValor("MONTURA" & n, "TIPO")) > 0 Then
    loopc = val(Leer.DarValor("MONTURA" & n, "TIPO"))
    
UserList(UserIndex).Montura.Tipo(loopc) = val(Leer.DarValor("MONTURA" & n, "TIPO"))
UserList(UserIndex).Montura.Nivel(loopc) = val(Leer.DarValor("MONTURA" & n, "NIVEL"))
UserList(UserIndex).Montura.exp(loopc) = val(Leer.DarValor("MONTURA" & n, "EXP"))
UserList(UserIndex).Montura.Elu(loopc) = val(Leer.DarValor("MONTURA" & n, "ELU"))
UserList(UserIndex).Montura.Vida(loopc) = val(Leer.DarValor("MONTURA" & n, "VIDA"))
UserList(UserIndex).Montura.Golpe(loopc) = val(Leer.DarValor("MONTURA" & n, "GOLPE"))
UserList(UserIndex).Montura.Nombre(loopc) = Trim$(Leer.DarValor("MONTURA" & n, "NOMBRE"))
UserList(UserIndex).Montura.AtCuerpo(loopc) = val(Leer.DarValor("MONTURA" & n, "ATCUERPO"))
UserList(UserIndex).Montura.Defcuerpo(loopc) = val(Leer.DarValor("MONTURA" & n, "DEFCUERPO"))
UserList(UserIndex).Montura.AtFlechas(loopc) = val(Leer.DarValor("MONTURA" & n, "ATFLECHAS"))
UserList(UserIndex).Montura.DefFlechas(loopc) = val(Leer.DarValor("MONTURA" & n, "DEFFLECHAS"))
UserList(UserIndex).Montura.AtMagico(loopc) = val(Leer.DarValor("MONTURA" & n, "ATMAGICO"))
UserList(UserIndex).Montura.DefMagico(loopc) = val(Leer.DarValor("MONTURA" & n, "DEFMAGICO"))
UserList(UserIndex).Montura.Evasion(loopc) = val(Leer.DarValor("MONTURA" & n, "EVASION"))
UserList(UserIndex).Montura.Libres(loopc) = val(Leer.DarValor("MONTURA" & n, "LIBRES"))
UserList(UserIndex).Montura.index(loopc) = n


    
    
    End If
    Next n






'---------------------------------------------

'loaduserreputacion---------------------------
UserList(UserIndex).Reputacion.AsesinoRep = val(Leer.DarValor("REP", "Asesino"))
UserList(UserIndex).Reputacion.BandidoRep = val(Leer.DarValor("REP", "Bandido"))
UserList(UserIndex).Reputacion.BurguesRep = val(Leer.DarValor("REP", "Burguesia"))
UserList(UserIndex).Reputacion.LadronesRep = val(Leer.DarValor("REP", "Ladrones"))
UserList(UserIndex).Reputacion.NobleRep = val(Leer.DarValor("REP", "Nobles"))
UserList(UserIndex).Reputacion.PlebeRep = val(Leer.DarValor("REP", "Plebe"))
UserList(UserIndex).Reputacion.Promedio = val(Leer.DarValor("REP", "Promedio"))
'pluto:2-3-04
If UserList(UserIndex).Faccion.FuerzasCaos > 0 And UserList(UserIndex).Reputacion.Promedio >= 0 Then Call ExpulsarCaos(UserIndex)
'------------------------------------------------------





Exit Sub
fallo:
Call LogError("LOADUSERINIT" & Err.number & " D: " & Err.Description)

End Sub





Function GetVar(file As String, Main As String, Var As String) As String
On Error GoTo fallo
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), file
  
GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  Exit Function
fallo:
Call LogError("GETVAR" & Err.number & " D: " & Err.Description)

End Function

Sub CargarBackUp()

'Call LogTarea("Sub CargarBackUp")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim Map As Integer
Dim loopc As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempInt As Integer
Dim SaveAs As String
Dim npcfile As String
Dim Porc As Long
Dim FileNamE As String
Dim C$
    
On Error GoTo man

 
NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumMaps
frmCargando.cargar.value = 0

MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo

Dim buffer(1 To ((YMaxMapSize - YMinMapSize + 1) * (XMaxMapSize - XMinMapSize + 1))) As TileMap
Dim buffer2(1 To ((YMaxMapSize - YMinMapSize + 1) * (XMaxMapSize - XMinMapSize + 1))) As TileInf
Dim idx As Integer

For Map = 1 To NumMaps
    
    FileNamE = App.Path & "\WorldBackUp\Map" & Map & ".map"
    
    If FileExist(FileNamE, vbNormal) Then
        Open App.Path & "\WorldBackUp\Map" & Map & ".map" For Binary As #1
        Open App.Path & "\WorldBackUp\Map" & Map & ".inf" For Binary As #2
        C$ = App.Path & "\WorldBackUp\Map" & Map & ".dat"
    Else
        Open App.Path & MapPath & "Mapa" & Map & ".map" For Binary As #1
        Open App.Path & MapPath & "Mapa" & Map & ".inf" For Binary As #2
        C$ = App.Path & MapPath & "Mapa" & Map & ".dat"
    End If
    
    Seek #1, 1
    Seek #2, 1
    'map Header
    Get #1, , MapInfo(Map).MapVersion
    Get #1, , MiCabecera
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    'inf Header
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    'Load arrays
                   
    Get #1, , buffer
    Get #2, , buffer2
    
    
    idx = 1
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            MapData(Map, X, Y).Blocked = buffer(idx).bloqueado
            MapData(Map, X, Y).Graphic(1) = buffer(idx).grafs(1)
            MapData(Map, X, Y).Graphic(2) = buffer(idx).grafs(2)
            MapData(Map, X, Y).Graphic(3) = buffer(idx).grafs(3)
            MapData(Map, X, Y).Graphic(4) = buffer(idx).grafs(4)
            MapData(Map, X, Y).trigger = buffer(idx).trigger
            
            MapData(Map, X, Y).TileExit.Map = buffer2(idx).dest_mapa
            MapData(Map, X, Y).TileExit.X = buffer2(idx).dest_x
            MapData(Map, X, Y).TileExit.Y = buffer2(idx).dest_y
            
            MapData(Map, X, Y).NpcIndex = buffer2(idx).npc
            If MapData(Map, X, Y).NpcIndex > 0 Then
                
                If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                       
                Else
                        npcfile = DatPath & "NPCs.dat"
                End If
                
                'Si el npc debe hacer respawn en la pos
                'original la guardamos
                If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                Else
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                End If
                
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                
                'Si existe el backup lo cargamos
                        If Npclist(MapData(Map, X, Y).NpcIndex).flags.BackUp = 1 Then
                                'cargamos el nuevo del backup
                                Call CargarNpcBackUp(MapData(Map, X, Y).NpcIndex, Npclist(MapData(Map, X, Y).NpcIndex).numero)
                                
                        End If
                        
                Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
            End If

            If buffer2(idx).obj_ind > 0 And buffer2(idx).obj_ind <= UBound(ObjData) Then
                MapData(Map, X, Y).OBJInfo.ObjIndex = buffer2(idx).obj_ind
                MapData(Map, X, Y).OBJInfo.Amount = buffer2(idx).obj_cant
            Else
                MapData(Map, X, Y).OBJInfo.ObjIndex = 0
                MapData(Map, X, Y).OBJInfo.Amount = 0
            End If
            
            idx = idx + 1
        Next X
    Next Y

    Close #1
    Close #2
    MapInfo(Map).Name = GetVar(C$, "Mapa" & Map, "Name")
    MapInfo(Map).Music = GetVar(C$, "Mapa" & Map, "MusicNum")
    MapInfo(Map).Dueño = val(GetVar(C$, "Mapa" & Map, "Dueño"))
'pluto:6.0A
 MapInfo(Map).Resucitar = val(GetVar(C$, "Mapa" & Map, "Resucitar"))
  MapInfo(Map).Invisible = val(GetVar(C$, "Mapa" & Map, "Invisible"))
   MapInfo(Map).Mascotas = val(GetVar(C$, "Mapa" & Map, "Mascotas"))
    MapInfo(Map).Domar = val(GetVar(C$, "Mapa" & Map, "Domar"))
    MapInfo(Map).Insegura = val(GetVar(C$, "Mapa" & Map, "Insegura"))
     MapInfo(Map).Lluvia = val(GetVar(C$, "Mapa" & Map, "Lluvia"))
    MapInfo(Map).Monturas = val(GetVar(C$, "Mapa" & Map, "Monturas"))
   
   
   ' MapInfo(Map).MagiaSinEfecto = val(GetVar(c$, "Mapa" & Map, "MagiaSinEfecto"))
    'MapInfo(Map).NoEncriptarMP = val(GetVar(c$, "Mapa" & Map, "NoEncriptarMP"))
    MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(C$, "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(C$, "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(C$, "Mapa" & Map, "StartPos"), 45))
    If val(GetVar(C$, "Mapa" & Map, "Pk")) = 0 Then
          MapInfo(Map).Pk = True
    Else
          MapInfo(Map).Pk = False
    End If
    MapInfo(Map).Restringir = GetVar(C$, "Mapa" & Map, "Restringir")
    
    MapInfo(Map).BackUp = val(GetVar(C$, "Mapa" & Map, "BackUp"))
    MapInfo(Map).Terreno = GetVar(C$, "Mapa" & Map, "Terreno")
    MapInfo(Map).Zona = GetVar(C$, "Mapa" & Map, "Zona")
    
    
    frmCargando.cargar.value = frmCargando.cargar.value + 1
    
    DoEvents
Next Map

FrmStat.Visible = False


Exit Sub

man:
    MsgBox ("Error durante la carga de mapas.")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

  

End Sub

Sub LoadMapData()


'Call LogTarea("Sub LoadMapData")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas."

Dim Map As Integer
Dim loopc As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempInt As Integer
Dim npcfile As String

On Error GoTo man

NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))


frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumMaps
frmCargando.cargar.value = 0

MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
Dim Calcu As Double
  
For Map = 1 To NumMaps
    DoEvents
    Calcu = Map
    Calcu = Calcu * 100
    Calcu = Calcu / NumMaps
    frmCargando.Label1(2).Caption = "Mapa: (" & Map & "/" & NumMaps & ") " & Round(Calcu, 1) & "%"

    
    Open App.Path & MapPath & "Mapa" & Map & ".map" For Binary As #1
    Seek #1, 1
    
    'inf
    Open App.Path & MapPath & "Mapa" & Map & ".inf" For Binary As #2
    Seek #2, 1
    
     'map Header
    Get #1, , MapInfo(Map).MapVersion
    Get #1, , MiCabecera
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt

    'inf Header
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
        
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            '.dat file
            Get #1, , MapData(Map, X, Y).Blocked
            
            For loopc = 1 To 4
                Get #1, , MapData(Map, X, Y).Graphic(loopc)
            Next loopc
            
            Get #1, , MapData(Map, X, Y).trigger
            Get #1, , TempInt
            
                                
            '.inf file
            Get #2, , MapData(Map, X, Y).TileExit.Map
            Get #2, , MapData(Map, X, Y).TileExit.X
            Get #2, , MapData(Map, X, Y).TileExit.Y
            
            'Get and make NPC
            Get #2, , MapData(Map, X, Y).NpcIndex
            If MapData(Map, X, Y).NpcIndex > 0 Then
                
                If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                    'quitar esto----------------------------
                       'Dim NpcUsado(1 To 1000) As String
                       ' NpcUsado(MapData(Map, X, Y).NpcIndex) = NpcUsado(MapData(Map, X, Y).NpcIndex) & Map & ","
                        '-----------------------------------------
                Else
                        npcfile = DatPath & "NPCs.dat"
                End If
                
                'Si el npc debe hacer respawn en la pos
                'original la guardamos
                If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                Else
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                End If
                
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                
                Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
            End If

            'Get and make Object
            Get #2, , MapData(Map, X, Y).OBJInfo.ObjIndex
            Get #2, , MapData(Map, X, Y).OBJInfo.Amount

            'Space holder for future expansion (Objects, ect.
            Get #2, , DummyInt
            Get #2, , DummyInt
        
        Next X
    Next Y

   
    Close #1
    Close #2

  
    MapInfo(Map).Name = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Name")
    MapInfo(Map).Music = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "MusicNum")
    
    MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    
    If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Pk")) = 0 Then
        MapInfo(Map).Pk = True
    Else
        MapInfo(Map).Pk = False
    End If
    
    
    MapInfo(Map).Terreno = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Terreno")

    MapInfo(Map).Zona = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Zona")
    
    MapInfo(Map).Restringir = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Restringir")
'pluto:6.0A
MapInfo(Map).Resucitar = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Resucitar"))
MapInfo(Map).Invisible = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Invisible"))
MapInfo(Map).Mascotas = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Mascotas"))
MapInfo(Map).Insegura = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Insegura"))
MapInfo(Map).Domar = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Domar"))
MapInfo(Map).Monturas = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Monturas"))
MapInfo(Map).Lluvia = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Lluvia"))

    MapInfo(Map).BackUp = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "BACKUP"))
'pluto:2.17
MapInfo(Map).Dueño = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Dueño"))
MapInfo(Map).Aldea = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Aldea"))

    frmCargando.cargar.value = frmCargando.cargar.value + 1
Next Map


'Dim n As Integer
'For n = 500 To 717
'Debug.Print n & "- " & NpcUsado(n)
'Call LogError(n & "- " & NpcUsado(n))
'Next

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

    
End Sub


Sub LoadSini()
On Error GoTo fallo
Dim Temporal As Long
Dim Temporal1 As Long

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."


BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

ServerIp = GetVar(IniPath & "Server.ini", "INIT", "ServerIp")
Temporal = InStr(1, ServerIp, ".")
Temporal1 = (Mid(ServerIp, 1, Temporal - 1) And &H7F) * 16777216
ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
Temporal = InStr(1, ServerIp, ".")
Temporal1 = Temporal1 + Mid(ServerIp, 1, Temporal - 1) * 65536
ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
Temporal = InStr(1, ServerIp, ".")
Temporal1 = Temporal1 + Mid(ServerIp, 1, Temporal - 1) * 256
ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))

Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")
'pluto:6.9
TOPELANZAR = val(GetVar(IniPath & "Server.ini", "INIT", "AvisoLanzar"))
TOPEFLECHA = val(GetVar(IniPath & "Server.ini", "INIT", "AvisoFlecha"))

ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))
ArmaduraImperial2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial2"))
ArmaduraImperial3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial3"))
TunicaMagoImperial = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperial"))
TunicaMagoImperialEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperialEnanos"))

ArmaduraCaos1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos1"))
ArmaduraCaos2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos2"))
ArmaduraCaos3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos3"))
TunicaMagoCaos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaos"))
TunicaMagoCaosEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaosEnanos"))
'ropa legion
ArmaduraLegion1 = val(GetVar(IniPath & "Server.ini", "INIT", "Armaduralegion1"))
ArmaduraLegion2 = val(GetVar(IniPath & "Server.ini", "INIT", "Armaduralegion2"))
ArmaduraLegion3 = val(GetVar(IniPath & "Server.ini", "INIT", "Armaduralegion3"))
TunicaMagoLegion = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagolegion"))
TunicaMagoLegionEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagolegionEnanos"))
'castillos clanes
castillo1 = GetVar(IniPath & "castillos.txt", "INIT", "Castillo1")
castillo2 = GetVar(IniPath & "castillos.txt", "INIT", "Castillo2")
castillo3 = GetVar(IniPath & "castillos.txt", "INIT", "Castillo3")
castillo4 = GetVar(IniPath & "castillos.txt", "INIT", "Castillo4")
fortaleza = GetVar(IniPath & "castillos.txt", "INIT", "fortaleza")
'ciudades dueños
'DueñoNix = val(GetVar(IniPath & "ciudades.txt", "INIT", "NIX"))
'DueñoCaos = val(GetVar(IniPath & "ciudades.txt", "INIT", "CAOS"))
'DueñoUlla = val(GetVar(IniPath & "ciudades.txt", "INIT", "ULLA"))
'DueñoBander = val(GetVar(IniPath & "ciudades.txt", "INIT", "BANDER"))
'DueñoDescanso = val(GetVar(IniPath & "ciudades.txt", "INIT", "DESCANSO"))
'DueñoQuest = val(GetVar(IniPath & "ciudades.txt", "INIT", "QUEST"))
'DueñoArghal = val(GetVar(IniPath & "ciudades.txt", "INIT", "ARGHAL"))
'DueñoLaurana = val(GetVar(IniPath & "ciudades.txt", "INIT", "LAURANA"))
'DueñoLindos = val(GetVar(IniPath & "ciudades.txt", "INIT", "LINDOS"))


hora1 = GetVar(IniPath & "castillos.txt", "INIT", "hora1")
hora2 = GetVar(IniPath & "castillos.txt", "INIT", "hora2")
hora3 = GetVar(IniPath & "castillos.txt", "INIT", "hora3")
hora4 = GetVar(IniPath & "castillos.txt", "INIT", "hora4")
hora5 = GetVar(IniPath & "castillos.txt", "INIT", "hora5")
date1 = GetVar(IniPath & "castillos.txt", "INIT", "date1")
date2 = GetVar(IniPath & "castillos.txt", "INIT", "date2")
date3 = GetVar(IniPath & "castillos.txt", "INIT", "date3")
date4 = GetVar(IniPath & "castillos.txt", "INIT", "date4")
date5 = GetVar(IniPath & "castillos.txt", "INIT", "date5")

ClientsCommandsQueue = val(GetVar(IniPath & "Server.ini", "INIT", "ClientsCommandsQueue"))

If ClientsCommandsQueue <> 0 Then
        frmMain.CmdExec.Enabled = True
Else
        frmMain.CmdExec.Enabled = False
End If

'Start pos
StartPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))

'Intervalos
SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
FrmInterv.txtIntervaloSed.Text = IntervaloSed

IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

IntervaloParalisisPJ = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalisisPJ"))
'FrmInterv.txtIntervaloParalisisPJ.Text = IntervaloParalisisPJ
IntervaloMorphPJ = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMorphPJ"))
Intervaloceguera = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "Intervaloceguera"))
'FrmInterv.txtIntervaloceguera.Text = Intervaloceguera

IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
FrmInterv.txtInvocacion.Text = IntervaloInvocacion

IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&


IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
'pluto:2.8.0
IntervaloUserPuedeFlechas = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechas"))
FrmInterv.TxtFlechas.Text = IntervaloUserPuedeFlechas

IntervaloRegeneraVampiro = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloRegeneraVampiro"))
FrmInterv.txtVampire.Text = IntervaloRegeneraVampiro

'pluto:2.10
IntervaloUserPuedeTomar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeTomar"))

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

frmMain.tLluvia.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval

frmMain.CmdExec.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerExec"))
FrmInterv.txtCmdExec.Text = frmMain.CmdExec.Interval

MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
If MinutosWs < 60 Then MinutosWs = 180

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))


'Ressurect pos
ResPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
  
recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
MaxUsers = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))

'pluto:2.17
TimeEmbarazo = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TimeEmbarazo"))
TimeAborto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TimeAborto"))
ProbEmbarazo = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "ProbEmbarazo"))
'pluto:6.0A
NumeroGranPoder = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "NumeroGranPoder"))

ReDim UserList(1 To MaxUsers) As User
ReDim Cuentas(1 To MaxUsers)
Call IniciaCuentas

Nix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

ciudadcaos.Map = GetVar(DatPath & "Ciudades.dat", "CAOS", "Mapa")
ciudadcaos.X = GetVar(DatPath & "Ciudades.dat", "CAOS", "X")
ciudadcaos.Y = GetVar(DatPath & "Ciudades.dat", "CAOS", "Y")
'pluto:2.17
Pobladohumano.Map = GetVar(DatPath & "Ciudades.dat", "humano", "Mapa")
Pobladohumano.X = GetVar(DatPath & "Ciudades.dat", "humano", "X")
Pobladohumano.Y = GetVar(DatPath & "Ciudades.dat", "humano", "Y")
Pobladoorco.Map = GetVar(DatPath & "Ciudades.dat", "orco", "Mapa")
Pobladoorco.X = GetVar(DatPath & "Ciudades.dat", "orco", "X")
Pobladoorco.Y = GetVar(DatPath & "Ciudades.dat", "orco", "Y")
Pobladoenano.Map = GetVar(DatPath & "Ciudades.dat", "enano", "Mapa")
Pobladoenano.X = GetVar(DatPath & "Ciudades.dat", "enano", "X")
Pobladoenano.Y = GetVar(DatPath & "Ciudades.dat", "enano", "Y")
Pobladoelfo.Map = GetVar(DatPath & "Ciudades.dat", "elfo", "Mapa")
Pobladoelfo.X = GetVar(DatPath & "Ciudades.dat", "elfo", "X")
Pobladoelfo.Y = GetVar(DatPath & "Ciudades.dat", "elfo", "Y")
Pobladovampiro.Map = GetVar(DatPath & "Ciudades.dat", "vampiro", "Mapa")
Pobladovampiro.X = GetVar(DatPath & "Ciudades.dat", "vampiro", "X")
Pobladovampiro.Y = GetVar(DatPath & "Ciudades.dat", "vampiro", "Y")
'-------------------------------

'pluto:2.24------------------------------------
WeB = GetVar(IniPath & "Server.ini", "INIT", "WebAodraG")
DifServer = val(GetVar(IniPath & "Server.ini", "INIT", "DificultadServer"))
DifOro = val(GetVar(IniPath & "Server.ini", "INIT", "DificultadOro"))
BaseDatos = val(GetVar(IniPath & "Server.ini", "INIT", "BaseDatos"))
ServerPrimario = val(GetVar(IniPath & "Server.ini", "INIT", "ServerPrimario"))
NumeroObjEvento = val(GetVar(IniPath & "Server.ini", "EVENTOS", "NumeroObjEvento"))
CantEntregarObjEvento = val(GetVar(IniPath & "Server.ini", "EVENTOS", "CantEntregarObjEvento"))
CantObjRecompensa = val(GetVar(IniPath & "Server.ini", "EVENTOS", "CantObjRecompensa"))
ObjRecompensaEventos(1) = val(GetVar(IniPath & "Server.ini", "EVENTOS", "ObjRecompensaEventos1"))
ObjRecompensaEventos(2) = val(GetVar(IniPath & "Server.ini", "EVENTOS", "ObjRecompensaEventos2"))
ObjRecompensaEventos(3) = val(GetVar(IniPath & "Server.ini", "EVENTOS", "ObjRecompensaEventos3"))
ObjRecompensaEventos(4) = val(GetVar(IniPath & "Server.ini", "EVENTOS", "ObjRecompensaEventos4"))
'------------------------------------------------


'Call SQLConnect("localhost", "aodrag", "root", "")
Call BDDConnect
'Call BDDResetGMsos
Call BDDSetUsersOnline

Call BDDSetCastillos


Exit Sub
fallo:
Call LogError("LOADSINI" & Err.number & " D: " & Err.Description)


End Sub

Sub WriteVar(file As String, Main As String, Var As String, value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************
On Error GoTo fallo
writeprivateprofilestring Main, Var, value, file
 Exit Sub
fallo:
Call LogError("WRITEVAR" & Err.number & " D: " & Err.Description)
   
End Sub

Sub SaveUser(UserIndex As Integer, userfile As String)
On Error GoTo errhandler

'pluto:6.2------------------------------------------
'Posicion de comienzo
'Dim x As Integer
'Dim Y As Integer
'Dim Map As Integer

'Select Case UserList(UserIndex).Pos.Map


'Case MAPATORNEO 'torneos
'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
'UserList(UserIndex).Pos.Map = 296
'UserList(UserIndex).Pos.X = 71
'UserList(UserIndex).Pos.Y = 64
'Case MapaTorneo2 'torneos
'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
'UserList(UserIndex).Pos.Map = 296
'UserList(UserIndex).Pos.X = 71
'UserList(UserIndex).Pos.Y = 64
'Case 291 To 295 'torneos
'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
'UserList(UserIndex).Pos.Map = 296
'UserList(UserIndex).Pos.X = 71
'UserList(UserIndex).Pos.Y = 64
'Case 277 'fabrica lingotes
'If UserList(UserIndex).Pos.X = 36 And UserList(UserIndex).Pos.Y = 70 Then UserList(UserIndex).Pos = Nix

'Case 186 'fortaleza
'If fortaleza <> UserList(UserIndex).GuildInfo.GuildName Then
'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
'End If

'Case 166 To 169 'castillos
'UserList(UserIndex).Pos.X = 26 + RandomNumber(1, 9)
'UserList(UserIndex).Pos.Y = 85 + RandomNumber(1, 5)

'Case 191 To 192 'dragfutbol o bloqueo
'UserList(UserIndex).Pos = Nix

'End Select
'---------------------------------



If FileExist(userfile, vbNormal) Then
       If UserList(UserIndex).flags.Muerto = 1 Then UserList(UserIndex).Char.Head = val(GetVar(userfile, "INIT", "Head"))
'       Kill UserFile
End If
'pluto:6.5 quito esto lo llevo a closeuser
'If UserList(UserIndex).flags.Montura = 1 Then
'Dim obj As ObjData
'Call UsaMontura(UserIndex, obj)
'End If
Dim loopc As Integer

Call WriteVar(userfile, "FLAGS", "Muerto", val(UserList(UserIndex).flags.Muerto))
Call WriteVar(userfile, "FLAGS", "Escondido", val(UserList(UserIndex).flags.Escondido))
Call WriteVar(userfile, "FLAGS", "Hambre", val(UserList(UserIndex).flags.Hambre))
Call WriteVar(userfile, "FLAGS", "Sed", val(UserList(UserIndex).flags.Sed))
Call WriteVar(userfile, "FLAGS", "Desnudo", val(UserList(UserIndex).flags.Desnudo))
Call WriteVar(userfile, "FLAGS", "Ban", val(UserList(UserIndex).flags.ban))
Call WriteVar(userfile, "FLAGS", "Navegando", val(UserList(UserIndex).flags.Navegando))
'pluto:6.0A---------------
Call WriteVar(userfile, "FLAGS", "Minotauro", val(UserList(UserIndex).flags.Minotauro))
Call WriteVar(userfile, "FLAGS", "MinOn", val(UserList(UserIndex).flags.MinutosOnline))
'pluto:7.0
Call WriteVar(userfile, "FLAGS", "Creditos", val(UserList(UserIndex).flags.Creditos))

Call WriteVar(userfile, "FLAGS", "DragC1", val(UserList(UserIndex).flags.DragCredito1))
Call WriteVar(userfile, "FLAGS", "DragC2", val(UserList(UserIndex).flags.DragCredito2))
Call WriteVar(userfile, "FLAGS", "DragC3", val(UserList(UserIndex).flags.DragCredito3))
Call WriteVar(userfile, "FLAGS", "DragC4", val(UserList(UserIndex).flags.DragCredito4))
Call WriteVar(userfile, "FLAGS", "DragC5", val(UserList(UserIndex).flags.DragCredito5))
Call WriteVar(userfile, "FLAGS", "DragC6", val(UserList(UserIndex).flags.DragCredito6))

Call WriteVar(userfile, "FLAGS", "Elixir", val(UserList(UserIndex).flags.Elixir))
'--------------------------


'pluto:2.3
'Call WriteVar(UserFile, "FLAGS", "Montura", val(UserList(UserIndex).Flags.Montura))
'Call WriteVar(UserFile, "FLAGS", "ClaseMontura", val(UserList(UserIndex).Flags.ClaseMontura))
'pluto:2.4.1
Call WriteVar(userfile, "FLAGS", "Montura", 0)
Call WriteVar(userfile, "FLAGS", "ClaseMontura", 0)

Call WriteVar(userfile, "FLAGS", "Envenenado", val(UserList(UserIndex).flags.Envenenado))
Call WriteVar(userfile, "FLAGS", "Paralizado", val(UserList(UserIndex).flags.Paralizado))
Call WriteVar(userfile, "FLAGS", "Morph", val(UserList(UserIndex).flags.Morph))
'pluto:hoy
Call WriteVar(userfile, "QUEST", "Estado", val(UserList(UserIndex).Mision.estado))
Call WriteVar(userfile, "QUEST", "TimeC", UserList(UserIndex).Mision.TimeComienzo)
Call WriteVar(userfile, "QUEST", "Numero", val(UserList(UserIndex).Mision.numero))
'pluto:7.0---------------------------------------------------------
Call WriteVar(userfile, "QUEST", "Actual1", val(UserList(UserIndex).Mision.Actual1))
Call WriteVar(userfile, "QUEST", "Actual2", val(UserList(UserIndex).Mision.Actual2))
Call WriteVar(userfile, "QUEST", "Actual3", val(UserList(UserIndex).Mision.Actual3))
Call WriteVar(userfile, "QUEST", "Actual4", val(UserList(UserIndex).Mision.Actual4))
Call WriteVar(userfile, "QUEST", "Actual5", val(UserList(UserIndex).Mision.Actual5))
Call WriteVar(userfile, "QUEST", "Actual6", val(UserList(UserIndex).Mision.Actual6))
Call WriteVar(userfile, "QUEST", "Actual7", val(UserList(UserIndex).Mision.Actual7))
Call WriteVar(userfile, "QUEST", "Actual8", val(UserList(UserIndex).Mision.Actual8))
Call WriteVar(userfile, "QUEST", "Actual9", val(UserList(UserIndex).Mision.Actual9))
Call WriteVar(userfile, "QUEST", "Actual10", val(UserList(UserIndex).Mision.Actual10))
Call WriteVar(userfile, "QUEST", "Actual11", val(UserList(UserIndex).Mision.Actual11))
Call WriteVar(userfile, "QUEST", "Actual12", val(UserList(UserIndex).Mision.Actual12))
Call WriteVar(userfile, "QUEST", "NpcQuest", val(UserList(UserIndex).Mision.NpcQuest))
For loopc = 1 To 5
Call WriteVar(userfile, "QUEST", "EC" & loopc, val(UserList(UserIndex).Mision.NEnemigosConseguidos(loopc)))
Next
Call WriteVar(userfile, "QUEST", "PJC", val(UserList(UserIndex).Mision.PjConseguidos))
'-------------------------------------------------------------------
'Call WriteVar(userfile, "QUEST", "Level", val(UserList(UserIndex).Mision.Level))
'Call WriteVar(userfile, "QUEST", "Entrega", val(UserList(UserIndex).Mision.Entrega))
'Call WriteVar(userfile, "QUEST", "Cantidad", val(UserList(UserIndex).Mision.Cantidad))
'Call WriteVar(userfile, "QUEST", "Objeto", val(UserList(UserIndex).Mision.Objeto))
'Call WriteVar(userfile, "QUEST", "Enemigo", val(UserList(UserIndex).Mision.Enemigo))
'Call WriteVar(userfile, "QUEST", "Clase", UserList(UserIndex).Mision.clase)


Call WriteVar(userfile, "FLAGS", "Angel", val(UserList(UserIndex).flags.Angel))
Call WriteVar(userfile, "FLAGS", "Demonio", val(UserList(UserIndex).flags.Demonio))

Call WriteVar(userfile, "COUNTERS", "Pena", val(UserList(UserIndex).Counters.Pena))

Call WriteVar(userfile, "FACCIONES", "EjercitoReal", val(UserList(UserIndex).Faccion.ArmadaReal))
Call WriteVar(userfile, "FACCIONES", "EjercitoCaos", val(UserList(UserIndex).Faccion.FuerzasCaos))
Call WriteVar(userfile, "FACCIONES", "CiudMatados", val(UserList(UserIndex).Faccion.CiudadanosMatados))
Call WriteVar(userfile, "FACCIONES", "CrimMatados", val(UserList(UserIndex).Faccion.CriminalesMatados))
Call WriteVar(userfile, "FACCIONES", "rArCaos", val(UserList(UserIndex).Faccion.RecibioArmaduraCaos))
Call WriteVar(userfile, "FACCIONES", "rArReal", val(UserList(UserIndex).Faccion.RecibioArmaduraReal))
'pluto:2.3
Call WriteVar(userfile, "FACCIONES", "rArLegion", val(UserList(UserIndex).Faccion.RecibioArmaduraLegion))
Call WriteVar(userfile, "FACCIONES", "rExCaos", val(UserList(UserIndex).Faccion.RecibioExpInicialCaos))
Call WriteVar(userfile, "FACCIONES", "rExReal", val(UserList(UserIndex).Faccion.RecibioExpInicialReal))
Call WriteVar(userfile, "FACCIONES", "recCaos", val(UserList(UserIndex).Faccion.RecompensasCaos))
Call WriteVar(userfile, "FACCIONES", "recReal", val(UserList(UserIndex).Faccion.RecompensasReal))


Call WriteVar(userfile, "GUILD", "EsGuildLeader", val(UserList(UserIndex).GuildInfo.EsGuildLeader))
Call WriteVar(userfile, "GUILD", "Echadas", val(UserList(UserIndex).GuildInfo.Echadas))
Call WriteVar(userfile, "GUILD", "Solicitudes", val(UserList(UserIndex).GuildInfo.Solicitudes))
Call WriteVar(userfile, "GUILD", "SolicitudesRechazadas", val(UserList(UserIndex).GuildInfo.SolicitudesRechazadas))
Call WriteVar(userfile, "GUILD", "VecesFueGuildLeader", val(UserList(UserIndex).GuildInfo.VecesFueGuildLeader))
Call WriteVar(userfile, "GUILD", "YaVoto", val(UserList(UserIndex).GuildInfo.YaVoto))
Call WriteVar(userfile, "GUILD", "FundoClan", val(UserList(UserIndex).GuildInfo.FundoClan))
'pluto:2.4.5
Call WriteVar(userfile, "STATS", "PClan", val(UserList(UserIndex).Stats.PClan))
Call WriteVar(userfile, "STATS", "GTorneo", val(UserList(UserIndex).Stats.GTorneo))

Call WriteVar(userfile, "GUILD", "GuildName", UserList(UserIndex).GuildInfo.GuildName)
Call WriteVar(userfile, "GUILD", "ClanFundado", UserList(UserIndex).GuildInfo.ClanFundado)
Call WriteVar(userfile, "GUILD", "ClanesParticipo", str(UserList(UserIndex).GuildInfo.ClanesParticipo))
Call WriteVar(userfile, "GUILD", "GuildPts", str(UserList(UserIndex).GuildInfo.GuildPoints))

'¿Fueron modificados los atributos del usuario?
If Not UserList(UserIndex).flags.TomoPocion Then
    For loopc = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        Call WriteVar(userfile, "ATRIBUTOS", "AT" & loopc, val(UserList(UserIndex).Stats.UserAtributos(loopc)))
    Next
Else
    For loopc = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        UserList(UserIndex).Stats.UserAtributos(loopc) = UserList(UserIndex).Stats.UserAtributosBackUP(loopc)
        Call WriteVar(userfile, "ATRIBUTOS", "AT" & loopc, val(UserList(UserIndex).Stats.UserAtributos(loopc)))
    Next
End If

'pluto:7.0
Call WriteVar(userfile, "PORC", "P1", str(UserList(UserIndex).UserDañoProyetilesRaza))
Call WriteVar(userfile, "PORC", "P2", str(UserList(UserIndex).UserDañoArmasRaza))
Call WriteVar(userfile, "PORC", "P3", str(UserList(UserIndex).UserDañoMagiasRaza))
Call WriteVar(userfile, "PORC", "P4", str(UserList(UserIndex).UserDefensaMagiasRaza))
Call WriteVar(userfile, "PORC", "P5", str(UserList(UserIndex).UserEvasiónRaza))
Call WriteVar(userfile, "PORC", "P6", str(UserList(UserIndex).UserDefensaEscudos))


For loopc = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
    Call WriteVar(userfile, "SKILLS", "SK" & loopc, val(UserList(UserIndex).Stats.UserSkills(loopc)))
Next


Call WriteVar(userfile, "CONTACTO", "Email", UserList(UserIndex).Email)
'pluto:2.10
Call WriteVar(userfile, "CONTACTO", "EmailActual", Cuentas(UserIndex).mail)

Call WriteVar(userfile, "INIT", "Genero", UserList(UserIndex).Genero)
Call WriteVar(userfile, "INIT", "Raza", UserList(UserIndex).raza)
Call WriteVar(userfile, "INIT", "Hogar", UserList(UserIndex).Hogar)
Call WriteVar(userfile, "INIT", "Clase", UserList(UserIndex).clase)
Call WriteVar(userfile, "INIT", "Desc", UserList(UserIndex).Desc)
Call WriteVar(userfile, "INIT", "Heading", str(UserList(UserIndex).Char.Heading))
Call WriteVar(userfile, "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))

If UserList(UserIndex).flags.Muerto = 0 Then
    Call WriteVar(userfile, "INIT", "Body", str(UserList(UserIndex).Char.Body))
End If
If UserList(UserIndex).flags.Morph > 0 Then
    Call WriteVar(userfile, "INIT", "Body", str(UserList(UserIndex).flags.Morph))
End If
If UserList(UserIndex).flags.Angel > 0 Then
    Call WriteVar(userfile, "INIT", "Body", str(UserList(UserIndex).flags.Angel))
End If
If UserList(UserIndex).flags.Demonio > 0 Then
    Call WriteVar(userfile, "INIT", "Body", str(UserList(UserIndex).flags.Demonio))
End If
Call WriteVar(userfile, "INIT", "Arma", str(UserList(UserIndex).Char.WeaponAnim))
Call WriteVar(userfile, "INIT", "Escudo", str(UserList(UserIndex).Char.ShieldAnim))
Call WriteVar(userfile, "INIT", "Casco", str(UserList(UserIndex).Char.CascoAnim))
'[GAU]
Call WriteVar(userfile, "INIT", "Botas", str(UserList(UserIndex).Char.Botas))
'[GAU]
Call WriteVar(userfile, "INIT", "RAZAREMORT", UserList(UserIndex).Remorted)
Call WriteVar(userfile, "INIT", "BD", val(UserList(UserIndex).BD))

Call WriteVar(userfile, "INIT", "LastIP", UserList(UserIndex).ip)
'pluto:2.14
Call WriteVar(userfile, "INIT", "LastSerie", UserList(UserIndex).Serie)
Call WriteVar(userfile, "INIT", "LastMac", UserList(UserIndex).MacPluto)

'Debug.Print userfile

'pluto:6.5---------
'If UserList(UserIndex).Pos.Map = 170 Or UserList(UserIndex).Pos.Map = 34 Then
'If UserList(UserIndex).Pos.X > 16 And UserList(UserIndex).Pos.X < 31 And UserList(UserIndex).Pos.Y > 42 And UserList(UserIndex).Pos.Y < 48 Then
'UserList(UserIndex).Pos.X = 36
'UserList(UserIndex).Pos.Y = 36
'End If
'End If
'------------------
Call WriteVar(userfile, "INIT", "Position", UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)

' pluto:2.15 -------------------
Call WriteVar(userfile, "INIT", "Esposa", UserList(UserIndex).Esposa)
Call WriteVar(userfile, "INIT", "Nhijos", val(UserList(UserIndex).Nhijos))
Dim X As Byte
For X = 1 To 5
Call WriteVar(userfile, "INIT", "Hijo" & X, UserList(UserIndex).Hijo(X))
Next
Call WriteVar(userfile, "INIT", "Amor", val(UserList(UserIndex).Amor))
Call WriteVar(userfile, "INIT", "Embarazada", val(UserList(UserIndex).Embarazada))
Call WriteVar(userfile, "INIT", "Bebe", val(UserList(UserIndex).Bebe))
Call WriteVar(userfile, "INIT", "NombreDelBebe", UserList(UserIndex).NombreDelBebe)
Call WriteVar(userfile, "INIT", "Padre", UserList(UserIndex).Padre)
Call WriteVar(userfile, "INIT", "Madre", UserList(UserIndex).Madre)
'-----------------------------------

'PLUTO:2-3-04
Call WriteVar(userfile, "STATS", "PUNTOS", str(UserList(UserIndex).Stats.Puntos))

Call WriteVar(userfile, "STATS", "GLD", str(UserList(UserIndex).Stats.GLD))
Call WriteVar(userfile, "STATS", "REMORT", str(UserList(UserIndex).Remort))
Call WriteVar(userfile, "STATS", "BANCO", str(UserList(UserIndex).Stats.Banco))

Call WriteVar(userfile, "STATS", "MET", str(UserList(UserIndex).Stats.MET))
Call WriteVar(userfile, "STATS", "MaxHP", str(UserList(UserIndex).Stats.MaxHP))
Call WriteVar(userfile, "STATS", "MinHP", str(UserList(UserIndex).Stats.MinHP))

Call WriteVar(userfile, "STATS", "FIT", str(UserList(UserIndex).Stats.FIT))
Call WriteVar(userfile, "STATS", "MaxSTA", str(UserList(UserIndex).Stats.MaxSta))
Call WriteVar(userfile, "STATS", "MinSTA", str(UserList(UserIndex).Stats.MinSta))

Call WriteVar(userfile, "STATS", "MaxMAN", str(UserList(UserIndex).Stats.MaxMAN))
Call WriteVar(userfile, "STATS", "MinMAN", str(UserList(UserIndex).Stats.MinMAN))

Call WriteVar(userfile, "STATS", "MaxHIT", str(UserList(UserIndex).Stats.MaxHIT))
Call WriteVar(userfile, "STATS", "MinHIT", str(UserList(UserIndex).Stats.MinHIT))

Call WriteVar(userfile, "STATS", "MaxAGU", str(UserList(UserIndex).Stats.MaxAGU))
Call WriteVar(userfile, "STATS", "MinAGU", str(UserList(UserIndex).Stats.MinAGU))

Call WriteVar(userfile, "STATS", "MaxHAM", str(UserList(UserIndex).Stats.MaxHam))
Call WriteVar(userfile, "STATS", "MinHAM", str(UserList(UserIndex).Stats.MinHam))

Call WriteVar(userfile, "STATS", "SkillPtsLibres", str(UserList(UserIndex).Stats.SkillPts))
  
Call WriteVar(userfile, "STATS", "EXP", str(UserList(UserIndex).Stats.exp))
Call WriteVar(userfile, "STATS", "ELV", str(UserList(UserIndex).Stats.ELV))
Call WriteVar(userfile, "STATS", "ELU", str(UserList(UserIndex).Stats.Elu))
'pluto:6.0A
Call WriteVar(userfile, "STATS", "LIBROSUSADOS", str(UserList(UserIndex).Stats.LibrosUsados))
Call WriteVar(userfile, "STATS", "FAMA", str(UserList(UserIndex).Stats.Fama))

Call WriteVar(userfile, "MUERTES", "UserMuertes", val(UserList(UserIndex).Stats.UsuariosMatados))
Call WriteVar(userfile, "MUERTES", "CrimMuertes", val(UserList(UserIndex).Stats.CriminalesMatados))
Call WriteVar(userfile, "MUERTES", "NpcsMuertes", val(UserList(UserIndex).Stats.NPCsMuertos))
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************

'pluto:7.0 quito esto que pasa a sistema cuentas
'Call WriteVar(userfile, "BancoInventory", "CantidadItems", val(UserList(UserIndex).BancoInvent.NroItems))
'Dim loopd As Integer
'pluto:7.0
'For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
 '   Call WriteVar(userfile, "BancoInventory", "Obj" & loopd, UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).Amount)
'Next loopd
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Call WriteVar(userfile, "Inventory", "CantidadItems", val(UserList(UserIndex).Invent.NroItems))

For loopc = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(userfile, "Inventory", "Obj" & loopc, UserList(UserIndex).Invent.Object(loopc).ObjIndex & "-" & UserList(UserIndex).Invent.Object(loopc).Amount & "-" & UserList(UserIndex).Invent.Object(loopc).Equipped)
Next

Call WriteVar(userfile, "Inventory", "WeaponEqpSlot", str(UserList(UserIndex).Invent.WeaponEqpSlot))
Call WriteVar(userfile, "Inventory", "ArmourEqpSlot", str(UserList(UserIndex).Invent.ArmourEqpSlot))
Call WriteVar(userfile, "Inventory", "CascoEqpSlot", str(UserList(UserIndex).Invent.CascoEqpSlot))
Call WriteVar(userfile, "Inventory", "EscudoEqpSlot", str(UserList(UserIndex).Invent.EscudoEqpSlot))
Call WriteVar(userfile, "Inventory", "BarcoSlot", str(UserList(UserIndex).Invent.BarcoSlot))
Call WriteVar(userfile, "Inventory", "MunicionSlot", str(UserList(UserIndex).Invent.MunicionEqpSlot))
'pluto:2.4.1
Call WriteVar(userfile, "Inventory", "AnilloEqpSlot", str(UserList(UserIndex).Invent.AnilloEqpSlot))

'[GAU]
Call WriteVar(userfile, "Inventory", "BotaEqpSlot", str(UserList(UserIndex).Invent.BotaEqpSlot))
'[GAU]


'Reputacion
Call WriteVar(userfile, "REP", "Asesino", val(UserList(UserIndex).Reputacion.AsesinoRep))
Call WriteVar(userfile, "REP", "Bandido", val(UserList(UserIndex).Reputacion.BandidoRep))
Call WriteVar(userfile, "REP", "Burguesia", val(UserList(UserIndex).Reputacion.BurguesRep))
Call WriteVar(userfile, "REP", "Ladrones", val(UserList(UserIndex).Reputacion.LadronesRep))
Call WriteVar(userfile, "REP", "Nobles", val(UserList(UserIndex).Reputacion.NobleRep))
Call WriteVar(userfile, "REP", "Plebe", val(UserList(UserIndex).Reputacion.PlebeRep))

Dim l As Long
l = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
l = l / 6
Call WriteVar(userfile, "REP", "Promedio", val(l))

Dim cad As String

For loopc = 1 To MAXUSERHECHIZOS
    cad = UserList(UserIndex).Stats.UserHechizos(loopc)
    Call WriteVar(userfile, "HECHIZOS", "H" & loopc, cad)
Next


For loopc = 1 To MAXMASCOTAS
    ' Mascota valida?
    If UserList(UserIndex).MascotasIndex(loopc) > 0 Then
        ' Nos aseguramos que la criatura no fue invocada
        If Npclist(UserList(UserIndex).MascotasIndex(loopc)).Contadores.TiempoExistencia = 0 Then
            cad = UserList(UserIndex).MascotasType(loopc)
        Else 'Si fue invocada no la guardamos
            cad = "0"
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
        End If
        Call WriteVar(userfile, "MASCOTAS", "MAS" & loopc, 0)
    End If

Next

Call WriteVar(userfile, "MASCOTAS", "NroMascotas", 0)

'pluto:6.0A -guardamos mascotas
Call WriteVar(userfile, "MONTURAS", "NroMonturas", val(UserList(UserIndex).Nmonturas))

loopc = 0
Dim n As Byte
For n = 1 To 12

loopc = UserList(UserIndex).Montura.index(n)
If loopc > 0 Then
Call WriteVar(userfile, "MONTURA" & loopc, "TIPO", val(UserList(UserIndex).Montura.Tipo(n)))

Call WriteVar(userfile, "MONTURA" & loopc, "NIVEL", val(UserList(UserIndex).Montura.Nivel(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "EXP", val(UserList(UserIndex).Montura.exp(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "ELU", val(UserList(UserIndex).Montura.Elu(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "VIDA", val(UserList(UserIndex).Montura.Vida(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "GOLPE", val(UserList(UserIndex).Montura.Golpe(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "NOMBRE", UserList(UserIndex).Montura.Nombre(n))

Call WriteVar(userfile, "MONTURA" & loopc, "ATCUERPO", val(UserList(UserIndex).Montura.AtCuerpo(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "DEFCUERPO", val(UserList(UserIndex).Montura.Defcuerpo(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "ATFLECHAS", val(UserList(UserIndex).Montura.AtFlechas(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "DEFFLECHAS", val(UserList(UserIndex).Montura.DefFlechas(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "ATMAGICO", val(UserList(UserIndex).Montura.AtMagico(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "DEFMAGICO", val(UserList(UserIndex).Montura.DefMagico(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "EVASION", val(UserList(UserIndex).Montura.Evasion(n)))
Call WriteVar(userfile, "MONTURA" & loopc, "LIBRES", val(UserList(UserIndex).Montura.Libres(n)))
End If
Next

'Delzak sistema premios
For n = 1 To 34
Call WriteVar(userfile, "PREMIOS", "L" & n, val(UserList(UserIndex).Stats.PremioNPC(n)))
Next


Exit Sub
errhandler:
Call LogError("Error en SaveUser")

End Sub

Function Criminal(ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
'Dim a As Integer
'If UserList(UserIndex).Reputacion.Promedio < 0 Then a = 1 Else a = 0
Dim l As Long
l = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
l = l / 6
Criminal = (l < 0)
UserList(UserIndex).Reputacion.Promedio = l
'If a = 0 And Criminal = True Then UserCrimi = UserCrimi + 1: UserCiu = UserCiu - 1
'If a = 1 And Criminal = False Then UserCiu = UserCiu + 1: UserCrimi = UserCrimi - 1

 

Exit Function
fallo:
Call LogError("CRIMINAL " & Err.number & " D: " & Err.Description)

End Function




Sub BackUPnPc(NpcIndex As Integer)
On Error GoTo fallo
'Call LogTarea("Sub BackUPnPc NpcIndex:" & NpcIndex)

Dim NpcNumero As Integer
Dim npcfile As String
Dim loopc As Integer


NpcNumero = Npclist(NpcIndex).numero

If NpcNumero > 499 Then
    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "bkNPCs.dat"
End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).Name)

Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).Desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.Heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Inflacion", val(Npclist(NpcIndex).Inflacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))
'pluto:6.0A
Call WriteVar(npcfile, "NPC" & NpcNumero, "Arquero", val(Npclist(NpcIndex).Arquero))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Anima", val(Npclist(NpcIndex).Anima))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Raid", val(Npclist(NpcIndex).Raid))
'pluto:7.0
Call WriteVar(npcfile, "NPC" & NpcNumero, "LogroTipo", val(Npclist(NpcIndex).LogroTipo))

'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.Def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados))




'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For loopc = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & loopc, Npclist(NpcIndex).Invent.Object(loopc).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(loopc).Amount)
   Next
End If

Exit Sub
fallo:
Call LogError("BACKUPNPC" & Err.number & " D: " & Err.Description)

End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)
On Error GoTo fallo
'Call LogTarea("Sub CargarNpcBackUp NpcIndex:" & NpcIndex & " NpcNumber:" & NpcNumber)

'Status
If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"


Dim npcfile As String

If NpcNumber > 499 Then
        npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
        npcfile = DatPath & "bkNPCs.dat"
End If

Npclist(NpcIndex).numero = NpcNumber
'pluto:2.17
Npclist(NpcIndex).Anima = val(GetVar(npcfile, "NPC" & NpcNumber, "Anima"))
Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
'pluto:6.0A
Npclist(NpcIndex).Arquero = val(GetVar(npcfile, "NPC" & NpcNumber, "Arquero"))
Npclist(NpcIndex).Raid = val(GetVar(npcfile, "NPC" & NpcNumber, "Raid"))
'pluto:7.0
Npclist(NpcIndex).LogroTipo = val(GetVar(npcfile, "NPC" & NpcNumber, "LogroTipo"))

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))


Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

'@Nati: NPCS vida a 1
'Npclist(NpcIndex).Stats.MaxHP = 1
'Npclist(NpcIndex).Stats.MinHP = 1
Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NpcNumber, "ImpactRate"))
'Npclist(NpcIndex).Premio = val(GetVar(npcfile, "NPC" & NpcNumber, "Premio")) 'Delzak sistema premios


Dim loopc As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For loopc = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & loopc)
        Npclist(NpcIndex).Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))
       
    Next loopc
Else
    For loopc = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(loopc).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(loopc).Amount = 0
    Next loopc
End If

Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Inflacion"))


Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False
Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "PosOrig"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))
Exit Sub
fallo:
Call LogError("CARGARNPCBACKUP" & Err.number & " D: " & Err.Description)

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal moTivo As String)
On Error GoTo fallo
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "Reason", moTivo)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "Fecha", Date)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile
Exit Sub
fallo:
Call LogError("LOGBAN" & Err.number & " D: " & Err.Description)

End Sub

Private Sub BuscaPosicionValida(UserIndex As Integer)

'Delzak (28-8-10)

Dim Leer As New clsLeerInis
Dim Mapa As Integer
Dim X As Integer
Dim Y As Integer
Dim MapaOK As Integer
Dim XOK As Integer
Dim YOK As Integer
Dim dn As Integer
Dim M As Integer
Dim User As Integer
Dim iNDiCe As Integer
Dim QueSumo As Boolean '0 para x, 1 para y
Dim PosicionValida As Boolean
Dim ControlBordes As Boolean

Mapa = UserList(UserIndex).Pos.Map
X = UserList(UserIndex).Pos.X
Y = UserList(UserIndex).Pos.Y
MapaOK = Mapa
XOK = X
YOK = Y
QueSumo = False
iNDiCe = 1
ControlBordes = True


'Busco un hueco donde no haya nadie y que no este bloqueado (OPTIMIZADO 14-9-10)

For dn = 1 To 6400 '80x80
    
PosicionValida = True
    'Compruebo que no haya nadie en la posicion que quiero logear
    For User = 1 To LastUser
        If UserList(User).Pos.Map = MapaOK And UserList(User).Pos.X = XOK And UserList(User).Pos.Y = YOK Then PosicionValida = False
    Next
    
    'Compruebo que no este bloqueado
    If PosicionValida = True Then
    
        If MapData(MapaOK, XOK, YOK).Blocked = 1 Then PosicionValida = False
    
    End If

    'Si la posicion es valida, salgo del bucle
    If PosicionValida = True And ControlBordes = True Then Exit For

    'Si no es valida, busco una trazando un espiral
   
    If QueSumo = False Then
        
        XOK = XOK + iNDiCe
        
    Else
        
        YOK = YOK + iNDiCe
        
        iNDiCe = iNDiCe * (-1)
        If iNDiCe < 0 Then iNDiCe = iNDiCe - 1
        If iNDiCe > 0 Then iNDiCe = iNDiCe + 1
    End If
    
    If QueSumo = True Then QueSumo = False
    If QueSumo = False Then QueSumo = True
    
    'Controlo que no me salga del borde
    If XOK < 4 Or XOK > 85 Or YOK < 4 Or YOK > 85 Then ControlBordes = False Else ControlBordes = True
    

    'Si termina el bucle y no he encontrado alternativa, que le den por culo
    If dn = 6400 Then
        MapaOK = Mapa
        XOK = X
        YOK = Y
    End If
    
Next

'Bloqueo la posicion donde voy a aparecer para que no me de por culo nadie
MapData(MapaOK, XOK, YOK).Blocked = 1

'Cargo mi posicion
UserList(UserIndex).Pos.Map = MapaOK
UserList(UserIndex).Pos.X = XOK
UserList(UserIndex).Pos.Y = YOK

'Desbloqueo la posicion
MapData(MapaOK, XOK, YOK).Blocked = 0

End Sub
