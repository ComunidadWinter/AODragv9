Attribute VB_Name = "ES"
Option Explicit

'*************************************
'Lorwik - Nuevo formato de mapas .CSM
'*************************************

Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    trigger As Integer
End Type

Private Type tDatosLuces
    X As Integer
    Y As Integer
    light_value(3) As Long
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NpcIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    music_number As String
    zone As String
    terrain As String
    Ambient As String
    lvlMinimo As String
    SePuedeDomar As Boolean
    ResuSinEfecto As Boolean
    MagiaSinEfecto As Boolean
    InviSinEfecto As Boolean
    NoEncriptarMP As Boolean
    version As Long
End Type


Public Sub CargarSpawnList()
    Dim n As Integer, LoopC As Integer
    n = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(n) As tCriaturasEntrenador
    For LoopC = 1 To n
        SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
End Sub

Function EsAdmin(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Admines"))

For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Admines", "Admin" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsAdmin = True
        Exit Function
    End If
Next WizNum
EsAdmin = False

End Function

Function EsDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
End Function

Function EsSemiDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False

End Function

Function EsConsejero(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
End Function

Function EsRolesMaster(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsRolesMaster = True
        Exit Function
    End If
Next WizNum
EsRolesMaster = False
End Function


Public Function TxtDimension(ByVal Name As String) As Long
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
End Function

Public Sub CargarForbidenWords()

ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim n As Integer, i As Integer
n = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #n

For i = 1 To UBound(ForbidenNames)
    Line Input #n, ForbidenNames(i)
Next i

Close n

End Sub

Public Sub CargarHechizos()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo errhandler

    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando Hechizos."
    
    Dim Hechizo As Integer
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(DatPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.value = 0

    'Llena la lista
    For Hechizo = 1 To NumeroHechizos
        With Hechizos(Hechizo)
            
            .nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            .GrhIndex = val(Leer.GetValue("Hechizo" & Hechizo, "GrhIndex"))
            If .GrhIndex = 0 Then .GrhIndex = 609 ' Imagen de Hechizo "generico"
            .desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
          
            .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            .targetMSG = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
            .loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
        '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHP = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHP = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
            .SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
            .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
            .SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
            .SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
            .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            
            
            .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
            '01/02/2016 - Lorwik> Ahora podemos añadir mas clases por hechizos.
            Dim i As Byte
            Dim NumCP As String
            
            NumCP = Leer.GetValue("Hechizo" & Hechizo, "NUMCP")
            
            If Not NumCP = "" Then
                For i = 1 To NumCP
                    .ExclusivoClase(i) = Leer.GetValue("Hechizo" & Hechizo, "CP" & i)
                Next i
            End If
            
            .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
            
        '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
        '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            
            'Barrin 30/9/03
            .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
            frmCargando.cargar.value = frmCargando.cargar.value + 1
       End With
    Next Hechizo

    Set Leer = Nothing
    Exit Sub

errhandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
End Sub

Sub LoadMotd()
Dim i As Integer

MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))

ReDim MOTD(1 To MaxLines)
For i = 1 To MaxLines
    MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
    MOTD(i).Formato = vbNullString
Next i

End Sub

Public Sub DoBackUp()
  HaciendoBackup = True
  Dim i As Integer


  Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())


  Call LimpiarMundo
  Call WorldSave
  'Call modGuilds.v_RutinaElecciones
  Call ResetCentinelaInfo     'Reseteamos al centinela
  
  
  Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
  
  'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)
  
  HaciendoBackup = False
  
  'Log
  On Error Resume Next
  Dim nfile As Integer
  nfile = FreeFile ' obtenemos un canal
  Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
  Print #nfile, Date & " " & time
  Close #nfile
End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByVal MAPFILE As String)
' Lorwik> PENDIENTE DE PROGRAMAR
End Sub
Sub LoadArmasHerreria()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To n) As Integer

For lc = 1 To n
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc

End Sub

Sub LoadArmadurasHerreria()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To n) As Integer

For lc = 1 To n
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadBalance()
    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        ModClase(i).Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
        ModClase(i).AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
        ModClase(i).AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
        ModClase(i).DañoArmas = val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
        ModClase(i).DañoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
        ModClase(i).DañoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
        ModClase(i).Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
        
        ModClase(i).Fuerza = val(GetVar(DatPath & "Balance.dat", "MODCLASES", ListaClases(i) + "Fuerza"))
        ModClase(i).Agilidad = val(GetVar(DatPath & "Balance.dat", "MODCLASES", ListaClases(i) + "Agilidad"))
        ModClase(i).Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODCLASES", ListaClases(i) + "Inteligencia"))
        ModClase(i).Energia = val(GetVar(DatPath & "Balance.dat", "MODCLASES", ListaClases(i) + "Energia"))
        ModClase(i).Constitucion = val(GetVar(DatPath & "Balance.dat", "MODCLASES", ListaClases(i) + "Constitucion"))
    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        ModRaza(i).Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
        ModRaza(i).Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
        ModRaza(i).Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
        ModRaza(i).Energia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Energia"))
        ModRaza(i).Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribución de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i
    For i = 1 To 4
        DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Party
    'ExponenteNivelParty = val(GetVar(DatPath & "Balance.dat", "PARTY", "ExponenteNivelParty"))
End Sub

Sub LoadObjCarpintero()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To n) As Integer

For lc = 1 To n
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc

End Sub



Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.Text = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
    Dim Object As Integer
    Dim Leer As New clsIniReader
    
    Dim i As Integer
    Dim n As Integer
    Dim S As String
    
    Call Leer.Initialize(DatPath & "Obj.dat")
    
    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.value = 0
    
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
    'Llena la lista
    For Object = 1 To NumObjDatas
        With ObjData(Object)
            .Name = Leer.GetValue("OBJ" & Object, "Name")
            
            'Pablo (ToxicWaste) Log de Objetos.
            .Log = val(Leer.GetValue("OBJ" & Object, "Log"))
            .NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
            '07/09/07
            
            .GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
            If .GrhIndex = 0 Then
                .GrhIndex = .GrhIndex
            End If
            
            .OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
            
            .Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
            
            .CantMateriales = val(Leer.GetValue("OBJ" & Object, "CantMaterial"))
            
            'NUEVO SISTEMA DE CRAFTEO
            If .CantMateriales > 0 Then
                For i = 1 To .CantMateriales
                    .Material(i).Material = val(Leer.GetValue("OBJ" & Object, ReadField(1, "Material" & i, Asc("-"))))
                    .Material(i).CantMaterial = val(Leer.GetValue("OBJ" & Object, ReadField(2, "Material" & i, Asc("-"))))
                Next i
            End If
            
            Select Case .OBJType
                Case eOBJType.otArmadura
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                
                Case eOBJType.otESCUDO
                    .ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otCASCO
                    .CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otWeapon
                    .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                    .Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                    .Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
                    
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otInstrumentos
                    .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                    .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                    .Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                    'Pablo (ToxicWaste)
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otMinerales
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                Case otPociones
                    .TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                    .MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                    .MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                    .DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                
                Case eOBJType.otBarcos
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                
                Case eOBJType.otMonturas
                    .IndiceMontura = val(Leer.GetValue("OBJ" & Object, "IndiceMontura"))
                    
                Case eOBJType.otFlechas
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                    
                Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .EfectoMagico = val(Leer.GetValue("OBJ" & Object, "EfectoMagico"))
                    
                Case eOBJType.otPasajes
                    .DesdeMap = val(Leer.GetValue("OBJ" & Object, "DesdeMap"))
                    .HastaMap = val(Leer.GetValue("OBJ" & Object, "HastaMap"))
                    .HastaX = val(Leer.GetValue("OBJ" & Object, "HastaX"))
                    .HastaY = val(Leer.GetValue("OBJ" & Object, "HastaY"))
                    .CantidadSkill = val(Leer.GetValue("OBJ" & Object, "CantidadSkill"))
                    
                Case eOBJType.otCabezaMontura
                    .IndiceMontura = val(Leer.GetValue("OBJ" & Object, "IndiceMontura"))
                    .Cabeza = val(Leer.GetValue("OBJ" & Object, "Cabeza"))
                    .skdomar = val(Leer.GetValue("OBJ" & Object, "SKDomar"))
                    
                Case eOBJType.otManual
                    .IndiceSkill = val(Leer.GetValue("OBJ" & Object, "IndiceSkill"))
                    .CuantosSkill = val(Leer.GetValue("OBJ" & Object, "CuantosSkill"))
                    .SkNecesarios = val(Leer.GetValue("OBJ" & Object, "SkNecesarios"))
                    
                Case eOBJType.otCofre
                    '01-03-2016 Lorwik: Sistema de cofre con items aleatorios.
                    'Maximo 255 obj por cofre.
                    .CantItems = val(Leer.GetValue("OBJ" & Object, "CantItems"))
                    .CofreCerrado = val(Leer.GetValue("OBJ" & Object, "CofreCerrado"))
                    
                    'Un cofre podria estar vacio, por eso comprobamos.
                    If .CantItems > 0 Then
                        For i = 1 To .CantItems
                            .ItemCofre(i).Obj = val(Leer.GetValue("OBJ" & Object, ReadField(1, "ItemCofre" & i, Asc("-"))))
                            .ItemCofre(i).cant = val(Leer.GetValue("OBJ" & Object, ReadField(2, "ItemCofre" & i, Asc("-"))))
                            .ItemCofre(i).Prob = val(Leer.GetValue("OBJ" & Object, ReadField(3, "ItemCofre" & i, Asc("-")))) 'Desde el 0% al 100%
                        Next i
                    End If
            End Select
            
            .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
            .HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
            .LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
            .MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
            .MaxHP = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
            .MinHP = val(Leer.GetValue("OBJ" & Object, "MinHP"))
            
            .Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
            .Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
            .MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
            .MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
            .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
            .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
            .def = (.MinDef + .MaxDef) / 2
            
            .RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
            .RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
            .RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
            .RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
            .RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
            
            .valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
                    
            .NoLimpiar = val(Leer.GetValue("OBJ" & Object, "NoLimpiar"))
            
            .Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
            
            .Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))
            If .Cerrada = 1 Then
                .Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
                .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            End If
            
            'Puertas y llaves
            .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            
            .texto = Leer.GetValue("OBJ" & Object, "Texto")
            .GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
            .Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
            .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
            For i = 1 To NUMCLASES
                S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
                n = 1
                Do While LenB(S) > 0 And UCase$(ListaClases(n)) <> S
                    n = n + 1
                Loop
                .ClaseProhibida(i) = IIf(LenB(S) > 0, n, 0)
            Next i
            
            .DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
            .DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
            .DañoMagico = val(Leer.GetValue("OBJ" & Object, "DañoMagico"))
            
            .SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            
            'Bebidas
            .MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
            
            .NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            
            .Nosetira = val(Leer.GetValue("OBJ" & Object, "Nosetira"))
            
            .NoComerciable = val(Leer.GetValue("OBJ" & Object, "NoComerciable"))
            
            '*************************************************
            '21/03/2016 Lorwik: Aportan aumentos en los stats
            
            .SumaVida = val(Leer.GetValue("OBJ" & Object, "SumaVida"))
            .SumaMana = val(Leer.GetValue("OBJ" & Object, "SumaMana"))
            .SumaFuerza = val(Leer.GetValue("OBJ" & Object, "SumaFuerza"))
            .SumaAgilidad = val(Leer.GetValue("OBJ" & Object, "SumaAgilidad"))
            '*************************************************
            
            '16/12/2018 Lorwik: Aumenta velocidad
            .Speed = val(Leer.GetValue("OBJ" & Object, "Speed"))
            
            frmCargando.cargar.value = frmCargando.cargar.value + 1
        End With
    Next Object
    Set Leer = Nothing
    
    Exit Sub

errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description


End Sub

Sub LoadUserMonturas(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
    Dim i As Byte
    
    With UserList(UserIndex)
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT count(id_personaje) as Total FROM montura WHERE id_personaje = '" & .id & "'")
    
        'Primero cargamos el numero de monturas
        .Stats.NUMMONTURAS = CByte(RS!Total)
    
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT id,tipo,nombre,level,skills,elu,exp,ataque,defensa,atmagia,defmagia,evasion,speed FROM montura WHERE id_personaje = '" & .id & "'")
    
        '¿Tiene alguna montura?
        If .Stats.NUMMONTURAS = 0 Then Exit Sub
        
        i = 1
        
        'Ya podemos cargar todas las monturas
        While Not RS.EOF
            .flags.Montura(i).id = CInt(RS!id)
            .flags.Montura(i).tipo = CByte(RS!tipo)
            Debug.Print .flags.Montura(i).tipo
            .flags.Montura(i).nombre = CStr(RS!nombre)
            .flags.Montura(i).MonturaLevel = CByte(RS!level)
            .flags.Montura(i).Skills = CByte(RS!Skills)
            .flags.Montura(i).ELU = CInt(RS!ELU)
            .flags.Montura(i).Exp = CInt(RS!Exp)
            .flags.Montura(i).Ataque = CByte(RS!Ataque)
            .flags.Montura(i).Defensa = CByte(RS!Defensa)
            .flags.Montura(i).AtMagia = CByte(RS!AtMagia)
            .flags.Montura(i).DefMagia = CByte(RS!DefMagia)
            .flags.Montura(i).Evasion = CByte(RS!Evasion)
            .flags.Montura(i).Speed = CByte(RS!Speed)
            i = i + 1
            RS.MoveNext
        Wend

    End With
End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef Name As String)

Dim LoopC As Long

    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT * FROM personaje WHERE nombre = '" & Name & "'")
    UserList(UserIndex).id = RS!id
    
    UserList(UserIndex).Stats.GLD = CLng(RS!GLD)
    UserList(UserIndex).Stats.Banco = CLng(RS!Banco)
    
    UserList(UserIndex).Stats.MaxHP = CInt(RS!MaxHP)
    UserList(UserIndex).Stats.MinHP = CInt(RS!MinHP)
    
    UserList(UserIndex).Stats.MinSta = CInt(RS!MinSta)
    UserList(UserIndex).Stats.MaxSta = CInt(RS!MaxSta)
    
    UserList(UserIndex).Stats.MaxMAN = CInt(RS!MaxMAN)
    UserList(UserIndex).Stats.MinMAN = CInt(RS!MinMAN)
    
    UserList(UserIndex).Stats.MaxHIT = CInt(RS!MaxHIT)
    UserList(UserIndex).Stats.MinHIT = CInt(RS!MinHIT)
    
    UserList(UserIndex).Stats.MaxAGU = CByte(RS!MaxAGU)
    UserList(UserIndex).Stats.MinAGU = CByte(RS!MinAGU)
    
    UserList(UserIndex).Stats.MaxHam = CByte(RS!MaxHam)
    UserList(UserIndex).Stats.MinHam = CByte(RS!MinHam)
    
    UserList(UserIndex).Stats.Exp = CDbl(RS!Exp)
    UserList(UserIndex).Stats.ELU = CLng(RS!ELU)
    UserList(UserIndex).Stats.ELO = CDbl(RS!ELO)
    UserList(UserIndex).Stats.ELV = CByte(RS!ELV)
    UserList(UserIndex).Stats.DragCredits = CInt(RS!DragCredits)
    
    UserList(UserIndex).Stats.UsuariosMatados = CLng(RS!usermuertes)
    UserList(UserIndex).Stats.NPCsMuertos = CInt(RS!npcsmuertes)
    
    If CByte(RS!pertenece) Then _
        UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoyalCouncil
    
    If CByte(RS!pertenececaos) Then _
        UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.ChaosCouncil

    RS.Close
    
    '05/11/2015 Irongete: Cargar los atributos del personaje
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT atributo,valor FROM rel_personaje_atributo WHERE id_personaje = '" & UserList(UserIndex).id & "'")
    While Not RS.EOF
        UserList(UserIndex).Stats.UserAtributos(RS!atributo) = CInt(RS!valor)
        UserList(UserIndex).Stats.UserAtributosBackUP(RS!atributo) = UserList(UserIndex).Stats.UserAtributos(RS!atributo)
        RS.MoveNext
    Wend
    RS.Close
    
    '05/11/2015 Irongete: Cargar los skills del personaje
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT skill,valor FROM rel_personaje_skill WHERE id_personaje = '" & UserList(UserIndex).id & "'")
    While Not RS.EOF
        UserList(UserIndex).Stats.UserSkills(RS!Skill) = CInt(RS!valor)
        RS.MoveNext
    Wend
    RS.Close
    
    '05/11/2015 Irongete: Cargar los hechizos
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT slot,hechizo FROM rel_personaje_hechizo WHERE id_personaje = '" & UserList(UserIndex).id & "'")
    While Not RS.EOF
        UserList(UserIndex).Stats.UserHechizos(RS!Slot) = CInt(RS!Hechizo)
        RS.MoveNext
    Wend
    RS.Close

    '17/12/2018 Irongete: Cargar las habilidades
    'Set RS = New ADODB.Recordset
    'Set RS = SQL.Execute("SELECT id_habilidad, slot FROM rel_personaje_habilidad WHERE id_personaje = '" & UserList(UserIndex).id & "'")
    'While Not RS.EOF
    '  UserList(UserIndex).Stats.UserHechizos(RS!Slot) = CInt(RS!id_habilidad)
    '  RS.MoveNext
    'Wend
    'RS.Close

End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef Name As String)

    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT asesino,bandido,burguesia,ladrones,nobles,plebe,promedio FROM personaje WHERE nombre = '" & Name & "'")

    UserList(UserIndex).Reputacion.AsesinoRep = val(RS!asesino)
    UserList(UserIndex).Reputacion.BandidoRep = val(RS!bandido)
    UserList(UserIndex).Reputacion.BurguesRep = val(RS!burguesia)
    UserList(UserIndex).Reputacion.LadronesRep = val(RS!ladrones)
    UserList(UserIndex).Reputacion.NobleRep = val(RS!nobles)
    UserList(UserIndex).Reputacion.PlebeRep = val(RS!plebe)
    UserList(UserIndex).Reputacion.Promedio = val(RS!Promedio)
    
    RS.Close


End Sub


Sub LoadUserInit(ByVal UserIndex As Integer, ByRef Name As String)
'*************************************************
'Author: Unknown
'Last modified: 19/11/2006
'Loads the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
'*************************************************
Dim LoopC As Long
Dim ln As String
With UserList(UserIndex)

    Dim PartyId As Integer
    
    '05/11/2015 Irongete: Cargar datos del personaje
    
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT j1.id_party, personaje.* FROM personaje LEFT JOIN rel_party_personaje j1 ON j1.id_personaje = personaje.id WHERE nombre = '" & Name & "'")
    
    PartyId = 0
    If Not IsNull(RS!id_party) Then
        PartyId = CInt(RS!id_party)
    End If
    
    '18/11/2015 Irongete: Compruebo si está en party y le asigno la ID
    If PartyId > 0 Then
        .PartyId = PartyId
    End If
    
    .id = RS!id
    .CuentaId = RS!id_Cuenta
    
    .genero = RS!genero
    .clase = RS!clase
    .raza = RS!raza
    .Char.heading = RS!heading
    
    .OrigChar.Head = CInt(RS!Head)
    .OrigChar.body = CInt(RS!body)
    .OrigChar.WeaponAnim = CInt(RS!Arma)
    .OrigChar.ShieldAnim = CInt(RS!Escudo)
    .OrigChar.CascoAnim = CInt(RS!casco)
    
    #If ConUpTime Then
        .UpTime = CLng(RS!UpTime)
    #End If
    
    .OrigChar.heading = eHeading.SOUTH
    
    .desc = RS!descripcion
    
    .GuildIndex = CInt(RS!id_clan)
    
    .Pos.Map = CInt(ReadField(1, RS!Position, 45))
    .Pos.X = CInt(ReadField(2, RS!Position, 45))
    .Pos.Y = CInt(ReadField(3, RS!Position, 45))
    
    If .Pos.Map = 43 Then
        If .GuildIndex <> Castillos(5).Dueño Then
            .Pos.Map = 1
            .Pos.X = 50
            .Pos.Y = 50
        End If
    End If
    
    


    
    .Faccion.ArmadaReal = CByte(RS!ejercitoreal)
    .Faccion.FuerzasCaos = CByte(RS!ejercitocaos)
    .Faccion.CiudadanosMatados = CLng(RS!ciudmatados)
    .Faccion.CriminalesMatados = CLng(RS!crimmatados)
    .Faccion.RecibioArmaduraCaos = CByte(RS!rarcaos)
    .Faccion.RecibioArmaduraReal = CByte(RS!rarreal)
    .Faccion.RecibioExpInicialCaos = CByte(RS!rexcaos)
    .Faccion.RecibioExpInicialReal = CByte(RS!rexreal)
    .Faccion.RecompensasCaos = CLng(RS!reccaos)
    .Faccion.RecompensasReal = CLng(RS!recreal)
    .Faccion.Reenlistadas = CByte(RS!Reenlistadas)
    .Faccion.NivelIngreso = CInt(RS!NivelIngreso)
    .Faccion.FechaIngreso = RS!FechaIngreso
    .Faccion.MatadosIngreso = CInt(RS!MatadosIngreso)
    .Faccion.NextRecompensa = CInt(RS!NextRecompensa)
    
    .flags.Muerto = CByte(RS!Muerto)
    .flags.Escondido = CByte(RS!Escondido)
    .flags.Advertencias = CByte(RS!Advertencias)
    
    .flags.Hambre = CByte(RS!Hambre)
    .flags.Sed = CByte(RS!Sed)
    .flags.Desnudo = CByte(RS!Desnudo)
    .flags.Navegando = CByte(RS!Navegando)
    '.flags.Montando = CByte(RS!montando)
    .flags.QueMontura = CByte(RS!QueMontura)
    .flags.Envenenado = CByte(RS!Envenenado)
    .flags.Inmovilizado = CByte(RS!Inmovilizado)
    .flags.Paralizado = CByte(RS!Paralizado)
    .flags.Seguro = CByte(RS!Seguro)
    '.SerialHD = val(RS!serialhd)
    
    If .flags.Paralizado = 1 Then
        .Counters.Paralisis = IntervaloParalizado
    End If
    
        
    If .flags.Muerto = 0 Then
        .Char = .OrigChar
    Else
        .Char.body = iCuerpoMuerto
        .Char.Head = iCabezaMuerto
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
    End If
    
    
    .Counters.Pena = CLng(RS!Pena)
      
    
    'Obtiene el indice-objeto del arma
    .Invent.WeaponEqpSlot = CByte(RS!WeaponEqpSlot)
    
    'Obtiene el indice-objeto del armadura
    .Invent.ArmourEqpSlot = CByte(RS!ArmourEqpSlot)
    
    'Obtiene el indice-objeto del escudo
    .Invent.EscudoEqpSlot = CByte(RS!EscudoEqpSlot)
    
    'Obtiene el indice-objeto del casco
    .Invent.CascoEqpSlot = CByte(RS!CascoEqpSlot)
    
    'Obtiene el indice-objeto barco
    .Invent.BarcoSlot = CByte(RS!BarcoSlot)
    
    'Obtiene el indice-objeto municion
    .Invent.MunicionEqpSlot = CByte(RS!municionslot)
    
    '[Alejo]
    'Obtiene el indice-objeto anilo
    .Invent.AnilloEqpSlot = CByte(RS!anilloslot)
    
    
    '12/11/15 Irongete: Desactivo el guardado de las mascotas entre logins y logouts
    '.NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))
    'Dim NpcIndex As Integer
    'For LoopC = 1 To MAXMASCOTAS
    '    .MascotasType(LoopC) = val(UserFile.GetValue("MASCOTAS", "MAS" & LoopC))
    'Next LoopC
  
    
    
    '05/11/2015 Irongete: Cargar los objetos de la boveda
    Dim i As Integer
    Set RS = SQL.Execute("SELECT slot,objeto FROM rel_cuenta_boveda WHERE id_cuenta = '" & .CuentaId & "'")
    UserList(UserIndex).BancoInvent.NroItems = CInt(RS.RecordCount)
    i = 1
    While Not RS.EOF
        ln = RS!Objeto
        UserList(UserIndex).BancoInvent.Object(i).ObjIndex = CInt(ReadField(1, ln, 45))
        UserList(UserIndex).BancoInvent.Object(i).Amount = CInt(ReadField(2, ln, 45))
        i = i + 1
        RS.MoveNext
    Wend
       
    '05/11/2015 Irongete: Cargar los objetos del inventario
    Set RS = SQL.Execute("SELECT slot,objeto FROM rel_personaje_inventario WHERE id_personaje = '" & .id & "'")
    .Invent.NroItems = CInt(RS.RecordCount)
    i = 1
    While Not RS.EOF
        ln = RS!Objeto
        .Invent.Object(i).ObjIndex = CInt(ReadField(1, ln, 45))
        .Invent.Object(i).Amount = CInt(ReadField(2, ln, 45))
        .Invent.Object(i).Equipped = CByte(ReadField(3, ln, 45))
        i = i + 1
        RS.MoveNext
    Wend
     
     
    '14/02/2016 Lorwik: Terminamos de cargar todo.
    '¡¡ESTO TIENES QUE IR AL FINAL O PODRAS EQUIPAR VARIOS EQUIPOS!!
    If .Invent.WeaponEqpSlot > 0 Then
        .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
    End If
     
    If .Invent.ArmourEqpSlot > 0 Then
        .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
        .flags.Desnudo = 0
    Else
        .flags.Desnudo = 1
    End If
     
    If .Invent.EscudoEqpSlot > 0 Then
        .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex
    End If
     
    If .Invent.CascoEqpSlot > 0 Then
        .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex
    End If
     
    If .Invent.BarcoSlot > 0 Then
        .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex
    End If
     
    If .Invent.MunicionEqpSlot > 0 Then
        .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex
    End If
    
    If .Invent.AnilloEqpSlot > 0 Then
        .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex
    End If
End With
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = vbNullString
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
  
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

If frmMain.Visible Then frmMain.txStatus.Text = "Cargando backup."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call ModAreas.generarIDAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
    
    For Map = 1 To NumMaps
        If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map
        Else
            tFileName = App.Path & MapPath & "Mapa" & Map
        End If
        
        Call CargarMapa(Map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Sub LoadMapData()

If frmMain.Visible Then frmMain.txStatus.Text = "Cargando mapas..."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call ModAreas.generarIDAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        
        tFileName = App.Path & MapPath & "Mapa" & Map
        Call CargarMapa(Map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByVal MAPFl As String)
On Error GoTo errh

Dim fh As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As Long
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim NPCs() As tDatosNPC
Dim TEs() As tDatosTE
Dim MapSize As tMapSize
Dim MapDat As tMapDat

Dim i As Long
Dim j As Long

    If Not FileExist(App.Path & "\Maps\Mapa" & Map & ".csm", vbNormal) Then
        Debug.Print "El arhivo " & App.Path & "\Maps\Mapa" & Map & ".csm" & " no existe."
        Exit Sub
    End If
    
    fh = FreeFile
    Open App.Path & "\Maps\Mapa" & Map & ".csm" For Binary Access Read As fh
        Get #fh, , MH
        Get #fh, , MapSize
        Get #fh, , MapDat
        
        ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
        
        Get #fh, , L1
        
        With MH
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs
                For i = 1 To .NumeroBloqueados
                    MapData(Map, Blqs(i).X, Blqs(i).Y).Blocked = 1
                Next i
            End If
            
            If .NumeroLayers(2) > 0 Then
                ReDim L2(1 To .NumeroLayers(2))
                Get #fh, , L2
                For i = 1 To .NumeroLayers(2)
                    MapData(Map, L2(i).X, L2(i).Y).Graphic(2) = L2(i).GrhIndex
                Next i
            End If
            
            If .NumeroLayers(3) > 0 Then
                ReDim L3(1 To .NumeroLayers(3))
                Get #fh, , L3
                For i = 1 To .NumeroLayers(3)
                    MapData(Map, L3(i).X, L3(i).Y).Graphic(3) = L3(i).GrhIndex
                Next i
            End If
            
            If .NumeroLayers(4) > 0 Then
                ReDim L4(1 To .NumeroLayers(4))
                Get #fh, , L4
                For i = 1 To .NumeroLayers(4)
                    MapData(Map, L4(i).X, L4(i).Y).Graphic(4) = L4(i).GrhIndex
                Next i
            End If
            
            If .NumeroTriggers > 0 Then
                ReDim Triggers(1 To .NumeroTriggers)
                Get #fh, , Triggers
                For i = 1 To .NumeroTriggers
                    MapData(Map, Triggers(i).X, Triggers(i).Y).trigger = Triggers(i).trigger
                Next i
            End If
            
            If .NumeroParticulas > 0 Then
                ReDim Particulas(1 To .NumeroParticulas)
                Get #fh, , Particulas
                'For i = 1 To .NumeroParticulas
                    'MapData(Particulas(i).x, Particulas(i).y).particle_group_index = General_Particle_Create(Particulas(i).Particula, Particulas(i).x, Particulas(i).y)
                'Next i
            End If
            
            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
                Get #fh, , Luces
                'For i = 1 To .NumeroLuces
                    'Call frmMain.engine.Light_Create(Luces(i).x, Luces(i).y, Luces(i).color, Luces(i).Rango)
                'Next i
            End If
            
            If .NumeroOBJs > 0 Then
                ReDim Objetos(1 To .NumeroOBJs)
                Get #fh, , Objetos
                For i = 1 To .NumeroOBJs
                    MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.ObjIndex = Objetos(i).ObjIndex
                    MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.Amount = Objetos(i).ObjAmmount
                Next i
            End If
                
            If .NumeroNPCs > 0 Then
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs
                For i = 1 To .NumeroNPCs
                    MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = NPCs(i).NpcIndex
                    If MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex > 0 Then
                        Dim npcfile As String
                
                        npcfile = DatPath & "NPCs.dat"
       
                        If val(GetVar(npcfile, "NPC" & MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex, "PosOrig")) = 1 Then
                            MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex)
                            NPCList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.Map = Map
                            NPCList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.X = NPCs(i).X
                            NPCList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.Y = NPCs(i).Y
                        Else
                            MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex)
                        End If
                        If Not MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = 0 Then
                            NPCList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.Map = Map
                            NPCList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.X = NPCs(i).X
                            NPCList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.Y = NPCs(i).Y
       
                            Call MakeNPCChar(True, 0, MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex, Map, NPCs(i).X, NPCs(i).Y)
                        End If
                    End If
                Next i
            End If
                
            If .NumeroTE > 0 Then
                ReDim TEs(1 To .NumeroTE)
                Get #fh, , TEs
                For i = 1 To .NumeroTE
                    MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                    MapData(Map, TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                    MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
                Next i
            End If
            
        End With
    
    Close fh
    
        
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If L1(i, j) > 0 Then
                MapData(Map, j, i).Graphic(1) = L1(j, i)
            End If
        Next i
    Next j
    
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    'Call Leer.Initialize(MAPFl & ".dat")
    
    'Cargamos los extras
    With MapInfo(Map)
        
        .MagiaSinEfecto = MapDat.MagiaSinEfecto
        .InviSinEfecto = MapDat.InviSinEfecto
        .ResuSinEfecto = MapDat.ResuSinEfecto
        .SePuedeDomar = MapDat.SePuedeDomar
        .NoEncriptarMP = MapDat.NoEncriptarMP

        .Pk = MapDat.battle_mode
        
        .Terreno = MapDat.terrain
        .zona = MapDat.zone
        .Restringir = RestrictStringToByte(MapDat.restrict_mode)
        .BackUp = MapDat.backup_mode
        .lvlMinimo = MapDat.lvlMinimo
        .MapName = MapDat.map_name
    End With
    
    'Set Leer = Nothing
    
Exit Sub

errh:
    'Call LogError("Error cargando mapa: " & map & " - Pos: " & .X & "," & Y & "." & Err.description)
    Set Leer = Nothing
End Sub

Function LoadSini() As Boolean

Dim Temporal As Long


If frmMain.Visible Then frmMain.txStatus.Text = "Cargando info de inicio del server."

'&&&&&&&&&&&&&&&&&&&&&&SQL&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

InfSQL.Driver = GetVar(IniPath & "Server.ini", "SQL", "Driver")
InfSQL.Server = GetVar(IniPath & "Server.ini", "SQL", "Server")
InfSQL.Database = GetVar(IniPath & "Server.ini", "SQL", "database")
InfSQL.Port = GetVar(IniPath & "Server.ini", "SQL", "Port")
InfSQL.Name = GetVar(IniPath & "Server.ini", "SQL", "Name")
InfSQL.Pass = GetVar(IniPath & "Server.ini", "SQL", "Pass")
InfSQL.Modo = GetVar(IniPath & "Server.ini", "SQL", "modo")
'*********************************************************

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

NPCReyCastle = val(GetVar(IniPath & "Server.ini", "INIT", "NPCReyCastle"))
NPCDefensorFortaleza = val(GetVar(IniPath & "Server.ini", "INIT", "NPCDefensorFortaleza"))

Puerto = val(GetVar(IniPath & "Server.ini", "CONEXION", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")

PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))
RequiereValidacionACC = val(GetVar(IniPath & "Server.ini", "ini", "RequiereValidacionACC"))

MultiplicadorGP = GetVar(IniPath & "Server.ini", "INIT", "GPUser")
MultiplicadorGPN = GetVar(IniPath & "Server.ini", "INIT", "GPNpc")

EnTesting = val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))

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

IntervaloPuedeMakrear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMakreo"))

IntervaloMorphPJ = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMorphPJ"))

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
IntervaloPuedeSerAtacado = 1000 '1 Seg
IntervaloPuedeCambiardeMapa = 5000 '5 Seg

frmMain.NPC_AI.interval = val(GetVar(IniPath & "Server.ini", "TIMER", "NPC_AI"))
FrmInterv.txtAI.Text = frmMain.NPC_AI.interval

frmMain.npcataca.interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.interval

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

User_AtacarMelee = val(GetVar(IniPath & "Server.ini", "INTERVALO", "User_AtacarMelee"))
FrmInterv.txtPuedeAtacar.Text = User_AtacarMelee


User_LanzarMagia = val(GetVar(IniPath & "Server.ini", "INTERVALO", "User_LanzarMagia"))
FrmInterv.txtPuedeAtacar.Text = User_LanzarMagia

'TODO : Agregar estos intervalos al form!!!
IntervaloMagiaGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe"))
IntervaloGolpeMagia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia"))
IntervaloGolpeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeUsar"))

MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
If MinutosWs < 60 Then MinutosWs = 30

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))

IntervaloOculto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto"))

'&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
  
recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
iniDragDrop = CByte(val(GetVar(IniPath & "Server.ini", "OPCIONES", "DragDrop")))
iniTirarOBJZonaSegura = CByte(val(GetVar(IniPath & "Server.ini", "OPCIONES", "TirarOBJZonaSegura")))
iniAutoSacerdote = IIf(GetVar(IniPath & "Server.ini", "OPCIONES", "AutoSacerdote") = 1, True, False)
iniSacerdoteCuraVeneno = IIf(GetVar(IniPath & "Server.ini", "OPCIONES", "SacerdoteCuraVeneno") = 1, True, False)
  
'Max users
Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As User
End If

'&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
'Se agregó en LoadBalance y en el Balance.dat
'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))

''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
Call Statistics.Initialize

Call CargarCanjes

Call ConsultaPopular.LoadData

Call LoadAntiCheat

LoadSini = True

End Function

Public Sub CargarCastillos()
    '*******************CASTILLOS***********************
    Castillos(1).nombre = "el Castillo Norte"
    Castillos(1).Ubicacion = 25
    Castillos(1).Mapa = 25
    Castillos(1).Dueño = GetIdDueñoCastillo(1)
    Castillos(1).FechaHora = GetFechaHoraCastillo(1)
    
    Castillos(2).nombre = "el Castillo Este"
    Castillos(2).Ubicacion = 26
    Castillos(2).Mapa = 26
    Castillos(2).Dueño = GetIdDueñoCastillo(2)
    Castillos(2).FechaHora = GetFechaHoraCastillo(2)
    
    Castillos(3).nombre = "el Castillo Sur"
    Castillos(3).Ubicacion = 15
    Castillos(3).Mapa = 15
    Castillos(3).Dueño = GetIdDueñoCastillo(3)
    Castillos(3).FechaHora = GetFechaHoraCastillo(3)
    
    Castillos(4).nombre = "el Castillo Oeste"
    Castillos(4).Ubicacion = 27
    Castillos(4).Mapa = 27
    Castillos(4).Dueño = GetIdDueñoCastillo(4)
    Castillos(4).FechaHora = GetFechaHoraCastillo(4)
    
    Castillos(5).nombre = "la Fortaleza"
    Castillos(5).Ubicacion = 28
    Castillos(5).Mapa = 28
    Castillos(5).Dueño = GetIdDueñoCastillo(5)
    Castillos(5).FechaHora = GetFechaHoraCastillo(5)
    '*******************CASTILLOS***********************
End Sub

Public Sub LoadAntiCheat()
    Dim i As Integer
    Dim CnfInterval As String
    
    CnfInterval = "intervalos.ini"
 
    Lac_Camina = CLng(val(GetVar$(IniPath & CnfInterval, "LACANTICHEAT", "Caminar")))
    Lac_Lanzar = CLng(val(GetVar$(IniPath & CnfInterval, "LACANTICHEAT", "Lanzar")))
    Lac_Usar = CLng(val(GetVar$(IniPath & CnfInterval, "LACANTICHEAT", "Usar")))
    Lac_Tirar = CLng(val(GetVar$(IniPath & CnfInterval, "LACANTICHEAT", "Tirar")))
    Lac_Pociones = CLng(val(GetVar$(IniPath & CnfInterval, "LACANTICHEAT", "Pociones")))
    Lac_Pegar = CLng(val(GetVar$(IniPath & CnfInterval, "LACANTICHEAT", "Pegar")))
 
    For i = 1 To MaxUsers
        ResetearLac i
    Next
   
End Sub

Sub CargarCanjes()
    Dim i As Integer
    CantPremios = val(GetVar(DatPath & "Premios.dat", "INIT", "CantPremios"))
    If CantPremios > 0 Then 'Evitamos el error
        ReDim PremiosInfo(1 To CantPremios) As tPremios
        For i = 1 To CantPremios
            PremiosInfo(i).ObjIndex = val(GetVar(DatPath & "Premios.dat", "PREMIO" & i, "ObjIndex"))
            PremiosInfo(i).Puntos = val(GetVar(DatPath & "Premios.dat", "PREMIO" & i, "Puntos"))
            PremiosInfo(i).cantidad = val(GetVar(DatPath & "Premios.dat", "PREMIO" & i, "Cantidad"))
        Next i
    End If
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, value, File
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 20/03/2016
'Saves the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************

On Error GoTo errhandler

Dim OldUserHead As Long

With UserList(UserIndex)

    'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
    'clase=0 es el error, porq el enum empieza de 1!!
    If .clase = 0 Or .Stats.ELV = 0 Then
        Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
        Exit Sub
    End If
    
    
    If .flags.Mimetizado = 1 Then
        .Char.body = .CharMimetizado.body
        .Char.Head = .CharMimetizado.Head
        .Char.CascoAnim = .CharMimetizado.CascoAnim
        .Char.ShieldAnim = .CharMimetizado.ShieldAnim
        .Char.WeaponAnim = .CharMimetizado.WeaponAnim
        .Counters.Mimetismo = 0
        .flags.Mimetizado = 0
    End If
    
    '13/02/2016 Lorwik: Si estas Morph te devuelve a la normalidad.
    If .flags.Morph = 1 Then
        .Char.body = .OrigChar.body
        .Char.Head = .OrigChar.Head
        .Char.CascoAnim = .OrigChar.CascoAnim
        .Char.ShieldAnim = .OrigChar.ShieldAnim
        .Char.WeaponAnim = .OrigChar.WeaponAnim
        .Counters.Morph = 0
        .flags.Morph = 0
     End If
     
    'Devuelve el head de muerto
    If .flags.Muerto = 1 Then
        .Char.Head = iCabezaMuerto
    End If
    

    If .flags.Muerto = 1 Then
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT head FROM personaje WHERE id = '" & .id & "'")
        OldUserHead = .Char.Head
        .Char.Head = RS!Head
        RS.Close
    End If
    
    #If ConUpTime Then
        Dim TempDate As Date
        TempDate = Now - .LogOnTime
        .LogOnTime = Now
        .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
        .UpTime = .UpTime
    #End If
    
    
    '04/11/2015 Irongete: Guardo los datos en SQL
    Set RS = New ADODB.Recordset
    Dim LoopC As Integer
    Dim query As String
    Dim NroMascotas As Long
    Dim cad As Integer
    
    'calculo el promedio de la reputacion
    Dim L As Long
    L = (-UserList(UserIndex).Reputacion.AsesinoRep) + (-UserList(UserIndex).Reputacion.BandidoRep) + UserList(UserIndex).Reputacion.BurguesRep + (-UserList(UserIndex).Reputacion.LadronesRep) + UserList(UserIndex).Reputacion.NobleRep + UserList(UserIndex).Reputacion.PlebeRep
    L = L / 6

    'VALUES
    query = query & "UPDATE personaje SET "
    query = query & "heading = '" & .OrigChar.heading & "', "
    query = query & "head = '" & .OrigChar.Head & "', "
    query = query & "body = '" & .Char.body & "', "
    query = query & "arma = '" & .Char.WeaponAnim & "', "
    query = query & "escudo = '" & .Char.ShieldAnim & "', "
    query = query & "casco = '" & .Char.CascoAnim & "', "
    query = query & "cantidaditems = '" & .BancoInvent.NroItems & "', "
    query = query & "npcsmuertes = '" & .Stats.NPCsMuertos & "', "
    query = query & "usermuertes = '" & .Stats.UsuariosMatados & "', "
    query = query & "nummonturas = '" & .Stats.NUMMONTURAS & "', "
    query = query & "dragcredits = '" & .Stats.DragCredits & "', "
    query = query & "elo = '" & .Stats.ELO & "', "
    query = query & "elu = '" & .Stats.ELU & "', "
    query = query & "elv = '" & .Stats.ELV & "', "
    query = query & "exp = '" & .Stats.Exp & "', "
    query = query & "minham = '" & .Stats.MinHam & "', "
    query = query & "maxham = '" & .Stats.MaxHam & "', "
    query = query & "minagu = '" & .Stats.MinAGU & "', "
    query = query & "maxagu = '" & .Stats.MaxAGU & "', "
    query = query & "minhit = '" & .Stats.MinHIT & "', "
    query = query & "maxhit = '" & .Stats.MaxHIT & "', "
    query = query & "minman = '" & .Stats.MinMAN & "', "
    query = query & "maxman = '" & .Stats.MaxMAN & "', "
    query = query & "minsta = '" & .Stats.MinSta & "', "
    query = query & "maxsta = '" & .Stats.MaxSta & "', "
    query = query & "minhp = '" & .Stats.MinHP & "', "
    query = query & "maxhp = '" & .Stats.MaxHP & "', "
    query = query & "banco = '" & .Stats.Banco & "', "
    query = query & "gld = '" & .Stats.GLD & "', "
    query = query & "promedio = '" & CStr(L) & "', "
    query = query & "plebe = '" & .Reputacion.PlebeRep & "', "
    query = query & "nobles = '" & .Reputacion.NobleRep & "', "
    query = query & "ladrones = '" & .Reputacion.LadronesRep & "', "
    query = query & "burguesia = '" & .Reputacion.BurguesRep & "', "
    query = query & "bandido = '" & .Reputacion.BandidoRep & "', "
    query = query & "asesino = '" & .Reputacion.AsesinoRep & "', "
    query = query & "anilloslot = '" & .Invent.AnilloEqpSlot & "', "
    query = query & "municionslot = '" & .Invent.MunicionEqpSlot & "', "
    query = query & "barcoslot = '" & .Invent.BarcoSlot & "', "
    query = query & "escudoeqpslot = '" & .Invent.EscudoEqpSlot & "', "
    query = query & "cascoeqpslot = '" & .Invent.CascoEqpSlot & "', "
    query = query & "armoureqpslot = '" & .Invent.ArmourEqpSlot & "', "
    query = query & "weaponeqpslot = '" & .Invent.WeaponEqpSlot & "', "
    query = query & "nextrecompensa = '" & .Faccion.NextRecompensa & "', "
    query = query & "matadosingreso = '" & .Faccion.MatadosIngreso & "', "
    query = query & "fechaingreso = '" & .Faccion.FechaIngreso & "', "
    query = query & "nivelingreso = '" & .Faccion.NivelIngreso & "', "
    query = query & "reenlistadas = '" & .Faccion.Reenlistadas & "', "
    query = query & "recreal = '" & .Faccion.RecompensasReal & "', "
    query = query & "reccaos = '" & .Faccion.RecompensasCaos & "', "
    query = query & "rexreal = '" & .Faccion.RecibioExpInicialReal & "', "
    query = query & "rexcaos = '" & .Faccion.RecibioExpInicialCaos & "', "
    query = query & "rarreal = '" & .Faccion.RecibioArmaduraReal & "', "
    query = query & "rarcaos = '" & .Faccion.RecibioArmaduraCaos & "', "
    query = query & "crimmatados = '" & .Faccion.CriminalesMatados & "', "
    query = query & "ciudmatados = '" & .Faccion.CiudadanosMatados & "', "
    query = query & "ejercitocaos = '" & .Faccion.FuerzasCaos & "', "
    query = query & "ejercitoreal = '" & .Faccion.ArmadaReal & "', "
    query = query & "pena = '" & .Counters.Pena & "', "
    query = query & "pertenececaos = '" & IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0") & "', "
    query = query & "pertenece = '" & IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0") & "', "
    query = query & "serialhd = '" & .flags.SerialHD & "', "
    query = query & "inmovilizado = '" & .flags.Inmovilizado & "', "
    query = query & "paralizado = '" & .flags.Paralizado & "', "
    query = query & "envenenado = '" & .flags.Envenenado & "', "
    query = query & "quemontura = '" & .flags.QueMontura & "', "
    query = query & "navegando = '" & .flags.Navegando & "', "
    query = query & "position = '" & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y & "', "
    query = query & "descripcion = '" & .desc & "', "
    query = query & "muerto = '" & .flags.Muerto & "', "
    query = query & "escondido = '" & .flags.Escondido & "', "
    query = query & "advertencias = '" & .flags.Advertencias & "', "
    query = query & "hambre = '" & .flags.Hambre & "', "
    query = query & "sed = '" & .flags.Sed & "', "
    query = query & "desnudo = '" & .flags.Desnudo & "', "
    query = query & "ban = '" & .flags.Ban & "', "
    query = query & "seguro = '" & .flags.Seguro & "' "
    query = query & "WHERE id = '" & .id & "'"
    Set RS = SQL.Execute(query)

    '05/11/2015 Irongete: Guardo los objetos del inventario
    Dim loopd As Integer
    For loopd = 1 To MAX_INVENTORY_SLOTS
        query = "UPDATE rel_personaje_inventario SET "
        query = query & "objeto = '" & .Invent.Object(loopd).ObjIndex & "-" & .Invent.Object(loopd).Amount & "-" & .Invent.Object(loopd).Equipped & "' "
        query = query & "WHERE id_personaje = '" & .id & "' AND slot = '" & loopd & "'"
        Set RS = SQL.Execute(query)
    Next loopd
    
        
    '05/11/2015 Irongete: Guardo los hechizos
    For LoopC = 1 To MAXUSERHECHIZOS
        query = "UPDATE rel_personaje_hechizo SET "
        query = query & "hechizo = '" & .Stats.UserHechizos(LoopC) & "' "
        query = query & "WHERE id_personaje = '" & .id & "' AND slot = '" & LoopC & "'"
        Set RS = SQL.Execute(query)
    Next
    
    '05/11/2015 Irongete: Guardar los atributos
    Dim valor As Integer
    For LoopC = 1 To UBound(.Stats.UserAtributos)
        query = "UPDATE rel_personaje_atributo SET "
        query = query & "valor = '" & valor & "' "
        query = query & "WHERE id_personaje = '" & .id & "' AND atributo = '" & valor & "'"
        Set RS = SQL.Execute(query)
    Next

    
    '05/11/2015 Irongete: Guardar las skills
    For LoopC = 1 To UBound(.Stats.UserSkills)
        query = "UPDATE rel_personaje_skill SET "
        query = query & "valor = '" & CStr(.Stats.UserSkills(LoopC)) & "' "
        query = query & "WHERE id_personaje = '" & .id & "' AND skill = '" & LoopC & "'"
        Set RS = SQL.Execute(query)
    Next

    
    '05/11/2015 Irongete: Guardo los objetos de la boveda
    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        query = "UPDATE rel_cuenta_boveda SET "
        query = query & "objeto = '" & .BancoInvent.Object(loopd).ObjIndex & "-" & .BancoInvent.Object(loopd).Amount & "' "
        query = query & "WHERE id_cuenta = '" & .CuentaId & "' AND slot = '" & loopd & "'"
        Set RS = SQL.Execute(query)
    Next loopd

    For LoopC = 1 To .Stats.NUMMONTURAS
        query = "UPDATE montura SET "
        query = query & "nombre = '" & .flags.Montura(LoopC).nombre & "', "
        query = query & "level = '" & .flags.Montura(LoopC).MonturaLevel & "', "
        query = query & "skills = '" & .flags.Montura(LoopC).Skills & "', "
        query = query & "elu = '" & .flags.Montura(LoopC).ELU & "', "
        query = query & "exp = '" & .flags.Montura(LoopC).Exp & "', "
        query = query & "ataque = '" & .flags.Montura(LoopC).Ataque & "', "
        query = query & "defensa = '" & .flags.Montura(LoopC).Defensa & "', "
        query = query & "atmagia = '" & .flags.Montura(LoopC).AtMagia & "', "
        query = query & "defmagia = '" & .flags.Montura(LoopC).DefMagia & "', "
        query = query & "evasion = '" & .flags.Montura(LoopC).Evasion & "' "
        query = query & "speed = '" & .flags.Montura(LoopC).Speed & "' "
        query = query & "WHERE id = '" & .flags.Montura(LoopC).id & "'"
        Set RS = SQL.Execute(query)
    Next
    
End With

Exit Sub

errhandler:
Debug.Print Err.Description

End Sub

Sub SaveNewUser(ByVal UserIndex As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 20/03/2016
'Saves the New Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************

On Error GoTo errhandler

Dim OldUserHead As Long

With UserList(UserIndex)

    'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
    'clase=0 es el error, porq el enum empieza de 1!!
    If .clase = 0 Or .Stats.ELV = 0 Then
        Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
        Exit Sub
    End If
    
    '04/11/2015 Irongete: Guardo los datos en SQL
    Set RS = New ADODB.Recordset
    Dim LoopC As Integer
    Dim query As String
    Dim NroMascotas As Long
    Dim cad As Integer
    
    'FIELDS
    query = "INSERT INTO personaje (id_cuenta,nombre,genero,raza,clase,heading,head,body,arma,escudo,casco,uptime,lastip1,position,"
    query = query & "descripcion,muerto,escondido,advertencias,hambre,sed,desnudo,ban,navegando,quemontura,envenenado,inmovilizado,paralizado,serialhd,"
    query = query & "pertenece,pertenececaos,pena,ejercitoreal,ejercitocaos,ciudmatados,crimmatados,rarcaos,rarreal,rexcaos,rexreal,reccaos,"
    query = query & "recreal,reenlistadas,nivelingreso,fechaingreso,matadosingreso,nextrecompensa,"
    query = query & "weaponeqpslot,armoureqpslot,cascoeqpslot,escudoeqpslot,barcoslot,municionslot,anilloslot,asesino,bandido,burguesia,ladrones,"
    query = query & "nobles,plebe,promedio,gld,banco,maxhp,minhp,maxsta,minsta,maxman,minman,maxhit,minhit,maxagu, "
    query = query & "minagu,maxham,minham,exp,elv,elu,elo,dragcredits,nummonturas,usermuertes,npcsmuertes,cantidaditems,seguro,borrado) "
    
    'calculo el promedio de la reputacion
    Dim L As Long
    L = (-UserList(UserIndex).Reputacion.AsesinoRep) + (-UserList(UserIndex).Reputacion.BandidoRep) + UserList(UserIndex).Reputacion.BurguesRep + (-UserList(UserIndex).Reputacion.LadronesRep) + UserList(UserIndex).Reputacion.NobleRep + UserList(UserIndex).Reputacion.PlebeRep
    L = L / 6

    'VALUES
    query = query & "VALUES ('" & UserList(UserIndex).CuentaId & "', '" & .Name & "', '" & .genero & "', '" & .raza & "', '" & .clase & "', '" & .Char.heading & "', "
    query = query & "'" & .Char.Head & "', '" & .Char.body & "', '" & .Char.WeaponAnim & "', '" & .Char.ShieldAnim & "', '" & .Char.CascoAnim & "', "
    query = query & "'" & .UpTime & "', '" & .IPLong & "', '" & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y & "', '" & .desc & "', '" & .flags.Muerto & "', '" & .flags.Escondido & "', "
    query = query & "'" & .flags.Advertencias & "', '" & .flags.Hambre & "', '" & .flags.Sed & "', '" & .flags.Desnudo & "', '" & .flags.Ban & "', '" & .flags.Navegando & "', "
    query = query & "'" & .flags.QueMontura & "', '" & .flags.Envenenado & "', '" & .flags.Inmovilizado & "', '" & .flags.Paralizado & "', '" & .flags.SerialHD & "', '" & IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0") & "', "
    query = query & "'" & IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0") & "', '" & .Counters.Pena & "', '" & .Faccion.ArmadaReal & "', '" & .Faccion.FuerzasCaos & "', "
    query = query & "'" & .Faccion.CiudadanosMatados & "', '" & .Faccion.CriminalesMatados & "', '" & .Faccion.RecibioArmaduraCaos & "', '" & .Faccion.RecibioArmaduraReal & "', '" & .Faccion.RecibioExpInicialCaos & "', "
    query = query & "'" & .Faccion.RecibioExpInicialReal & "', '" & .Faccion.RecompensasCaos & "', '" & .Faccion.RecompensasReal & "', '" & .Faccion.Reenlistadas & "', '" & .Faccion.NivelIngreso & "', "
    query = query & "'" & .Faccion.FechaIngreso & "', '" & .Faccion.MatadosIngreso & "', '" & .Faccion.NextRecompensa & "', '" & .Invent.WeaponEqpSlot & "', "
    query = query & "'" & .Invent.ArmourEqpSlot & "', '" & .Invent.CascoEqpSlot & "', '" & .Invent.EscudoEqpSlot & "', '" & .Invent.BarcoSlot & "', "
    query = query & "'" & .Invent.MunicionEqpSlot & "', '" & .Invent.AnilloEqpSlot & "', '" & .Reputacion.AsesinoRep & "', '" & .Reputacion.BandidoRep & "', "
    query = query & "'" & .Reputacion.BurguesRep & "', '" & .Reputacion.LadronesRep & "', '" & .Reputacion.NobleRep & "', '" & .Reputacion.PlebeRep & "', '" & CStr(L) & "', '" & .Stats.GLD & "', '" & .Stats.Banco & "', "
    query = query & "'" & .Stats.MaxHP & "', '" & .Stats.MinHP & "', '" & .Stats.MaxSta & "', '" & .Stats.MinSta & "', '" & .Stats.MaxMAN & "', '" & .Stats.MinMAN & "', '" & .Stats.MaxHIT & "', "
    query = query & "'" & .Stats.MinHIT & "', '" & .Stats.MaxAGU & "', '" & .Stats.MinAGU & "', '" & .Stats.MaxHam & "', '" & .Stats.MinHam & "', '" & .Stats.Exp & "', '" & .Stats.ELV & "', "
    query = query & "'" & .Stats.ELU & "', '" & .Stats.ELO & "', '" & .Stats.DragCredits & "', '" & .Stats.NUMMONTURAS & "', '" & .Stats.UsuariosMatados & "', '" & .Stats.NPCsMuertos & "', '" & .BancoInvent.NroItems & "', '" & .flags.Seguro & "', '0' ) "
    
    Debug.Print query
    
    Set RS = SQL.Execute(query)
    
    '05/11/2015 Irongete: Obtener la ID del personaje que se está guardando
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id FROM personaje WHERE nombre = '" & .Name & "'")
    .id = RS!id

    '05/11/2015 Irongete: Guardo los objetos del inventario
    Dim loopd As Integer
    For loopd = 1 To MAX_INVENTORY_SLOTS
        query = "INSERT INTO rel_personaje_inventario (id_personaje, slot, objeto) VALUES ('" & .id & "', '" & loopd & "', '" & .Invent.Object(loopd).ObjIndex & "-" & .Invent.Object(loopd).Amount & "-" & .Invent.Object(loopd).Equipped & "') "
        Set RS = SQL.Execute(query)
    Next loopd
    
        
    '05/11/2015 Irongete: Guardo los hechizos
    For LoopC = 1 To MAXUSERHECHIZOS
        query = "INSERT INTO rel_personaje_hechizo (id_personaje, slot, hechizo) VALUES ('" & .id & "', '" & LoopC & "', '" & .Stats.UserHechizos(LoopC) & "') "
        'Debug.Print query
        Set RS = SQL.Execute(query)
        
    Next
    
    '05/11/2015 Irongete: Guardar los atributos
    Dim valor As Integer
    For LoopC = 1 To UBound(.Stats.UserAtributos)
        If Not .flags.TomoPocion Then
            valor = CInt(.Stats.UserAtributos(LoopC))
        Else
            valor = CInt(.Stats.UserAtributosBackUP(LoopC))
        End If
        query = "INSERT INTO rel_personaje_atributo (id_personaje, atributo, valor) VALUES ('" & .id & "', '" & LoopC & "', '" & valor & "') "
        Set RS = SQL.Execute(query)
    Next

    
    '05/11/2015 Irongete: Guardar las skills
    For LoopC = 1 To UBound(.Stats.UserSkills)
        query = "INSERT INTO rel_personaje_skill (id_personaje, skill, valor) VALUES ('" & .id & "', '" & LoopC & "', '" & CStr(.Stats.UserSkills(LoopC)) & "') "
        Set RS = SQL.Execute(query)
    Next

    
    '05/11/2015 Irongete: Guardo los objetos de la boveda
    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        query = "INSERT INTO rel_cuenta_boveda (id_cuenta, slot, objeto) VALUES ('" & .CuentaId & "', '" & loopd & "', '" & .BancoInvent.Object(loopd).ObjIndex & "-" & .BancoInvent.Object(loopd).Amount & "') "
        Set RS = SQL.Execute(query)
    Next loopd
    
End With

Exit Sub

errhandler:
Debug.Print Err.Description

End Sub
Function criminal(ByVal UserIndex As Integer) As Boolean

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
criminal = (L < 0)

End Function

Sub BackUPnPc(NpcIndex As Integer)

Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer


NpcNumero = NPCList(NpcIndex).Numero

'If NpcNumero > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", NPCList(NpcIndex).Name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", NPCList(NpcIndex).desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Nivel", NPCList(NpcIndex).Nivel)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(NPCList(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(NPCList(NpcIndex).Char.body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "AnimAtaque", val(NPCList(NpcIndex).Char.AnimAtaque))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(NPCList(NpcIndex).Char.heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(NPCList(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(NPCList(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(NPCList(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(NPCList(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(NPCList(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(NPCList(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(NPCList(NpcIndex).GiveGLD))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(NPCList(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(NPCList(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(NPCList(NpcIndex).NPCType))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Quest", val(NPCList(NpcIndex).Quest))

'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(NPCList(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoNick", val(NPCList(NpcIndex).Stats.TipoNick))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(NPCList(NpcIndex).Stats.def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(NPCList(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(NPCList(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(NPCList(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(NPCList(NpcIndex).Stats.MinHP))


Call WriteVar(npcfile, "NPC" & NpcNumero, "SoloParty", val(NPCList(NpcIndex).flags.SoloParty))
Call WriteVar(npcfile, "NPC" & NpcNumero, "LanzaMensaje", val(NPCList(NpcIndex).flags.LanzaMensaje))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Mensaje", val(NPCList(NpcIndex).flags.Mensaje))
Call WriteVar(npcfile, "NPC" & NpcNumero, "AumentaPotencia", val(NPCList(NpcIndex).flags.AumentaPotencia))

Call WriteVar(npcfile, "NPC" & NpcNumero, "TiempoRetardoMax", val(NPCList(NpcIndex).flags.TiempoRetardoMax))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TiempoRetardoMin", val(NPCList(NpcIndex).flags.TiempoRetardoMin))

Call WriteVar(npcfile, "NPC" & NpcNumero, "Retardo", val(NPCList(NpcIndex).flags.Retardo))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Explota", val(NPCList(NpcIndex).flags.Explota))
Call WriteVar(npcfile, "NPC" & NpcNumero, "VerInvi", val(NPCList(NpcIndex).flags.VerInvi))
'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(NPCList(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(NPCList(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(NPCList(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(NPCList(NpcIndex).Invent.NroItems))
If NPCList(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, NPCList(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & NPCList(NpcIndex).Invent.Object(LoopC).Amount & "-" & NPCList(NpcIndex).Invent.Object(LoopC).ProbTirar)
   Next
End If


End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'Status
If frmMain.Visible Then frmMain.txStatus.Text = "Cargando backup Npc"

Dim npcfile As String

'If NpcNumber > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

NPCList(NpcIndex).Numero = NpcNumber
NPCList(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
NPCList(NpcIndex).desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
NPCList(NpcIndex).Nivel = GetVar(npcfile, "NPC" & NpcNumber, "Nivel")
If Not NPCList(NpcIndex).Nivel Then
    Debug.Print GetVar(npcfile, "NPC" & NpcNumber, "Name")

End If


NPCList(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
NPCList(NpcIndex).NPCType = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

NPCList(NpcIndex).Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
NPCList(NpcIndex).Char.AnimAtaque = val(GetVar(npcfile, "NPC" & NpcNumber, "AnimAtaque"))
NPCList(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
NPCList(NpcIndex).Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

NPCList(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
NPCList(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
NPCList(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
NPCList(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))


NPCList(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

NPCList(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

NPCList(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
NPCList(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
NPCList(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
NPCList(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
NPCList(NpcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
NPCList(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
NPCList(NpcIndex).Stats.TipoNick = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoNick"))

NPCList(NpcIndex).flags.SoloParty = val(GetVar(npcfile, "NPC" & NpcNumber, "SoloParty"))
NPCList(NpcIndex).flags.LanzaMensaje = GetVar(npcfile, "NPC" & NpcNumber, "LanzaMensaje")
NPCList(NpcIndex).flags.Mensaje = val(GetVar(npcfile, "NPC" & NpcNumber, "Mensaje"))
NPCList(NpcIndex).flags.AumentaPotencia = val(GetVar(npcfile, "NPC" & NpcNumber, "AumentaPotencia"))
NPCList(NpcIndex).flags.TiempoRetardoMax = val(GetVar(npcfile, "NPC" & NpcNumber, "TiempoRetardoMax"))
NPCList(NpcIndex).flags.TiempoRetardoMin = val(GetVar(npcfile, "NPC" & NpcNumber, "TiempoRetardoMin"))
NPCList(NpcIndex).flags.Retardo = val(GetVar(npcfile, "NPC" & NpcNumber, "Retardo"))
NPCList(NpcIndex).flags.Explota = val(GetVar(npcfile, "NPC" & NpcNumber, "Explota"))
NPCList(NpcIndex).flags.VerInvi = val(GetVar(npcfile, "NPC" & NpcNumber, "VerInvi"))

NPCList(NpcIndex).flags.ActivoPotencia = False
NPCList(NpcIndex).flags.DijoMensaje = False

Dim LoopC As Integer
Dim ln As String
NPCList(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If NPCList(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
        NPCList(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        NPCList(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
        NPCList(NpcIndex).Invent.Object(LoopC).ProbTirar = val(ReadField(3, ln, 45))
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        NPCList(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        NPCList(NpcIndex).Invent.Object(LoopC).Amount = 0
        NPCList(NpcIndex).Invent.Object(LoopC).ProbTirar = 0
    Next LoopC
End If



NPCList(NpcIndex).flags.Active = True
NPCList(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
NPCList(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
NPCList(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
NPCList(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
NPCList(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)


'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub
