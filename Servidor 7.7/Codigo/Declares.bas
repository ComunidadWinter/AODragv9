Attribute VB_Name = "Declaraciones"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'nati:
Public miembros As Integer
Public PorcentajeC As Integer
Public variablepuntos As Integer
Public puntosX As Integer

Public CuentaRegresiva As Integer
Public indexCuentaRegresiva As Integer
Public mapainvasion As Integer
'pluto:7.0
Public NOmbrelogro(1 To 34) As String


'PLUTO:6.9
Public TOPELANZAR As Integer
Public TOPEFLECHA As Integer
'pluto:6.2
Public COMPROBANDOMACRO As Boolean
Public CodigoMacro As Integer
Type Tclan
Nombre As String
numero As Byte
'Muertos() As String
End Type
'pluto:6.8
Public EventoDia As Byte
Public RecordEventoDia As Byte
Public BichoDelDia As Integer
Public NombreBichoDelDia As String
Public HOYESDIA As String
Public MsgTorneo As String
'pluto:6.5
Type RaidVivo
Activo As Byte
MiniRaids As Byte
End Type

Public NoDomarMontura As Boolean
Public TorneoClan(2) As Tclan
Public TClanOcupado As Byte
Public TClanNumero As Byte

Public Arx As String
Public RaidVivos(1 To 6) As RaidVivo
Public Minotauro As String
Public MinutosMinotauro As Byte
Public EstadoMinotauro As Byte

Public SolidoGirando As Byte
Public camaralenta As Integer
Public Puerto As Integer
'pluto:2.15
Public Joputa As String
Public joputa2 As Byte
Public AtaNorte As Byte
Public AtaSur As Byte
Public AtaEste As Byte
Public AtaOeste As Byte
Public AtaForta As Byte
Public MinutosPoder As Byte
'Public Conquistas As Boolean
'Public DueñoNix As Byte
'Public DueñoCaos As Byte
'Public DueñoUlla As Byte
'Public DueñoBander As Byte
'Public DueñoLindos As Byte
'Public DueñoQuest As Byte
'Public DueñoArghal As Byte
'Public DueñoDescanso As Byte
'Public DueñoLaurana As Byte
'pluto:2.4.5
Public MinutosOnline As Long
'pluto:2.9.0
Public Balon As Integer
Public Alarma As Byte
'pluto:2.14
Public DobleExp As Byte
Public HeroeExp As String

'pluto:2.15
Public MediaUsers As Integer
Public AyerMediaUsers As Integer
Public MediaVez As Integer
Public MediaUser As Integer
Public Caballero As Boolean
Public AyerReNumUsers As Integer
Public HoraHoy As String
Public Horayer As String

Public MixedKey As Long
Public ServerIp As String
Public CrcSubKey As String
Public castillo1 As String
Public castillo2 As String
Public castillo3 As String
Public castillo4 As String
Public fortaleza As String

Public hora1 As String
Public hora2 As String
Public hora3 As String
Public hora4 As String
Public hora5 As String
Public date1 As String
Public date2 As String
Public date3 As String
Public date4 As String
Public date5 As String
'pluto:2.9.0
Public PartidoFutbol As Byte
'pluto:hoy
Public UltimoBan As String
'pluto:2-3-04
Public Cotilla As String
Public Tesoromomia As Integer
Public Tesorocaballero As Integer
'pluto:2.3
Public SoloGm As Byte
'pluto:2.15
Public MapaSeguro As Integer
Public MapaAngel As Integer
'pluto:2.14
Public BodyTorneo As Integer

Type tEstadisticasDiarias
    segundos As Double
    Maxusuarios As Integer
    Promedio As Integer
End Type
    
Public DayStats As tEstadisticasDiarias

'pluto:2.12
Public TorneoBote As Long
Public Torneo2Record As Byte
Public Torneo2Name As String
'pluto:2.9.0
Type Torneito
Ttip As Byte
Tcua As Byte
Tpj As Byte
Tmax As Byte
Tmin As Byte
Tins As Long
Creador As String
FaseTorneo As Byte
Participantes(1 To 8) As String
End Type

Public TorneoPluto As Torneito
Public Futboleros(1 To 6) As String

Public aDos As New clsAntiDoS
Public aClon As New clsAntiMassClon
Public TrashCollector As New Collection

Public Const MAXSPAWNATTEMPS = 60
Public Const MAXUSERMATADOS = 9000000
Public Const LoopAdEternum = 999
Public Const FXSANGRE = 14
Public Const MAPATORNEO = 164
'pluto:2.12
Public Const MapaTorneo2 = 194
Public Const MapaAnteTorneo2 = 193
'pluto:2.15
Public WeB As String
Public ActualizaWeb As Byte
Public DifServer As Byte
Public DifOro As Byte
Public BaseDatos As Byte
'pluto:2.24---------------------------
Public ServerPrimario As Byte
Public NumeroObjEvento As Integer
Public CantEntregarObjEvento As Integer
Public CantObjRecompensa As Integer
Public ObjRecompensaEventos(1 To 4) As Integer
'-------------------------------------
Public Const iFragataFantasmal = 87
Public Const pluto1 = 2
Public Const pluto2 = 1
Public Const pluto3 = 10
Public Type tLlamadaGM
    Usuario As String * 255
    Desc As String * 255
End Type
'nivel al que deja de ser nw
Public Const LimiteNewbie = 14

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Const NingunEscudo = 2
Public Const NingunCasco = 2
'[GAU]
Public Const NingunBota = 2
'[GAU]
Public Const EspadaMataDragonesIndex = 402

Public Const MAXMASCOTASENTRENADOR = 7

Public Const FXWARP = 1
Public Const FXCURAR = 2

Public Const FXMEDITARCHICO = 4
Public Const FXMEDITARMEDIANO = 5
Public Const FXMEDITARGRANDE = 6
Public Const FXMEDITARRAYOS = 16
Public Const FXMEDITARRAYOSFUEGO = 17
Public Const FXMEDITARorbitalazul = 18
Public Const FXMEDITARorbitalrojo = 21
Public Const POSINVALIDA = 8

Public Const Bosque = "BOSQUE"
Public Const Nieve = "NIEVE"
Public Const Desierto = "DESIERTO"

Public Const Ciudad = "CIUDAD"
Public Const Campo = "CAMPO"
Public Const Dungeon = "DUNGEON"
Public Const Casa = "CASA"
Public Const BOSQUETERROR = "BOSQUE TERROR"
'PLUTO:2-3-04
Public Const ALCANTARILLA = "ALCANTARILLA"


' <<<<<< Targets >>>>>>
Public Const uUsuarios = 1
Public Const uNPC = 2
Public Const uUsuariosYnpc = 3
Public Const uTerreno = 4

' <<<<<< Acciona sobre >>>>>>
Public Const uPropiedades = 1
Public Const uEstado = 2
Public Const uInvocacion = 4
Public Const uMaterializa = 3

Public Const DRAGON = 6
Public Const MataDragones = 1

Public Const MAX_MENSAJES_FORO = 35

Public Const MAXUSERHECHIZOS = 50


Public Const EsfuerzoTalarGeneral = 4
Public Const EsfuerzoTalarLeñador = 2

Public Const EsfuerzoPescarPescador = 1
Public Const EsfuerzoPescarGeneral = 3

Public Const EsfuerzoExcavarMinero = 2
Public Const EsfuerzoExcavarGeneral = 5


Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

Public Const Guardias = 6

Public Const MAXREP = 6000000
Public Const MAXORO = 999999999
Public Const MAXEXP = 999999999

Public Const MAXATRIBUTOS = 35
Public Const MINATRIBUTOS = 6

Public Const LingoteHierro = 386
Public Const LingotePlata = 387
Public Const LingoteOro = 388
Public Const Leña = 58
'[MerLiNz:6]
Public Const GemaI = 598
Public Const Diamante = 695
'[\END]

Public Const MAXNPCS = 10000
Public Const MAXCHARS = 10000

Public Const HACHA_LEÑADOR = 127
Public Const PIQUETE_MINERO = 187

Public Const DAGA = 15
Public Const FOGATA_APAG = 136
Public Const FOGATA = 63
Public Const ORO_MINA = 194
Public Const PLATA_MINA = 193
Public Const HIERRO_MINA = 192
Public Const MARTILLO_HERRERO = 389
Public Const SERRUCHO_CARPINTERO = 198
'[MerLiNz:6]
Public Const SERRUCHOMAGICO_ermitano = 837
'[\END]
Public Const ObjArboles = 4

Public Const NPCTYPE_COMUN = 0
Public Const NPCTYPE_REVIVIR = 1
Public Const NPCTYPE_GUARDIAS = 2
Public Const NPCTYPE_ENTRENADOR = 3
Public Const NPCTYPE_BANQUERO = 4
Public Const NPCTYPE_EXP = 7
Public Const NPCTYPE_CASINO = 8
Public Const NPCTYPE_CIRUJANO = 9
Public Const NPCTYPE_TORNEO = 10
Public Const NPCTYPE_GUARDIAS2 = 11
Public Const NPCTYPE_CHISMOSO = 12
Public Const FX_TELEPORT_INDEX = 1
'pluto:6.9
Public Const NPCTYPE_VIAJERO = 43

Public Const Criatura_Suerte = 585
Public Const mapa_castillo1 = 166
Public Const mapa_castillo2 = 167
Public Const mapa_castillo3 = 168
Public Const mapa_castillo4 = 169
Public Const Criatura_1 = 604
Public Const Criatura_2 = 605
Public Const Criatura_3 = 606
Public Const mapi = 165
Public Const mapix1 = 18
Public Const mapix2 = 24
Public Const mapix3 = 24
Public Const mapix4 = 29
Public Const mapiy1 = 20
Public Const mapiy2 = 16
Public Const mapiy3 = 25
Public Const mapiy4 = 20
Public Const MIN_APUÑALAR = 10

'********** CONSTANTANTES ***********
Public Const NUMSKILLS = 31
Public Const NUMATRIBUTOS = 5
'[MerLiNz:X]
Public Const NUMCLASES = 21
'[\END]
Public Const NUMRAZAS = 9

Public Const MAXSKILLPOINTS = 200

Public Const FLAGORO = 777

Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4

'PLUTO:2.3
Public Const MAXMONTURA = 12
Public Const MAXMASCOTAS = 3
'[Tite]Party
Public Const MAXPARTYS = 30
Public Const MAXMIEMBROS = 10
'[\Tite]

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO = 100
Public Const vlASESINO = 1000
Public Const vlCAZADOR = 5
Public Const vlNoble = 5
Public Const vlLadron = 25
Public Const vlProleta = 2



'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto = 8
Public Const iCuerpoMuerto2 = 145
Public Const iCabezaMuerto = 500
Public Const iCabezaMuerto2 = 499

Public Const iORO = 12
Public Const Pescado = 139
Public Const Pescado2 = 544
Public Const Pescado3 = 545

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const suerte = 1
Public Const Magia = 2
Public Const Robar = 3
Public Const Tacticas = 4
Public Const Armas = 5
Public Const Meditar = 6
Public Const Apuñalar = 7
Public Const Ocultarse = 8
Public Const Supervivencia = 9
Public Const Talar = 10
Public Const Comerciar = 11
Public Const Defensa = 12 'escudos
Public Const Pesca = 13
Public Const Mineria = 14
Public Const Carpinteria = 15
Public Const Herreria = 16
Public Const Liderazgo = 17
Public Const Domar = 18
Public Const Proyectiles = 19 'Acertar Proyec.
Public Const Navegacion = 21

'pluto:2.15
Public Const DobleArma = 20 'Posibilidad de golpear con la segunda



'Requerido es Magia (2)

'no habilitados------------
Public Const DanoArma = 25 'vale para dos manos también
Public Const DefArma = 26    'vale para dos manos también
Public Const DañoProyec = 28
Public Const DefProyec = 29
Public Const DañoMagia = 22
Public Const DefMagia = 23
Public Const EvitaMagia = 24
'------------------------------
Public Const RequeArma = 27 ' vale para dos manos.
'acertar es Armas (5)
'evitar es tactica (4)

Public Const RequeProyec = 30
Public Const EvitarProyec = 31
'acertar es Proyectiles (19)

'pluto:2.17
Public MinutosCastilloNorte As Long
Public MinutosCastilloSur As Long
Public MinutosCastilloEste As Long
Public MinutosCastilloOeste As Long
Public MinutosFortaleza As Long

'-----------------------------
Public Const FundirMetal = 88

Public Const XA = 40
Public Const XD = 10
Public Const Balance = 9

Public Const Fuerza = 1
Public Const Agilidad = 2
Public Const Inteligencia = 3
Public Const Carisma = 4
Public Const Constitucion = 5


Public Const AdicionalHPGuerrero = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalSTLadron = 3

Public Const AdicionalSTLeñador = 23
Public Const AdicionalSTPescador = 20
Public Const AdicionalSTMinero = 25

'Tamaño del mapa
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'Tamaño del tileset
Public Const TileSizeX = 32
Public Const TileSizeY = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow = 17
Public Const YWindow = 13

'Sonidos
Public Const SOUND_BUMP = 1
Public Const SOUND_SWING = 2
Public Const SOUND_TALAR = 13
Public Const SOUND_PESCAR = 14
Public Const SOUND_MINERO = 15
Public Const SND_WARP = 3
Public Const SND_PUERTA = 5
Public Const SOUND_NIVEL = 6
Public Const SOUND_COMIDA = 7
Public Const SOUND_tele = 100
Public Const SOUND_resu = 101
Public Const SOUND_sana = 102
Public Const SOUND_para = 103
Public Const SND_USERMUERTE = 11
Public Const SND_IMPACTO = 10
Public Const SND_IMPACTO2 = 12
Public Const SND_LEÑADOR = 13
Public Const SND_FOGATA = 14
Public Const SND_AVE = 21
Public Const SND_AVE2 = 22
Public Const SND_AVE3 = 34
Public Const SND_GRILLO = 28
Public Const SND_GRILLO2 = 29
Public Const SOUND_SACARARMA = 25
Public Const SND_ESCUDO = 37
Public Const MARTILLOHERRERO = 41
Public Const LABUROCARPINTERO = 42
Public Const SND_CREACIONCLAN = 44
Public Const SND_ACEPTADOCLAN = 43
Public Const SND_DECLAREWAR = 45
Public Const SND_BEBER = 46
Public Const SND_DINERO = 104
Public Const SND_TORNEO = 105
Public Const SND_CASA1 = 110
Public Const SND_CASA2 = 111
Public Const SND_CASA3 = 112
Public Const SND_CASA4 = 113
Public Const SND_CASA5 = 114

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 25

'<------------------CATEGORIAS PRINCIPALES--------->
Public Const OBJTYPE_USEONCE = 1
Public Const OBJTYPE_WEAPON = 2
Public Const OBJTYPE_ARMOUR = 3
Public Const OBJTYPE_ARBOLES = 4
Public Const OBJTYPE_GUITA = 5
Public Const OBJTYPE_PUERTAS = 6
Public Const OBJTYPE_CONTENEDORES = 7
Public Const OBJTYPE_CARTELES = 8
Public Const OBJTYPE_LLAVES = 9
Public Const OBJTYPE_FOROS = 10
Public Const OBJTYPE_POCIONES = 11
Public Const OBJTYPE_BEBIDA = 13
Public Const OBJTYPE_LEÑA = 14
Public Const OBJTYPE_FOGATA = 15
Public Const OBJTYPE_HERRAMIENTAS = 18
Public Const OBJTYPE_YACIMIENTO = 22
Public Const OBJTYPE_PERGAMINOS = 24
Public Const OBJTYPE_teleport = 19
Public Const OBJTYPE_YUNQUE = 27
Public Const OBJTYPE_FRAGUA = 28
Public Const OBJTYPE_MINERALES = 23
Public Const OBJTYPE_CUALQUIERA = 1000
Public Const OBJTYPE_INSTRUMENTOS = 26
Public Const OBJTYPE_BARCOS = 31
Public Const OBJTYPE_FLECHAS = 32
Public Const OBJTYPE_BOTELLAVACIA = 33
Public Const OBJTYPE_BOTELLALLENA = 34
Public Const OBJTYPE_MANCHAS = 35
Public Const OBJTYPE_tele = 36
Public Const OBJTYPE_resu = 37
Public Const OBJTYPE_sana = 38
Public Const OBJTYPE_para = 39
Public Const OBJTYPE_regalo = 40
'pluto:2.4
Public Const OBJTYPE_Anillo = 41

Public Const OBJTYPE_Montura = 60
'<------------------SUB-CATEGORIAS----------------->
Public Const OBJTYPE_ARMADURA = 0
Public Const OBJTYPE_CASCO = 1
Public Const OBJTYPE_ESCUDO = 2
Public Const OBJTYPE_CAÑA = 138
'[GAU]
Public Const OBJTYPE_BOTA = 3
'[GAU]


'Tipo de posicones
'1 Modifica la Agilidad
'2 Modifica la Fuerza
'3 Repone HP
'4 Repone Mana
Public Enum FontTypeNames
    FONTTYPE_talk
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_info
    FONTTYPE_INFObold
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO2
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_guildmsg
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesa
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_VENENO
    FONTTYPE_pluto
    FONTTYPE_COMERCIO
    FONTTYPE_GLOBAL
    FONTTYPE_CAOS
    FONTTYPE_ARMADA
End Enum
'Texto
'Public Const FONTTYPENAMES.FONTTYPE_TALK = "~255~255~255~0~0"
'Public Const FONTTYPENAMES.FONTTYPE_fight = "~255~0~0~1~0"
'Public Const FONTTYPENAMES.FONTTYPE_WARNING = "~32~51~223~1~1"
'Public Const FONTTYPENAMES.FONTTYPE_INFO = "~65~190~156~0~0"
'Public Const FONTTYPENAMES.FONTTYPE_VENENO = "~0~255~0~0~0"
'Public Const FONTTYPENAMES.FONTTYPE_GUILD = "~255~255~255~1~0"
'Public Const FONTTYPENAMES.FONTTYPE_PLUTO = "~255~150~0~1~0"
'pluto:2.8.0
'Public Const FONTTYPENAMES.FONTTYPE_COMERCIO = "~221~216~9~1~0"

'Estadisticas
Public Const STAT_MAXELV = 99
Public Const STAT_MAXHP = 999
Public Const STAT_MAXSTA = 999
Public Const STAT_MAXMAN = 2000
Public Const STAT_MAXHIT = 99
Public Const STAT_MAXDEF = 99

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1

Public Const SND_NODEFAULT = &H2

Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10


'**************************************************************
'**************************************************************
'************************ TIPOS *******************************
'**************************************************************
'**************************************************************


'pluto:6.0A
Type PMascotas
Tipo As String
AumentoCuerpo As Byte
AumentoMagia As Byte
ReduceCuerpo As Byte
ReduceMagia As Byte
AumentoFlecha As Byte
ReduceFlecha As Byte
AumentoEvasion As Byte
TopeLevel As Byte
VidaporLevel As Integer
GolpeporLevel As Integer
exp(1 To 30) As Long
TopeAtMagico As Byte
TopeDefMagico As Byte
TopeAtFlechas As Byte
TopeDefFlechas As Byte
TopeAtCuerpo As Byte
TopeDefCuerpo As Byte
TopeEvasion As Byte
End Type

Type tHechizo
    Nombre As String
    Desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    Resis As Byte
    
    Tipo As Byte
    WAV As Integer
    FXgrh As Integer
    loops As Byte
    
    SubeHP As Byte
    MinHP As Integer
    MaxHP As Integer
    
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    Invisibilidad As Byte
    Paraliza As Byte
    Paralizaarea As Byte
    RemoverParalisis As Byte
    CuraVeneno As Byte
    Envenena As Byte
    'pluto:2.15
    Protec As Byte
    Ron As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    
    invoca As Byte
    NumNpc As Integer
    Cant As Integer
    
    MinNivel As Byte
    itemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    Target As Byte
End Type

Type LevelSkill

LevelValue As Integer

End Type

Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    HerramientaEqpObjIndex As Integer
    HerramientaEqpSlot As Integer
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    NroItems As Integer
  
    'pluto:2.4
  AnilloEqpObjIndex As Integer
  AnilloEqpSlot As Byte
  
  '[GAU]
    BotaEqpObjIndex As Integer
    BotaEqpSlot As Byte
    '[GAU]
End Type


Type Position
    X As Integer
    Y As Integer
End Type

Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Type FXdata
    Nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

'Datos de user o npc
Type Char
    CharIndex As Integer
    Head As Integer
    Body As Integer
    '[GAU]
    Botas As Integer
    '[GAU]
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    loops As Integer
    
    Heading As Byte

End Type

'Tipos de objetos
Public Type ObjData
    'pluto:6.0A
    ParaCarpin As Byte
    ParaErmi As Byte
    ParaHerre As Byte
    
    ArmaNpc As Integer
    Name As String 'Nombre del obj
    'pluto:2.8.0
    Vendible As Integer
    
    OBJType As Integer 'Tipo enum que determina cuales son las caract del obj
    SubTipo As Integer 'Tipo enum que determina cuales son las caract del obj
    'pluto:7.0
    Drop As Byte
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    'pluto:2.3
    Peso As Double
    
    Respawn As Byte
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    
    MineralIndex As Integer
    'LingoteInex As Integer
    
    
    proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    Newbie As Integer
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    
    Def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    'pluto:7.0
    Defmagica As Integer
    'nati: agrego defcuerpo
    Defcuerpo As Integer
    'Defproyectil As Integer
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
     '[GAU]
    Botas As Integer
    '[GAU]
    Valor As Long     ' Precio
    objetoespecial As Integer
    Cerrada As Integer
    Llave As Byte
    Clave As Long 'si clave=llave la puerta se abre o cierra
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    Mujer As Byte
    Hombre As Byte
    Envenena As Byte
    Magia As Byte
    Resistencia As Long
    Agarrable As Byte
    
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    '[MerLiNz:6]
    Gemas As Integer
    Diamantes As Integer
    'pluto:2.10
    ObjetoClan As String
    '[\END]
    
    SkHerreria As Byte
    SkCarpinteria As Byte
    
    texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As String
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    MinInt As Integer
    
    Real As Byte
    Caos As Byte
    nocaer As Byte
    razaelfa As Byte
    razavampiro As Byte
    razahumana As Byte
    razaorca As Byte
    SkArco As Byte
    SkArma As Byte
    Cregalos As Integer
    Pregalo As Byte
    End Type


Public Type obj
    ObjIndex As Integer
    Amount As Integer
'nati: agrego esto para el modulo de quest, para la entrega de items altos o bajos.
    ObjIndex2 As Integer
    Amount2 As Integer
    Separador As Integer
End Type

'pluto:6.0A
Public ObjRegalo1(200) As Integer
Public ObjRegalo2(200) As Integer
Public ObjRegalo3(200) As Integer
Public Reo1 As Integer
Public Reo2 As Integer
Public Reo3 As Integer
'Banco Objs pluto:7.0
Public Const MAX_BANCOINVENTORY_SLOTS = 20
'pluto:6.0A
Public Const MAX_BOVEDACLAN_SLOTS = 40
'----------

'[KEVIN]
Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    'NroItems As Integer
End Type
'[/KEVIN]

'Delzak sistema premios

Type Premios
    
    MataArañas As Integer
    MataNoMuertos As Integer
    MataAnimales As Integer
    MataNavidad As Integer
    MataDarks As Integer
    MataGoblins As Integer
    MataDragones As Integer
    MataOgros As Integer
    MataOrcos As Integer
    MataHechiceros As Integer
    MataPuertas As Integer
    MataReyes As Integer
    MataDefensores As Integer
    MataEttins As Integer
    MataDemonios As Integer
    MataBeholders As Integer
    MataMarinos As Integer
    MataMedusas As Integer
    MataCiclopes As Integer
    MataHobbits As Integer
    MataGenios As Integer
    MataPolares As Integer
    MataGollums As Integer
    MataDevirs As Integer
    MataUruks As Integer
    MataDevastadores As Integer
    MataEnts As Integer
    MataLicantropos As Integer
    MataPiratas As Integer
    MataLagartos As Integer
    MataTrolls As Integer
    MataGolems As Integer
    MataRaids As Integer
    
End Type


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************


Type tReputacion 'Fama del usuario
    NobleRep As Double
    BurguesRep As Double
    PlebeRep As Double
    LadronesRep As Double
    BandidoRep As Double
    AsesinoRep As Double
    Promedio As Double
    
End Type



'Estadisticas de los usuarios
Type UserStats
     
      'pluto:2.4
    Fama As Long
    GTorneo As Long
    PClan As Long
     
     'pluto:2.9.0
    Torneo1 As Integer
    Torneo2 As Integer
    
    'pluto:2-3-04
    Puntos As Long
    GLD As Long 'Dinero
    Banco As Long
    'Delzak
    'Remort As Boolean
    MET As Integer
    
    MaxHP As Integer
    MinHP As Integer
    'pluto:2.3
    Peso As Single
    PesoMax As Integer
    
    FIT As Integer
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    
    MaxHam As Integer
    MinHam As Integer
    
    MaxAGU As Integer
    MinAGU As Integer
        
    Def As Integer
    exp As Long
    ELV As Byte
    Elu As Long
    LibrosUsados As Integer
    PremioNPC(1 To 34) As Integer  'Delzak sistema premios
    
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Integer
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Integer
    CriminalesMatados As Integer
    NPCsMuertos As Integer
    SkillPts As Integer
    
End Type

'pluto:hoy

'pluto:7.0----------------

Type Usermision
SoloClase As String
TimeComienzo As String
TimeMision As Integer
Titulo As String
tX As String
NivelMinimo As Byte
NivelMaximo As Byte
NEnemigos As Byte
Cargada As Boolean
Enemigo(1 To 5) As Integer
EnemigoCantidad(1 To 5) As Byte
NEnemigosConseguidos(1 To 5) As Byte
PjConseguidos As Byte
NObjetos As Byte
Objeto(1 To 5) As String
Level As Integer
Entrega As Integer
exp As Long
oro As Long
NObjetosR As Byte
ObjetoR(1 To 5) As String
Actual As Integer
Actual1 As Integer
Actual2 As Integer
Actual3 As Integer
Actual4 As Integer
Actual5 As Integer
Actual6 As Integer
Actual7 As Integer
Actual8 As Integer
Actual9 As Integer
Actual10 As Integer
Actual11 As Integer
Actual12 As Integer
NpcQuest As Integer
'-------------------------
estado As Integer
numero As Integer
'Enemigo As Integer
'Objeto As Integer
Cantidad As Integer
'Entrega As Integer
'Level As Integer
clase As String
End Type

'PLUTO:6.0A
Type UserMONTURA
Nivel(1 To MAXMONTURA) As Integer
exp(1 To MAXMONTURA) As Long
Elu(1 To MAXMONTURA) As Long
Vida(1 To MAXMONTURA) As Integer
Golpe(1 To MAXMONTURA) As Integer
Nombre(1 To MAXMONTURA) As String
AtMagico(1 To MAXMONTURA) As Byte
DefMagico(1 To MAXMONTURA) As Byte
AtCuerpo(1 To MAXMONTURA) As Byte
Defcuerpo(1 To MAXMONTURA) As Byte
AtFlechas(1 To MAXMONTURA) As Byte
DefFlechas(1 To MAXMONTURA) As Byte
Evasion(1 To MAXMONTURA) As Byte
Libres(1 To MAXMONTURA) As Byte
Tipo(1 To MAXMONTURA) As Byte
index(1 To MAXMONTURA) As Byte
End Type


'Flags
Type UserFlags

    'pluto:6.8
    Intentos As Byte
    NoTorneos As Boolean
    MapaIncor As Integer
    Incor As Boolean
    
    Pitag As Byte
    Arqui As Byte
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Protec As Byte
    Ron As Byte
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    ModoCombate As Boolean
    'pluto:5.2
    CMuerte As Byte
'[Tite]flag seguro golpe critico
    SegCritico As Boolean
'[\Tite]
'[Tite]flags party
    partyNum As Integer
    party As Boolean
    invitado As String
    privado As Byte
'[\Tite]
    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeAtacar As Byte
    PuedeMoverse As Byte
    PuedeLanzarSpell As Byte
    'pluto:2.8.0
    PuedeFlechas As Byte
    'pluto:2.10
    PuedeTomar As Byte
    ClienteValido As Integer
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    Angel As Integer
    Demonio As Integer
    Morph As Integer
    Vuela As Byte
    Navegando As Byte
    Seguro As Boolean
    'pluto:2.3
    Montura As Integer
    ClaseMontura As Integer
    'pluto:6.2
    Macreanda As Byte
    ComproMacro As Byte
    ParejaTorneo As Integer
    
    DuracionEfecto As Long
    TargetNpc As Integer ' Npc señalado por el usuario
    TargetNpcTipo As Integer ' Tipo del npc señalado
    NpcInv As Integer
    
    ban As Byte
    AdministrativeBan As Byte
    
    TargetUser As Integer ' Usuario señalado
    
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    
    StatsChanged As Byte
    Privilegios As Byte
 
    'pluto:6.0a----------------
    MinutosOnline As Long
    Minotauro As Byte
    Creditos As Integer
    DragCredito1 As Byte
    DragCredito2 As Byte
    DragCredito3 As Byte
    DragCredito4 As Byte
    DragCredito5 As Byte
    DragCredito6 As Byte
    'pluto:7.0
  
    NCaja As Byte
    Elixir As Byte
    '--------------------------
    
    
    ValCoDe As Integer
    
    LastCrimMatado As String
    LastCrimMatado2 As String
    LastCiudMatado As String
    LastCiudMatado2 As String
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    Torneo As Integer
'pluto:2.9.0
TorneoPluto As Byte
Futbol As Byte

End Type

Type UserCounters
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    COMCounter As Integer
    AGUACounter As Integer
    veneno As Integer
    Paralisis As Integer
    Morph As Integer
    Angel As Integer
    Ceguera As Integer
    Estupidez As Integer
    Invisibilidad As Integer
    PiqueteC As Long
    bloqueo As Long
    Pena As Long
    Incor As Long
    Macrear As Integer
    'pluto:2.23
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerTomar As Long
    TimerUsarArco As Long
    '-----------------------
    SendMapCounter As WorldPos
    Pasos As Integer
 Saliendo As Boolean
    Salir As Integer
    Protec As Integer
    Ron As Integer
    'pluto:6.7
    'UserEnvia As Byte
    'UserRecibe As Byte
End Type

Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Double
    CiudadanosMatados As Double
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    RecibioArmaduraLegion As Byte
End Type

Type tGuild
    GuildName As String
    Solicitudes As Long
    SolicitudesRechazadas As Long
    Echadas As Long
    VecesFueGuildLeader As Long
    YaVoto As Byte
    EsGuildLeader As Byte
    FundoClan As Byte
    ClanFundado As String
    ClanesParticipo As Long
    GuildPoints As Double
End Type
    


'Tipo de los Usuarios
Type User
    'pluto:7.0
    BonusElfoOscuro As Integer
    'pluto:6.8
    PoSum As WorldPos
    'pluto:2.12
    Torneo2 As Byte
    'pluto:6.0A
    Nmonturas As Byte
    'pluto:2.9.0
    ObjetosTirados As Integer
    Alarma As Integer
    Name As String
    ID As Long
    'pluto:2.10
    GranPoder As Byte
    'pluto:2.14
    Chetoso As Byte
    'pluto:2.18------------------
    UserDañoMagiasRaza As Byte
    UserDañoArmasRaza As Byte
    UserDañoProyetilesRaza As Byte
    UserEvasiónRaza As Byte
    UserDefensaMagiasRaza As Byte
    UserDefensaEscudos As Byte
    '-------------------------
    modName As String
    Password As String
    
    Char As Char 'Define la apariencia
    OrigChar As Char
    Remort As Integer
    Remorted As String
    'pluto:6.0A
    BD As Byte
 
    Desc As String ' Descripcion
    clase As String
    raza As String
    Genero As String
    Email As String
    'pluto:2.10
    EmailActual As String
     Hogar As String
    'pluto:2.14-------
    Serie As String
    Paquete As String
    MacPluto As String
    MacPluto2 As String
    MacClave As Integer
    Nhijos As Integer
    Padre As String
    Madre As String
    Hijo(1 To 5) As String
   
    Esposa As String
    Amor As Byte
    Embarazada As Integer
    Bebe As Byte
    NombreDelBebe As String
    '-----------------
    
    Invent As Inventario
    
    Pos As WorldPos
    
    ConnIDValida As Boolean
    ConnID As Integer 'ID
    RDBuffer As String 'Buffer roto
    
    CommandsBuffer As New CColaArray
     ColaSalida As New Collection
    SockPuedoEnviar As Boolean
    
    '[KEVIN]
    BancoInvent(1 To 6) As BancoInventario
    '[/KEVIN]
    
    
    Counters As UserCounters
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMacotas As Integer
    
    'pluto:hoy
    Mision As Usermision

    'pluto:2.3
    Montura As UserMONTURA
    
    Stats As UserStats
    flags As UserFlags
    NumeroPaquetesPorMiliSec As Long
    BytesTransmitidosUser As Long
    BytesTransmitidosSvr As Long
    
    'pluto:2.4.5
    ShTime As Integer
   'pluto.2.5.0
    MuertesTime As Integer
    
    Reputacion As tReputacion
    
    Faccion As tFacciones
    GuildInfo As tGuild
    GuildRef  As cGuild
    
    PrevCRC As Long
    PacketNumber As Long
    RandKey As Long
    
    ip As String

     '[Alejo]
    ComUsu As tCOmercioUsuario
    '[/Alejo]
End Type

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S  D E  P A R T Y  **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'[Tite] Party
Type clMiembro
ID As Integer
privi As Byte
End Type
Type ClParty
    privada As Byte
    lider As Integer
    expAc As Long
    miembros(1 To MAXMIEMBROS) As clMiembro
    Solicitudes(1 To MAXMIEMBROS) As Integer
    numMiembros As Byte
    numSolicitudes As Byte
    reparto As Byte
End Type
'[\Tite]


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Type NPCStats
    Alineacion As Integer
    MaxHP As Long
    MinHP As Long
    MaxHIT As Integer
    MinHIT As Integer
    Def As Integer
    UsuariosMatados As Integer
    ImpactRate As Integer
End Type

Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
    
End Type

Type NPCFlags
    'pluto:2.14
    PoderEspecial1 As Byte
    PoderEspecial2 As Byte
    PoderEspecial3 As Byte
    PoderEspecial4 As Byte
    PoderEspecial5 As Byte
    PoderEspecial6 As Byte
    
    AfectaParalisis As Byte
    Magiainvisible As Byte
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    LanzaSpells As Byte
    
    OldMovement As Byte
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    UseAINow As Boolean
    Sound As Integer
    Attacking As Integer
    AttackedBy As String
    Category1 As String
    Category2 As String
    Category3 As String
    Category4 As String
    Category5 As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    Snd4 As Integer
    
End Type

Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

'<--------- New type for holding the pathfinding info ------>
Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type
'<--------- New type for holding the pathfinding info ------>


Type npc
    Name As String
    Char As Char 'Define como se vera
    Desc As String
    
    NPCtype As Integer
    'Premio As Integer 'Delzak sistema premios
    numero As Integer
    Anima As Byte
    'pluto:6.0A
    Arquero As Byte
    Level As Integer
    Raid As Byte
    'pluto:7.0
    LogroTipo As Byte
    
    InvReSpawn As Byte
    
    Comercia As Integer
    Target As Long
    TargetNpc As Long
    TipoItems As Integer
    
    veneno As Byte
    
    Pos As WorldPos 'Posicion
    Orig As WorldPos
    SkillDomar As Integer
    
    Movement As Integer
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long
    clan As String
    Inflacion As Long
    
    GiveEXP As Long
    GiveGLD As Long
    
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
    
    '<---------New!! Needed for pathfindig----------->
    PFINFO As NpcPathFindingInfo

End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    UserIndex As Integer
    NpcIndex As Integer
    OBJInfo As obj
    TileExit As WorldPos
    trigger As Integer
End Type

'Info del mapa
Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Pk As Boolean
    invocado As Integer
    'pluto:2.15
    Dueño As Byte
    Aldea As Byte
    
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    'pluto:6.0A
    Mascotas As Byte
    Invisible As Byte
    Resucitar As Byte
    Insegura As Byte
    Lluvia As Byte
    Monturas As Byte
    Domar As Byte
'pluto: 2.9.0
'Objetos As Integer
End Type



'********** V A R I A B L E S     P U B L I C A S ***********

'[Tite] Party
Public numPartys As Byte
'[\Tite]
'pluto:hoy
Public ResTrivial As String
Public PreTrivial As String
Public ResEgipto As Integer
Public PreEgipto As String
'pluto:2.4
Public NivCrimi As Integer
Public NivCiu As Integer
Public NivCrimiON As Integer
Public NivCiuON As Integer
Public NNivCrimi As String
Public NNivCiu As String
Public NNivCrimiON As String
Public NNivCiuON As String
Public UserCiu As Integer
Public UserCrimi As Integer
Public MaxTorneo As Integer
Public Moro As Long
Public MoroOn As Long
Public NMaxTorneo As String
'pluto:6.9
Public Clan1Torneo As String
Public Clan2Torneo As String
Public PClan1Torneo As Long
Public PClan2Torneo As Long
Public NomClan() As String
Public PuntClan() As Integer


Public NMoro As String
Public NMoroOn As String

Public SERVERONLINE As Boolean
Public ULTIMAVERSION As String
Public BackUp As Boolean

Public ListaRazas() As String
Public SkillsNames() As String
Public ListaClases() As String


Public ENDL As String
Public ENDC As String

Public recordusuarios As Long

'Directorios
Public IniPath As String
Public CharPath As String
Public MapPath As String
Public DatPath As String

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public ResPos As WorldPos
Public StartPos As WorldPos 'Posicion de comienzo

'pluto:2.8.0
Public ReNumUsers As Integer
Public NumUsers As Integer 'Numero de usuarios actual
Public LastUser As Integer
Public LastChar As Integer
'Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public HideMe As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean
Public haciendoBKPJ As Boolean
Public Oscuridad As Integer
Public NocheDia As Integer
'pluto:2.12
Public MinutoSinMorir As Byte
'pluto:2.15
'Public yaya As Byte
'*****************ARRAYS PUBLICOS*************************
'pluto:2.17
Public PMascotas(1 To MAXMONTURA) As PMascotas

Public UserList() As User 'USUARIOS
Public Npclist() As npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList() As Integer
Public ObjData() As ObjData
'pluto:6.0A--------------------
Public ObjetosClan(1 To 255) As OClan
Public NameClan(1 To 255) As String
Type OClan
ObjSlot(40) As obj
End Type
'------------------------------
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 70) As LevelSkill
Public ForbidenNames() As String
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
'[MerLiNz:6]
Public Objermitano() As Integer
'[\END]
'[Tite]Party
Public partylist(1 To MAXPARTYS) As ClParty
'[\Tite]
'*********************************************************

Public Nix As WorldPos
Public ciudadcaos As WorldPos
Public Ullathorpe As WorldPos
Public Banderbill As WorldPos
Public Lindos As WorldPos
Public Prision As WorldPos
Public Libertad As WorldPos
'pluto:2.17------------
Public Pobladohumano As WorldPos
Public Pobladoorco As WorldPos
Public Pobladoelfo As WorldPos
Public Pobladoenano As WorldPos
Public Pobladovampiro As WorldPos
'pluto:2.9.0
Public MsgEntra As String
Public GolesLocal As Byte
Public GolesVisitante As Byte
Public Vezz As Byte
'pluto:2.10
Public GranPoder As Byte
Public UserGranPoder As String
'pluto:6.0A
Public NumeroGranPoder As Byte

Public Ayuda As New cCola

'pluto:2.23
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GenCrC Lib "crc" Alias "GenCrc" (ByVal CrcKey As Long, ByVal CrcString As String) As Long


Sub PlayWaveAPI(file As String)

On Error GoTo fallo
Dim rc As Integer

rc = sndPlaySound(file, SND_ASYNC)

Exit Sub
fallo:
Call LogError("PLAYWAVEAPI" & Err.number & " D: " & Err.Description)


End Sub

