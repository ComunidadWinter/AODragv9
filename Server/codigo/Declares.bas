Attribute VB_Name = "Declaraciones"
Option Explicit

''
' Modulo de declaraciones. Aca hay de todo.
'

Enum eMessages
    Lejos = 0 '¡Esta demasiado lejos!
    Muerto = 1 '¡Estas muerto!
    NoComerciar = 2 '¡No puedes comerciar con eso!
    NoCantidad = 3 '¡No tienes esa cantidad!
    NoMana = 4 '¡No tienes suficiente maná!
    NoEnergia = 5 '¡No tienes suficiente energia!
    ClaseNoUsa = 6 '¡Tu clase no puede usar este objeto!
    MapaReal = 7 'Mapa exclusivo para miembros del ejército Real
    MapaNewbie = 8 'Mapa exclusivo para newbies.
    MapaCaos = 9 'Mapa exclusivo para miembros del ejército Oscuro
    MapaFaccion = 10 'Solo se permite entrar al Mapa si eres miembro de alguna Facción
    Invocacion = 11 'Se ha invocado una criatura en la Sala de Invocaciones.
    Frio = 12 '¡¡Estas muriendo de frio, abrigate o moriras!!
    Murio = 13 '¡Has muerto!
    Quema = 14 '¡Te estas quemando!
    Mimetico = 15 'Recuperas tu apariencia normal.
    Paralizado = 16 '¡Estas paralizado!
    EstarSegura = 17 '¡¡Tienes que estar en zona segura!!
    DuelosClasicos = 18 '¡Estas en duelos clasicos!
    DuelosClasicosELO = 19 '¡Estas en duelos clasicos!
    SalaDuelos = 20 '¡Ya estas en la sala de duelos!
    MinLVLDuelos = 21 'Tu nivel debe de ser 15 o superior.
    invisible = 22 '¡Estas invisible!
    DuelosOcupado = 23 'El mapa de duelos está ocupado ahora mismo.
    Cegado = 24 '¡Estas cegado!
    Carcel = 25 '¡Estas en la carcel!
    Piquete = 26 '¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!
    NoTrabaja = 27 'Dejas de trabajar.
    SolicitudPropia = 28 '¡No puedes enviarte una solicitud a ti mismo!
End Enum

Type tSQL
    Driver As String
    Server As String
    Database As String
    Port As String
    Name As String
    Pass As String
    Modo As String
End Type

Public InfSQL As tSQL

Public UserTrabajando As Boolean
Public SkillTrabajando As Byte

'*************************
'INVOCACIONES
'*************************
Public Const MapInvocacion As Byte = 23
Public Const XInvocacion As Byte = 27
Public Const YInvocacion As Byte = 46
Public Const NPCInvocacion As Integer = 542
'*************************

'********************
'*Sistema de Torneos organizados*
Public Hay_Torneo As Boolean
Public MinLevel As Byte
Public MaxLevel As Byte
Public Cupos As Byte
Public AutoSum As Byte
Public Mapa As Integer
Public X As Byte
Public Y As Byte
Public Torneo_Inscriptos As Long
'********************************

Public aLimpiarMundo As clsLimpiarMundo

'Sistema AntiCheat LAC - Para evitar cortar intervalos
Public Lac_Camina As Long
Public Lac_Pociones As Long
Public Lac_Pegar As Long
Public Lac_Lanzar As Long
Public Lac_Usar As Long
Public Lac_Tirar As Long
 
Public Type TLac
    LCaminar As New Cls_InterGTC
    LPociones As New Cls_InterGTC
    LPegar As New Cls_InterGTC
    LUsar As New Cls_InterGTC
    LTirar As New Cls_InterGTC
    LLanzar As New Cls_InterGTC
End Type
'/Sistema AntiCheat LAC - Para evitar cortar intervalos

'Sistema de Chat Global
Public HayGlobal As Boolean

'**********Casa encantada********************
Public Const Casa1 As Byte = 10
Public Const CasaRayoX1 As Byte = 49
Public Const CasaRayoX2 As Byte = 50
Public Const CasaRayoY As Byte = 55
Public Const CasaSpiritsORO As Integer = 30000
Public Const FxCASA As Byte = 10
'********************************************

Public Const RANGO_VISION_X As Byte = 11
Public Const RANGO_VISION_Y As Byte = 9

'**************Castillo**********************
Public NPCReyCastle As Integer 'REY DEL CASTILLO
Public NPCDefensorFortaleza As Integer 'REY DEL CASTILLO
Public Const NUMCASTILLOS As Byte = 5

'<Edurne>
Public Type CastillosInfo
    nombre As String
    Ubicacion As Integer
    Atacando As Boolean
    Dueño As Integer
    PuertaDie As Boolean
    Mapa As Integer
    FechaHora As String
End Type
Public Castillos(1 To 5) As CastillosInfo       '1= Norte   2=Este  3=Sur  4=Oeste   5=Forta
'<Edurne>
'*******************************************

Public aClon As New clsAntiMassClon
Public TrashCollector As New Collection

Public Const FxGranPoder As Byte = 54

Public Const MAXSPAWNATTEMPS = 60
Public Const INFINITE_LOOPS As Integer = -1
Public Const FXSANGRE = 48
Public Const FXSANGREXXL = 50
''
' The color of chats over head of dead characters.
Public Const CHAT_COLOR_DEAD_CHAR As Long = &HC0C0C0

''
' The color of yells made by any kind of game administrator.
Public Const CHAT_COLOR_GM_YELL As Long = &HF82FF

''
' Coordinates for normal sounds (not 3D, like rain)
Public Const NO_3D_SOUND As Byte = 0

Public Const iFragataFantasmal As Byte = 87
Public Const iFragataReal As Byte = 190
Public Const iFragataCaos As Byte = 189
Public Const iBarca As Byte = 84
Public Const iGalera As Byte = 85
Public Const iGaleon As Byte = 86
Public Const iBarcaCiuda = 395
Public Const iBarcaPk = 396
Public Const iGaleraCiuda = 397
Public Const iGaleraPk = 398
Public Const iGaleonCiuda = 399
Public Const iGaleonPk = 400

Public Enum iMinerales
    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eClass
    Mage = 1       'Mago
    Cleric = 2    'Clérigo
    Warrior = 3   'Guerrero
    Assasin = 4   'Asesino
    Bard = 5      'Bardo
    Druid = 6     'Druida
    Paladin = 7   'Paladín
    Hunter = 8    'Cazador
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal
End Enum

Public Enum eRaza
    Humano = 1
    Elfo
    Drow
    Gnomo
    Enano
    Orco
    NoMuerto
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Const LimiteNewbie As Byte = 14

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

'Barrin 3/10/03
'Cambiado a 2 segundos el 30/11/07
Public Const TIEMPO_SEND_PING As Byte = 200

Public Const NingunEscudo As Byte = 2
Public Const NingunCasco As Byte = 2
Public Const NingunArma As Byte = 2

Public Const EspadaMataDragonesIndex As Integer = 402

Public Const MAXMASCOTASENTRENADOR As Byte = 7

Public Enum FXIDs
    FXWARP = 1
    FXMEDITARCHICO = 46
    FXMEDITARMEDIANO = 3
    FXMEDITARGRANDE = 49
    FXMEDITARXGRANDE = 40
    FXMEDITARXXGRANDECIU = 34
    FXMEDITARXXGRANDECRI = 35
End Enum

Public Const TIEMPO_CARCEL_PIQUETE As Long = 10

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'

'11/12/2018 Irongete: Cambio los triggers
Public Enum eTrigger
    Nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
    DUELO = 7
    AUTORESU = 8
    TRAMPA_1 = 9
    TRAMPA_2 = 10
    SALASANGRE = 11
End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6
    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3
End Enum

'TODO : Reemplazar por un enum
Public Const Bosque As String = "BOSQUE"
Public Const Nieve As String = "NIEVE"
Public Const Desierto As String = "DESIERTO"
Public Const Ciudad As String = "CIUDAD"
Public Const Campo As String = "CAMPO"
Public Const Dungeon As String = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Enum TargetType
    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3    'Nose usa
    uInvocacion = 4
    uArea = 5
End Enum

Public Const MAXUSERHECHIZOS As Byte = 18


' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarLeñador As Byte = 2
Public Const EsfuerzoPescarPescador As Byte = 1
Public Const EsfuerzoExcavarMinero As Byte = 2
Public Const FX_TELEPORT_INDEX As Integer = 1

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Const Guardias As Integer = 6

Public Const MAX_ORO_EDIT As Long = 5000000

Public Const TAG_USER_INVISIBLE As String = "[INVISIBLE]"

Public Const MAXREP As Long = 6000000
Public Const MAXORO As Long = 90000000
Public Const MAXEXP As Long = 99999999

Public Const MAXUSERMATADOS As Long = 65000

Public Const MAXATRIBUTOS As Byte = 55
Public Const MINATRIBUTOS As Byte = 6

Public Const LingoteHierro As Integer = 386
Public Const LingotePlata As Integer = 387
Public Const LingoteOro As Integer = 388
Public Const Leña As Integer = 58


Public Const MAXNPCS As Integer = 10000
Public Const MAXCHARS As Integer = 10000

Public Const HACHA_LEÑADOR As Byte = 127
Public Const PIQUETE_MINERO As Byte = 187

Public Const DAGA As Byte = 15
Public Const FOGATA_APAG As Byte = 136
Public Const FOGATA As Byte = 63
Public Const ORO_MINA As Byte = 194
Public Const PLATA_MINA As Byte = 193
Public Const HIERRO_MINA As Byte = 192
Public Const MARTILLO_HERRERO As Integer = 389
Public Const SERRUCHO_CARPINTERO As Byte = 198
Public Const ObjArboles As Byte = 4
Public Const RED_PESCA As Integer = 543
Public Const CAÑA_PESCA As Byte = 138
Public Const DAMAGE_PUÑAL    As Byte = 1
Public Const DAMAGE_NORMAL   As Byte = 2
Public Const DAMAGE_MAGIC    As Byte = 3

'08/11/2015 Irongete: Nuevos colores
Public Const COLOR_DAÑO As Byte = 5
Public Const COLOR_CURACION As Byte = 6
Public Const COLOR_ORO As Byte = 7
Public Const COLOR_MENSAJE As Byte = 8


Public Enum eNPCType
    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Noble = 5
    DRAGON = 6
    Timbero = 7
    Guardiascaos = 8
    ResucitadorNewbie = 9
    Pirata = 10
    MonsterDrag = 11
    Cirujano = 12
    Guardiafalso = 13
    Quest = 14
    Arbol = 15
    Yacimiento = 16
    Subastador = 17
End Enum

Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS As Byte = 17

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES As Byte = 8

''
' Cantidad de Razas
Public Const NUMRAZAS As Byte = 7


''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''
'Direccion
'
' @param NORTH Norte 2
' @param EAST Este 4
' @param SOUTH Sur 1
' @param WEST Oeste 3
'
Public Enum eHeading
    SOUTH = 1
    NORTH = 2
    WEST = 3
    EAST = 4
End Enum

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS As Byte = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO As Byte = 100
Public Const vlASESINO As Integer = 1000
Public Const vlCAZADOR As Byte = 5
Public Const vlNoble As Byte = 5
Public Const vlLadron As Byte = 25
Public Const vlProleta As Byte = 2

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto As Byte = 8
Public Const iCabezaMuerto As Integer = 622


Public Const iORO As Byte = 12
Public Const Pescado As Byte = 139

Public Enum PECES_POSIBLES
    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546
End Enum

Public Enum eMakro '(El 0 es no activado)
    Ninguno = 0
    PESCAR = 1
    PescarRed = 2
    Lingotear = 3
End Enum

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Enum eSkill
    Magia = 1
    Robar
    Tacticas
    Armas
    Meditar
    Apuñalar
    Ocultarse
    talar
    Defensa
    pesca
    Mineria
    Carpinteria
    Herreria
    Domar
    Proyectiles
    Wrestling
    Navegacion
End Enum

Public Const FundirMetal = 88

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Energia = 4
    Constitucion = 5
End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef As Byte = 15
Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
Public Const AumentoSTMago As Byte = AumentoSTDef - 1
Public Const AumentoSTLeñador As Byte = AumentoSTDef + 23
Public Const AumentoSTPescador As Byte = AumentoSTDef + 20
Public Const AumentoSTMinero As Byte = AumentoSTDef + 25

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 200
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 200
Public Const YMinMapSize As Byte = 1

'Tamaño del tileset
Public Const TileSizeX As Byte = 32
Public Const TileSizeY As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow As Byte = 23
Public Const YWindow As Byte = 17

'Sonidos
Public Const SND_SWING As Byte = 2
Public Const SND_TALAR As Byte = 13
Public Const SND_PESCAR As Byte = 14
Public Const SND_MINERO As Byte = 223
Public Const SND_WARP As Byte = 3
Public Const SND_PUERTA As Byte = 5
Public Const SND_NIVEL As Byte = 128

Public Const SND_USERMUERTE As Byte = 11
Public Const SND_IMPACTO As Byte = 10
Public Const SND_IMPACTO2 As Byte = 12
Public Const SND_LEÑADOR As Byte = 13
Public Const SND_FOGATA As Byte = 244
Public Const SND_AVE As Byte = 21
Public Const SND_AVE2 As Byte = 22
Public Const SND_AVE3 As Byte = 34
Public Const SND_GRILLO As Byte = 28
Public Const SND_GRILLO2 As Byte = 29
Public Const SND_SACARARMA As Byte = 25
Public Const SND_ESCUDO As Byte = 37
Public Const MARTILLOHERRERO As Byte = 41
Public Const LABUROCARPINTERO As Byte = 170
Public Const SND_BEBER As Byte = 135

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS As Integer = 10000

''
' Cantidad de "slots" en el inventario
Public Const MAX_INVENTORY_SLOTS As Byte = 28

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1


' CATEGORIAS PRINCIPALES
Public Enum eOBJType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otEsferadeExp = 36
    otPasajes = 37
    otPiedraResu = 38
    otMonturas = 39
    otCabezaMontura = 40
    otManual = 41
    otCofre = 42
    otRopaMontura = 43
    otCualquiera = 1000
End Enum

'Texto
Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
Public Const FONTTYPE_FIGHT As String = "~255~0~0~1~0"
Public Const FONTTYPE_WARNING As String = "~32~51~223~1~1"
Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
Public Const FONTTYPE_EJECUCION As String = "~130~130~130~1~0"
Public Const FONTTYPE_PARTY As String = "~255~180~255~0~0"
Public Const FONTTYPE_VENENO As String = "~0~255~0~0~0"
Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
Public Const FONTTYPE_SERVER As String = "~0~185~0~0~0"
Public Const FONTTYPE_GUILDMSG As String = "~228~199~27~0~0"
Public Const FONTTYPE_CONSEJO As String = "~130~130~255~1~0"
Public Const FONTTYPE_CONSEJOCAOS As String = "~255~60~00~1~0"
Public Const FONTTYPE_CONSEJOVesA As String = "~0~200~255~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~50~0~1~0"
Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"

'Estadisticas
Public Const STAT_MAXELV As Byte = 250
Public Const STAT_MAXHP As Integer = 999
Public Const STAT_MAXSTA As Integer = 5000
Public Const STAT_MAXMAN As Integer = 9999
Public Const STAT_MAXHIT_UNDER36 As Byte = 99
Public Const STAT_MAXHIT_OVER36 As Integer = 999
Public Const STAT_MAXDEF As Byte = 99
Public CantPremios As Integer


' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************

Public Type tHechizo
    nombre As String
    GrhIndex As Integer
    desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    targetMSG As String
    PropioMsg As String
'    Resis As Byte
    
    tipo As TipoHechizo
    
    wav As Integer
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
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    ExclusivoClase(1 To NUMCLASES) As eClass
    Morph As Byte
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    
    Invoca As Byte
    NumNpc As Integer
    cant As Integer

'    Materializa As Byte
'    ItemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    'Barrin 29/9/03
    StaRequerido As Integer

    Target As TargetType
End Type

Public Type LevelSkill
    LevelValue As Integer
End Type

Public Type UserObj
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
    ProbTirar As Integer
End Type

Public Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserObj
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
    AnilloEqpObjIndex As Integer
    AnilloEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    MonturaObjIndex As Integer
    MonturaSlot As Byte
    NroItems As Integer
End Type

Public Type tPartyData
    PIndex As Integer
    RemXP As Double 'La exp. en el server se cuenta con Doubles
    targetUser As Integer 'Para las invitaciones
End Type

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type worldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type FXdata
    nombre As String
    GrhIndex As Long
    Delay As Integer
End Type

'Datos de user o npc
Public Type Char
    CharIndex As Integer
    Head As Integer
    body As Integer
    AnimAtaque As Long
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    fx As Integer
    loops As Integer
    
    heading As eHeading
End Type

Private Type tObjCofre
    Obj As Integer
    cant As Integer
    Prob As Byte
End Type

Public Type tMaterial
    Material As Integer
    CantMaterial As Integer
End Type

'Tipos de objetos
Public Type ObjData
    Name As String 'Nombre del obj
    
    OBJType As eOBJType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Long ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    
    proyectil As Integer
    Municion As Integer
    
    NoLimpiar As Byte
    
    Crucial As Byte
    Newbie As Integer
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina

    'Pasajes
    DesdeMap As Long
    HastaMap As Long
    HastaY As Byte
    HastaX As Byte
    NecesitaSkill As Byte
    CantidadSkill As Byte
    
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
    
    def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    valor As Long     ' Precio
    
    Cerrada As Integer
    Llave As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    RazaDrow As Byte
    RazaElfa As Byte
    RazaGnoma As Byte
    RazaHumana As Byte
    
    Mujer As Byte
    Hombre As Byte
    
    Envenena As Byte
    Paraliza As Byte
    
    Agarrable As Byte
    
    CantMateriales As Byte
    Material(1 To 20) As tMaterial
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    skdomar As Integer
    
    texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As eClass
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    
    NoSeCae As Byte

    DañoMagico As Byte
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
    
    Log As Byte 'es un objeto que queremos loguear? Pablo (ToxicWaste) 07/09/07
    NoLog As Byte 'es un objeto que esta prohibido loguear?
    
    EfectoMagico As Byte
    
    IndiceMontura As Byte 'Indica cual es la montura que va aprender
    Cabeza As Integer
    
    IndiceSkill As Byte 'El indice del Skills
    CuantosSkill As Byte 'Cantidad de Skills que va a sumar
    SkNecesarios As Byte 'Cantidad de Skills que requiere para poder aprender el manual
    
    CantItems As Byte
    CofreCerrado As Byte '¿tiene llave?
    ItemCofre() As tObjCofre
    
    Nosetira As Byte 'Para que el objeto no se pueda tirar
    NoComerciable As Byte 'Para que el objeto no se pueda comerciar
    
    '********************
    'Suma stats
    SumaVida As Integer
    SumaMana As Integer
    SumaFuerza As Integer
    SumaAgilidad As Integer
    '********************
    Speed As Byte 'Aumenta la velocidad
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'[Pablo ToxicWaste]
Public Type ModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    DañoArmas As Double
    DañoProyectiles As Double
    DañoWrestling As Double
    Escudo As Double
    
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Energia As Single
    Constitucion As Single
End Type

Public Type ModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Energia As Single
    Constitucion As Single
End Type
'[/Pablo ToxicWaste]

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 80
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserObj
    NroItems As Integer
End Type
'[/KEVIN]


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    Promedio As Long
End Type

Public Type tPremios
    Puntos As Long
    ObjIndex As Long
    cantidad As Integer
End Type

'Estadisticas de los usuarios
Public Type UserStats
    GLD As Long 'Dinero
    Banco As Long
    
    MaxHP As Integer
    MinHP As Integer
    
    MaxSta As Integer
    MinSta As Integer
    
    MaxMAN As Integer
    MinMAN As Integer
    
    MaxHIT As Integer
    MinHIT As Integer
    
    MaxHam As Byte
    MinHam As Byte
    
    MaxAGU As Byte
    MinAGU As Byte
        
    def As Integer
    Exp As Double
    ELV As Byte
    ELU As Long
    ELO As Double
    UserSkills(1 To NUMSKILLS) As Byte
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Byte
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Long
    CriminalesMatados As Long
    NPCsMuertos As Integer
    
    DragCredits As Integer
    NUMMONTURAS As Byte
End Type

'Flags
Public Type UserFlags
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    Hambre As Byte
    Sed As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    
    NoPuedeSerAtacado As Boolean
    PuedeCambiarMapa As Boolean
    
    Vuela As Byte
    Navegando As Byte
    'Montando As Byte
    QueMontura As Byte 'Nos dice en que montura esta montando de su lista
    
    Seguro As Byte
    SeguroResu As Boolean
    
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    NpcInv As Integer
    
    Ban As Byte
    AdministrativeBan As Byte
    
    targetUser As Integer ' Usuario señalado
    
    targetObj As Integer ' Obj señalado
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
    NPCAtacado As Integer
    
    StatsChanged As Byte
    Privilegios As PlayerType
    
    ValCoDe As Integer
    
    LastCrimMatado As String
    LastCiudMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    AdminPerseguible As Boolean
    
    ChatColor As Long
    
    'Lorwik AntiSH
    Anomalia As Byte
    
    '[Barrin 30-11-03]
    TimesWalk As Long
    StartWalk As Long
    CountSH As Long
    '[/Barrin 30-11-03]
    
    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]
    
    Silenciado As Byte
    Mimetizado As Byte
    
    CentinelaIndex As Byte ' Indice del centinela que lo revisa ' 0.13.3
    CentinelaOK As Boolean 'Centinela
    
    Makro As eMakro
    EnTorneo As Boolean
    
    Montura(1 To 5) As tMontus
    
    EsperandoDueloSet As Boolean
    EstaDueleandoSet As Boolean
    OponenteSet As Integer
    PerdioRondaSet As Byte
    TimeDueloSet As Byte
    GanoDueloSet As Boolean
    Advertencias As Byte
    SerialHD As Long
    Morph As Integer
    
    DuelosClasicos As Integer
    
    AumentodeVida As Integer
    AumentodeMana As Integer
    AumentodeFuerza As Integer
    AumentodeAgilidad As Integer
    
    ArenaRinkel As Boolean
    Speed As Byte 'Modificador de la velocidad
End Type

Public Type UserCounters
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    Lava As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Makro As Integer
    Veneno As Integer
    Paralisis As Integer
    Morph As Integer
    Ceguera As Integer
    Estupidez As Integer
    
    Invisibilidad As Integer
    TiempoOculto As Integer
    
    Mimetismo As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As worldPos
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    
    'Barrin 3/10/03
    bPuedeMeditar As Boolean
    'Barrin
    
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeUsarArco As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerMagiaGolpe As Long
    TimerGolpeMagia As Long
    TimerGolpeUsar As Long
    TimerPuedeSerAtacado As Long
    TimerPuedeCambiardeMapa As Long
    TimerPuedeSendPing As Long
    
    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela
    
    failedUsageAttempts As Long
End Type

'Cosas faccionarias.
Public Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    Legion As Byte
    CriminalesMatados As Long
    CiudadanosMatados As Long
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    Reenlistadas As Byte
    NivelIngreso As Integer
    FechaIngreso As String
    MatadosIngreso As Integer 'Para Armadas nada mas
    NextRecompensa As Integer
End Type

'Tipo de los Usuarios
Public Type User
    'IndexPJ As Long ' Este es el ID que se le asignara en la DB
    Name As String
    id As Long
    
    
    
    '12/12/2018 Irongete: Guarda las zonas que está pisando el jugador
    Zona As New Collection
    
    '14/12/2018 Irongete: Guarda los efectos que tiene el jugador
    efecto As New Collection
    
    
    
    PartyIndex As Integer
    
    CuentaId As Integer
    
    showName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    Char As Char 'Define la apariencia
    CharMimetizado As Char
    OrigChar As Char
    
    desc As String ' Descripcion
    DescRM As String
    
    clase As eClass
    raza As eRaza
    genero As eGenero
        
    Invent As Inventario
    
    Pos As worldPos
    
    ConnIDValida As Boolean
    ConnID As Long 'ID
    
    '[KEVIN]
    BancoInvent As BancoInventario
    '[/KEVIN]
    
    Counters As UserCounters
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMascotas As Integer
    
    Stats As UserStats
    flags As UserFlags
    elpedidor As Integer
    
    Reputacion As tReputacion
    
    Faccion As tFacciones

#If ConUpTime Then
    LogOnTime As Date
    UpTime As Long
#End If

    ip As String
    IPLong As Long
    
     '[Alejo]
    ComUsu As tCOmercioUsuario
    '[/Alejo]
    
    GuildIndex As Integer   'puntero al array global de guilds
    FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    EscucheClan As Integer
    
    partyInvitacionIndex As Integer
    
    PartyId As Integer
    
    KeyCrypt As Integer
    
    AreasInfo As AreaInfo
    
    Accounted As String
    
    'Outgoing and incoming messages
    outgoingData As clsByteQueue
    incomingData As clsByteQueue
    
    Lac As TLac
End Type


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type NPCStats
    Alineacion As Integer
    TipoNick As Byte 'No pinta nada aqui, pero lo tenia que poner en algun sitio xD
    MaxHP As Long
    MinHP As Long
    MaxHIT As Integer
    MinHIT As Integer
    def As Integer
    defM As Integer
End Type

Public Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
End Type

Public Type NPCFlags
    Zona As Integer
    AfectaParalisis As Byte
    Domable As Integer
    Respawn As Byte
    Active As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    AtacaDoble As Byte
    LanzaSpells As Byte
    
    ExpCount As Long
    
    OldMovement As TipoAI
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    Sound As Integer
    AttackedBy As String
    AttackedFirstBy As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    SoloParty As Boolean
    
    LanzaMensaje As Byte
    Mensaje As String
    DijoMensaje As Boolean
    
    ActivoPotencia As Boolean
    AumentaPotencia As Boolean
    
    TiempoRetardoMax As Long
    TiempoRetardoMin As Long
    Retardo As Byte
    
    Explota As Byte
    VerInvi As Byte
    
    ArenasRinkel As Byte 'Identifica si un NPC pertenece al evento de arenas de Rinkel
  
    Speed As Byte
End Type

Public Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    targetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type
' New type for holding the pathfinding info


Public Type npc
    Name As String
    Char As Char 'Define como se vera
    desc As String
    Nivel As Byte
    
    NPCType As eNPCType
    Numero As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer

    Veneno As Byte
    
    efecto As New Collection

    Pos As worldPos 'Posicion
    Orig As worldPos
    SkillDomar As Integer

    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long


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
    
    ' New!! Needed for pathfindig
    PFINFO As NpcPathFindingInfo
    AreasInfo As AreaInfo
    
    
    Quest As Byte
    
    
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Long
    UserIndex As Integer
    NpcIndex As Integer
    ObjInfo As Obj
    TileExit As worldPos
    trigger As eTrigger
End Type

Public Type MapInfo
    NumUsers As Integer
    MapVersion As Integer
    Pk As Boolean
    Invocado As Boolean
    MagiaSinEfecto As Byte
    NoEncriptarMP As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    lvlMinimo As Byte
    MapName As String
    SePuedeDomar As Byte
    
    Zonas() As Long
End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public ULTIMAVERSION As String
Public BackUp As Boolean ' TODO: Se usa esta variable ?

'OPCIONES
Public iniDragDrop As Byte
Public iniTirarOBJZonaSegura As Byte
Public iniAutoSacerdote As Boolean
Public iniSacerdoteCuraVeneno As Boolean

Public ListaRazas(1 To NUMRAZAS) As String
Public SkillsNames(1 To NUMSKILLS) As String
Public ListaClases(1 To NUMCLASES) As String
Public ListaAtributos(1 To NUMATRIBUTOS) As String


Public recordusuarios As Integer 'Lorwik> Lo pongo en Integer por que dudo mucho que superemos su valor jajaja

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath As String

''
'Ruta base para guardar los chars
Public CharPath As String

''
'Ruta base para los archivos de mapas
Public MapPath As String

''
'Ruta base para los DATs
Public DatPath As String

''
'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

''
'Numero de usuarios actual
Public NumUsers As Integer
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public TotalNPCDat As Integer
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
Public HaciendoBackup As Boolean
Public PuedeCrearPersonajes As Integer
Public ServerSoloGMs As Integer
Public RequiereValidacionACC As Byte

Public EnPausa As Boolean
Public EnTesting As Boolean


'*****************ARRAYS PUBLICOS*************************
Public UserList() As User 'USUARIOS
Public NPCList(1 To MAXNPCS) As npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList(1 To MAXCHARS) As Integer
Public ObjData() As ObjData
Public fx() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 50) As LevelSkill
Public ForbidenNames() As String
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
Public BanIPs As New Collection
Public BanHDs As New Collection
Public PremiosInfo() As tPremios
'Public Parties(1 To MAX_PARTIES) As clsParty
Public ModClase(1 To NUMCLASES) As ModClase
Public ModRaza(1 To NUMRAZAS) As ModRaza
Public ModVida(1 To NUMCLASES) As Double
Public DistribucionEnteraVida(1 To 5) As Integer
Public DistribucionSemienteraVida(1 To 4) As Integer
'*********************************************************

Public Nix As worldPos
Public Ullathorpe As worldPos
Public Banderbill As worldPos
Public Lindos As worldPos
Public Arghal As worldPos

Public Prision As worldPos
Public Libertad As worldPos

Public Ayuda As New cCola
Public ConsultaPopular As New ConsultasPopulares
Public SonidosMapas As New SoundMapInfo

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)

Public Enum e_ObjetosCriticos
    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467
End Enum

'Lorwik> Sistema de retardo de Spawn de NPC
Type tRetarded
    Tiempo As Long
    Mapa As Byte
    X As Byte
    Y As Byte
    NPCNUM As Integer
End Type

Public RetardoSpawn(1 To MAXNPCS) As tRetarded
