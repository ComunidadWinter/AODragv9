Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

Public KeyFisico As Boolean

'*************************************
'CONSEJOS
'*************************************
Public Consejos(1 To 100) As String
Public ListaConsejos() As String
Public ConsejoSeleccionado As String
'*************************************

'**********************
'DIRECTORIOS
'**********************
Public Const DirCursores As String = "\Recursos\Cursores\"
Public Const MouseHand As String = "\recursos\cursores\d.ico"
'**********************

Type Rank
    name As String
    ELO As Double
End Type

Public Ranking(5) As Rank
Public LlegoRank  As Boolean

'************************************************
'IP y PUERTO del server
'************************************************
Public CurServerIp As String
Public CurServerPort As Integer

Public Type Servidores
    Nombre As String
    Ip As String
    Puerto As Integer
End Type

Public Servidor(0 To 100) As Servidores
Public ServIndSel As Byte
'************************************************

'Almacenaremos los mensajes predefinidos
Public MultiMensaje(0 To 255) As tMultiMessage
    
Type tMultiMessage
    mensaje As String
End Type

'*****Lorwik Clima***********
Public DayStatus As Byte
'****************************
Public Const OFFSET_HEAD As Integer = -5 'De los dialogos

Public HeadSeleccion As Long  'Seleccion de Cabezas
Public Form_Caption As String
Public Win2kXP As Boolean

'******Conectar renderizado*********
Private Type tMapaConnect
    Map As Byte
    X As Byte
    Y As Byte
End Type

Public MapaConnect As tMapaConnect

Public lbFuerza As Byte
Public lbAgilidad As Byte
Public lbInteligencia As Byte
Public lbEnergia As Byte
Public lbConstitucion As Byte

'***********************************

'Inventarios de comercio con usuario
Public InvComUsu As New clsGraphicalInventory ' Inventario del usuario visible en el comercio
Public InvComNpc As New clsGraphicalInventory ' Inventario con los items que ofrece el npc

'Objetos públicos
Public Spells As New clsGraphicalSpells
Public Inventario As New clsGraphicalInventory
Public InventarioComNpc As New clsGraphicalInventory
Public InventarioComUser As New clsGraphicalInventory
Public InvBanco(1) As New clsGraphicalInventory

Public SurfaceDB As clsSurfaceManager   'No va new porque es una interfaz, el new se pone al decidir que clase de objeto es
Public Sound As clsSoundEngine
Public CustomKeys As New clsCustomKeys
Public texto As New clsDX8Font 'Textos renderizado
Public Dialogos As New clsDialogs

Public incomingData As New clsByteQueue
Public outgoingData As New clsByteQueue

''
'The main timer of the game.
Public MainTimer As New clsTimer


'*************************************************
'Sonidos
'*************************************************

Public Enum TipoPaso
    CONST_BOSQUE = 1
    CONST_NIEVE = 2
    CONST_CABALLO = 3
    CONST_DUNGEON = 4
    CONST_PISO = 5
    CONST_DESIERTO = 6
    CONST_PESADO = 7
End Enum

Public Type tPaso
    CantPasos As Byte
    Wav() As Integer
End Type

Public Const NUM_PASOS As Byte = 7
Public Pasos() As tPaso

Public Const MUS_Inicio As String = "6"
Public Const MUS_CrearPersonaje As String = "7"
Public Const MUS_VolverInicio As String = "53"

Public Const SND_CLICK As Byte = 190
Public Const SND_NAVEGANDO As Byte = 50
Public Const SND_OVER As Byte = 0
Public Const SND_DICE As Byte = 188
Public Const SND_FUEGO As Byte = 79
Public Const SND_AMBIENTE_NOCHE As Integer = 20
Public Const SND_AMBIENTE_NOCHE_CIU As Integer = 21

'Musica
Public Const Musica_Inicio As Byte = 6
Public Const Music_Desconectar As Byte = 53
'*************************************************

' Head index of the casper. Used to know if a char is killed

'*************************************************
' Constantes de intervalo
'*************************************************
Public Const INT_ATTACK As Integer = 100
Public Const INT_ARROWS As Integer = 100
Public Const INT_CAST_SPELL As Integer = 500
Public Const INT_CAST_ATTACK As Integer = 0
Public Const INT_WORK As Integer = 100
Public Const INT_USEITEMU As Integer = 300
Public Const INT_USEITEMDCK As Integer = 300
Public Const INT_SENTRPU As Integer = 2000

Public MacroBltIndex As Integer

Public Const CASPER_HEAD As Integer = 500
Public Const FRAGATA_FANTASMAL As Integer = 87

Public Const NUMATRIBUTES As Byte = 5

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public ColoresPJ(0 To 50) As Long

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean
Public UserCharIndex As Integer
Public UserSpeed As Byte

Public Enum E_SISTEMA_MUSICA
    CONST_DESHABILITADA = 0
    CONST_MP3 = 1
    CONST_MIDI = 2
End Enum

Private Type tOption
    NoRes As Byte 'no cambiar la resolucion
    BaseTecho As Byte
    bCursores As Byte
    MovEscritura As Byte
    URLCON As Byte
    NamePlayers As Byte
    PrimeraVez As Byte
    GuildNews As Byte
    VSynC As Byte
    VProcessing As Byte
    MusicVolume As Long
    HechizosClasicos As Byte
    Ambient As Byte
    AmbientVol As Long
    Audio As Byte
    FxNavega As Long
    InvertirSonido As Byte
    FXVolume As Long
    sMusica As E_SISTEMA_MUSICA
    BloqCruceta As Byte
End Type

Public Opciones As tOption

Public Const GRH_FOGATA As Integer = 1521

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer
Public NumAtaques As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

Public Versiones(1 To 7) As Integer

Public UsaMacro As Boolean
Public CnTd As Byte


'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 80
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]


Public Tips() As String * 255
Public Const LoopAdEternum As Integer = 999

'Direcciones
Public Enum E_Heading
    SOUTH = 1
    NORTH = 2
    WEST = 3
    EAST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 28
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAX_SPELL_SLOTS As Byte = 18
Public PremiosInv(1 To 20) As PremiosList
Public Const MAX_LEVEL As Byte = 250

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1

Public Const FOgata As Integer = 1521

Type PremiosList
    name As String
    Puntos As Integer
    cantidad As Integer
End Type

Public Enum eClass
    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Bard        'Bardo
    Druid       'Druida
    Paladin     'Paladín
    Hunter      'Cazador
End Enum

Enum eRaza
    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano
    orco
    nomuerto
End Enum

Public Enum eSkill
    magia = 1
    Robar
    Tacticas
    Armas
    Meditar
    Apuñalar
    Ocultarse
    Talar
    Defensa
    Pesca
    Mineria
    Carpinteria
    Herreria
    Domar
    Proyectiles
    Wrestling
    Navegacion
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Energia = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer
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

Public Enum eObjType
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
    otescudo = 16
    otcasco = 17
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
    otCualquiera = 1000
End Enum

Public Const FundirMetal As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "¡¡¡La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "¡¡¡La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "¡¡¡Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "¡¡¡El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "¡Has fallado el golpe!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = ">>SEGURO ACTIVADO<<"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = ">>SEGURO DESACTIVADO<<"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_SEGURO_RESU_ON As String = "SEGURO DE RESURRECCION ACTIVADO"
Public Const MENSAJE_SEGURO_RESU_OFF As String = "SEGURO DE RESURRECCION DESACTIVADO"

Public Const MENSAJE_GOLPE_CABEZA As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "Causas "
Public Const MENSAJE_GOLPE_CRIATURA_2 As String = " puntos de daño a  "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la víctima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

'Inventario
Type Inventory
    OBJIndex As Integer
    name As String
    GrhIndex As Long
    '[Alejo]: tipo de datos ahora es Long
    amount As Long
    '[/Alejo]
    Equipped As Byte
    valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    PuedeUsar As Integer
End Type

Type NpCinV
    OBJIndex As Integer
    name As String
    GrhIndex As Long
    amount As Integer
    valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
    PuedeUsar As Byte
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

'*******************************
'Nombre de mapa y su transparencia
Public LastMapName As String
Public TransMapAB As Byte
'*******************************

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAX_SPELL_SLOTS) As Spells

Type Spells
    Index As Integer
    name As String
    GrhIndex As Long
End Type

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public UserMeditar As Boolean
Public UserName As String
Public UserNameClan As String
Public UserPartyId As Integer
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte

Public UserFuerza As Byte
Public UserAgilidad As Byte

Public UserWeaponEqpSlot As Byte
Public UserArmourEqpSlot As Byte
Public UserHelmEqpSlot As Byte
Public UserShieldEqpSlot As Byte

Public UserGLD As Long
Public UserELO As Double
Public UserDragCreditos As Integer
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public FPSFLAG As Boolean
Public MiPing As Long
Public pausa As Boolean
Public IsAttacking As Boolean
Public NPCAtaqueIndex As Integer
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserMontando As Boolean
Public UserAvisado As Boolean

Public MonturaSeleccionada As Byte
'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public lblexpactivo As Boolean

Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 5
Public Const NUMSKILLS As Byte = 17
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 8
Public Const NUMRAZAS As Byte = 7

Public UserSkills(1 To NUMSKILLS) As Byte
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String
Public SendingType As Byte
Public sndPrivateTo As String
    
Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public Logged As Boolean

Public UsingSkill As Integer

Public pingTime As Long

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    LoginCuenta = 4
    BorrandoPJ = 5
End Enum

Public EstadoLogin As E_MODO

'**********Cuentas*************
Public Type pjs
    NamePJ As String
    LvlPJ As Integer
    ClasePJ As eClass
    
    Acuerpo As Integer
    rcvHead As Integer
    rcvCasco As Integer
    rcvShield As Integer
    rcvWeapon As Integer
    rcvRaza As Integer
    PJLogged As Byte
    mapa As String
End Type

Public Type acc
    name As String
    Pass As String
    Email As String
    preg As String
    resp As String
   
    CantPJ As Byte
    pjs(1 To 8) As pjs
End Type

Public Cuenta As acc
Public IndexSelectedUSer As Byte
Public PJName As String
Public NameAccount As String

'*************************

Public Enum FxMeditar
    CHICO = 46
    MEDIANO = 3
    GRANDE = 49
    XGRANDE = 40
    XXGRANDECIU = 34
    XXGRANDECRI = 35
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Speed
End Enum

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
Public Enum eTrigger
    Nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'Control
Public prgRun As Boolean 'When true the program ends
Public IPdelServidor As String
Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'
'Graficos
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Lista de cabezas
Public Type tHead
    Texture As Integer
    startX As Integer
    startY As Integer
End Type

Public heads() As tHead
Public Cascos() As tHead

Public Type tIndiceCuerpo
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceAtaque
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Long
    offsetX As Integer
    offsetY As Integer
    FXTransparente As Boolean
End Type

Public Enum eMoveType
    Inventory = 1
    Bank = 2
    SpellsI = 3
End Enum

Public Enum eCursorState
    cur_Normal = 0
    cur_Action
    cur_Wait
    cur_Npc
    cur_Npc_Hostile
    cur_User
    cur_User_Danger
    cur_Obj
End Enum

Public CurrentCursor As eCursorState
Public picMouseIcon As Picture

Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public SolicitudParty As Boolean

'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream
 
'RGB Type
Public Type RGB
    r As Long
    g As Long
    b As Long
End Type
 
Public Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    Angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    Speed As Single
    life_counter As Long

    Radio As Integer
End Type

'CopyMemory Kernel Function
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
