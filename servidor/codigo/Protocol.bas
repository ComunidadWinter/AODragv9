Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Public Const SEPARATOR As String * 1 = vbNullChar

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 246

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue


Private Enum ServerPacketID
    logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    CreateRenderText
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCSwing                ' N1
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    UserSwing               ' U1
    NobilityLost            ' PN
    CantUseWhileMeditating  ' M!
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ChangeSpeed
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMidi                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong
    UpdateTagAndStatus
    ShowPuntosClanes
    
    'Paquetes de party
    PeticionInvitacionAParty
    ActualizacionParty
    SetPartyId
    
    '15/12/2018 Irongete: Paquetes de Efectos
    CrearEfecto
    QuitarEfecto
    RenovarEfecto
    
    '15/12/2018 Irongete:
    CrearZona
    QuitarZona
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    Noche
    CuentaSV
    CastleAttack
    MontateToggle
    EnvioListaMonturas
    EnvioInfoMontura
    EsteTioQuiereParty
    UneteamiParty
    muestropremios
    AtaqueNPC
    UpdateGLDBank
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    rCaptions
    ShowCaptions
    SegNoti
    EnviaManual
    EnviarRanking
    PartyClanIndex
    ShowReciclarHerreria
    limpiaracc
    MultiMessage
    SubInit
End Enum

Private Enum ClientPacketID
    LoginExistingChar       'OLOGIN
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    MoveSpell               'DESPHE
    MoveBank
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    PartyLeave              '/SALIRPARTY
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    PartyOnline             '/ONLINEPARTY
    CouncilMessage          '/BMSG
    GMRequest               '/GM
    ChangeDescription       '/DESC
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    GuildFundate            '/FUNDARCLAN
    Ping                    '/PING
    CastillosCMD            '/CASTILLOS
    PuntosClanes
    IrAlCementerio
    
    '18/11/2015 Irongete: Paquetes de party
    InvitacionAParty
    RespuestaPeticionAParty
    ExpulsarDeParty
    NuevoLiderDeParty
    AbandonarParty
    ActualizarParty
    
    GlobalMsg               '/GLOBAL
    
    'GM messages
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterName               '/ANAME
    ToggleCentinelActivated '/CENTINELAACTIVADO
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    night                   '/NOCHE
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    
    Noche
    LoginCuenta
    BorrarPJ
    DueloSet
    MoveItem                'Drag and drop 0.13.3
    DropObjTo               'Drop to pos.
    SolInfMonturaClient
    SolParty
    QuieroParty
    RespuestaParty
    Premios
    rCaptions
    SCAPTIONS
    AdvChar
    GMCommands
    TPCastle
    SolicitarRank
    Torneo
    MovePj                  '/MOVER
    ReleasePet
    MotdCastillos
    OlvidaSpell
    SalirSubasta
    NuevaSubasta
End Enum

Public Enum eGMCommands
    VerHD                   '/VERHD USUARIO
    BanHD                   '/BANHD USUARIO
    UnBanHD                 '/UNBANHD NROHD
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    BanACC              '/BANACC
    UnbanACC                '/UNBANACC
    SearchObj               '/BUSCAR [OBJETO]
    Searchnpc               '/BUSCARNPC [NPC]
    EntregarDC              '/ENTREGARDC
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    HacerTorneo
    Cerrartorneo
    GlobalOn                '/GLOBALON
    GlobalOff               '/GLOBALOFF
    Jail                    '/CARCEL
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_oro
    FONTTYPE_exp
    FONTTYPE_global
    FONTTYPE_PRIVADO
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
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
On Error Resume Next
    Dim packetID As Byte
    
    packetID = UserList(UserIndex).incomingData.PeekByte()
    
    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.LoginExistingChar _
      Or packetID = ClientPacketID.LoginNewChar _
      Or packetID = ClientPacketID.LoginCuenta _
      Or packetID = ClientPacketID.BorrarPJ) Then
        
        'Is the user actually logged?
        If Not UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Function
        
        'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            UserList(UserIndex).Counters.IdleCount = 0
        End If
    ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
        UserList(UserIndex).Counters.IdleCount = 0
        
        'Is the user logged?
        If UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Function
        End If
    End If
    
    ' Ante cualquier paquete, pierde la proteccion de ser atacado.
    UserList(UserIndex).flags.NoPuedeSerAtacado = False
    
    Select Case packetID
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Debug.Print "OLOGIN"
            Call HandleLoginExistingChar(UserIndex)
            
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Debug.Print "NLOGIN"
            Call HandleLoginNewChar(UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
        
        Case ClientPacketID.Walk
            Call HandleWalk(UserIndex)
            
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
            
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(UserIndex)
        
        Case ClientPacketID.ResuscitationSafeToggle
            Call HandleResuscitationToggle(UserIndex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(UserIndex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(UserIndex)
        
        Case ClientPacketID.RequestFame             'FAMA
            Call HandleRequestFame(UserIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(UserIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(UserIndex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(UserIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(UserIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(UserIndex)
            
        Case ClientPacketID.MoveBank
            Call HandleMoveBank(UserIndex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
        
        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(UserIndex)
        
        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(UserIndex)
        
        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(UserIndex)
        
        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(UserIndex)
        
        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(UserIndex)
        
        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(UserIndex)
        
        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(UserIndex)
        
        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(UserIndex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(UserIndex)
        
        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(UserIndex)
        
        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(UserIndex)
        
        Case ClientPacketID.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(UserIndex)
        
        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(UserIndex)
        
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(UserIndex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(UserIndex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(UserIndex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(UserIndex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(UserIndex)
        
        Case ClientPacketID.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(UserIndex)
        
        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(UserIndex)
        
        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(UserIndex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(UserIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(UserIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(UserIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(UserIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
            
        Case ClientPacketID.GlobalMsg               '/GLOBAL
            Call HandleGlobalMsg(UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(UserIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(UserIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(UserIndex)
        
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)
        
        Case ClientPacketID.PartyLeave              '/SALIRPARTY
            Call HandlePartyLeave(UserIndex)
        
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(UserIndex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(UserIndex)
        
        Case ClientPacketID.PartyMessage            '/PMSG
            Call HandlePartyMessage(UserIndex)
        
        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(UserIndex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(UserIndex)
        
        Case ClientPacketID.PartyOnline             '/ONLINEPARTY
            Call HandlePartyOnline(UserIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(UserIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(UserIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)

        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(UserIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(UserIndex)
        
        Case ClientPacketID.GuildFundate            '/FUNDARCLAN
            Call HandleGuildFundate(UserIndex)
                
        Case ClientPacketID.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
            
        Case ClientPacketID.CastillosCMD                   '/CASTILLOS
            Call HandleCastillosCMD(UserIndex)
            
        Case ClientPacketID.PuntosClanes
            Call HandlePuntosClanes(UserIndex)
            
        Case ClientPacketID.IrAlCementerio
            Call HandleIrAlCementerio(UserIndex)
            
        Case ClientPacketID.InvitacionAParty
            Call HandleInvitacionAParty(UserIndex)
            
        Case ClientPacketID.RespuestaPeticionAParty
            Call HandleRespuestaPeticionAParty(UserIndex)
            
        Case ClientPacketID.AbandonarParty
            Call HandleAbandonarParty(UserIndex)
            
            
        Case ClientPacketID.ActualizarParty
            Call HandleActualizarParty(UserIndex)
            
        Case ClientPacketID.ExpulsarDeParty
            Call HandleExpulsarDeParty(UserIndex)
            
        Case ClientPacketID.NuevoLiderDeParty
            Call HandleNuevoLiderDeParty(UserIndex)
        
        
        'GM messages
        
        Case ClientPacketID.RequestUserList         'LISTUSU
            Call HandleRequestUserList(UserIndex)
        
        Case ClientPacketID.Working                 '/TRABAJANDO
            Call HandleWorking(UserIndex)
        
        Case ClientPacketID.Hiding                  '/OCULTANDO
            Call HandleHiding(UserIndex)
        
        Case ClientPacketID.KillNPC                 '/RMATA
            Call HandleKillNPC(UserIndex)
        
        Case ClientPacketID.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(UserIndex)
        
        Case ClientPacketID.EditChar                '/MOD
            Call HandleEditChar(UserIndex)
            
        Case ClientPacketID.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(UserIndex)
        
        Case ClientPacketID.RequestCharStats        '/STAT
            Call HandleRequestCharStats(UserIndex)
            
        Case ClientPacketID.RequestCharGold         '/BAL
            Call HandleRequestCharGold(UserIndex)
            
        Case ClientPacketID.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(UserIndex)
            
        Case ClientPacketID.RequestCharBank         '/BOV
            Call HandleRequestCharBank(UserIndex)
        
        Case ClientPacketID.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(UserIndex)
        
        Case ClientPacketID.ReviveChar              '/REVIVIR
            Call HandleReviveChar(UserIndex)
        
        Case ClientPacketID.Forgive                 '/PERDON
            Call HandleForgive(UserIndex)
            
        Case ClientPacketID.Kick                    '/ECHAR
            Call HandleKick(UserIndex)
            
        Case ClientPacketID.Execute                 '/EJECUTAR
            Call HandleExecute(UserIndex)
            
        Case ClientPacketID.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(UserIndex)
            
        Case ClientPacketID.SummonChar              '/SUM
            Call HandleSummonChar(UserIndex)
            
        Case ClientPacketID.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(UserIndex)
            
        Case ClientPacketID.SpawnCreature           'SPA
            Call HandleSpawnCreature(UserIndex)
            
        Case ClientPacketID.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(UserIndex)
            
        Case ClientPacketID.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(UserIndex)
            
        Case ClientPacketID.ServerMessage           '/RMSG
            Call HandleServerMessage(UserIndex)
            
        Case ClientPacketID.NickToIP                '/NICK2IP
            Call HandleNickToIP(UserIndex)
        
        Case ClientPacketID.IPToNick                '/IP2NICK
            Call HandleIPToNick(UserIndex)
            
        Case ClientPacketID.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(UserIndex)
        
        Case ClientPacketID.TeleportCreate          '/CT
            Call HandleTeleportCreate(UserIndex)
            
        Case ClientPacketID.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(UserIndex)
        
        Case ClientPacketID.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(UserIndex)
        
        Case ClientPacketID.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(UserIndex)
            
        Case ClientPacketID.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(UserIndex)
            
        Case ClientPacketID.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(UserIndex)
                        
        Case ClientPacketID.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(UserIndex)
            
        Case ClientPacketID.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(UserIndex)
            
        Case ClientPacketID.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(UserIndex)
            
        Case ClientPacketID.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(UserIndex)
        
        Case ClientPacketID.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(UserIndex)
            
        Case ClientPacketID.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(UserIndex)
            
        Case ClientPacketID.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(UserIndex)
            
        Case ClientPacketID.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(UserIndex)
            
        Case ClientPacketID.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(UserIndex)
            
        Case ClientPacketID.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(UserIndex)
            
        Case ClientPacketID.DumpIPTables            '/DUMPSECURITY"
            Call HandleDumpIPTables(UserIndex)
            
        Case ClientPacketID.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(UserIndex)
        
        Case ClientPacketID.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(UserIndex)
        
        Case ClientPacketID.AskTrigger               '/TRIGGER
            Call HandleAskTrigger(UserIndex)
            
        Case ClientPacketID.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(UserIndex)
        
        Case ClientPacketID.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(UserIndex)
        
        Case ClientPacketID.GuildBan                '/BANCLAN
            Call HandleGuildBan(UserIndex)
        
        Case ClientPacketID.BanIP                   '/BANIP
            Call HandleBanIP(UserIndex)
        
        Case ClientPacketID.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(UserIndex)
        
        Case ClientPacketID.CreateItem              '/CI
            Call HandleCreateItem(UserIndex)
        
        Case ClientPacketID.DestroyItems            '/DEST
            Call HandleDestroyItems(UserIndex)
            
        Case ClientPacketID.MovePj                  '/MOVER
            Call HandleMovePj(UserIndex)
        
        Case ClientPacketID.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(UserIndex)
        
        Case ClientPacketID.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(UserIndex)
        
        Case ClientPacketID.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(UserIndex)
        
        Case ClientPacketID.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(UserIndex)
        
        Case ClientPacketID.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(UserIndex)
        
        Case ClientPacketID.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(UserIndex)
        
        Case ClientPacketID.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(UserIndex)
        
        Case ClientPacketID.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(UserIndex)
        
        Case ClientPacketID.LastIP                  '/LASTIP
            Call HandleLastIP(UserIndex)
        
        Case ClientPacketID.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(UserIndex)
        
        Case ClientPacketID.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(UserIndex)
        
        Case ClientPacketID.SystemMessage           '/SMSG
            Call HandleSystemMessage(UserIndex)
        
        Case ClientPacketID.CreateNPC               '/ACC
            Call HandleCreateNPC(UserIndex)
        
        Case ClientPacketID.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(UserIndex)
        
        Case ClientPacketID.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(UserIndex)
        
        Case ClientPacketID.ServerOpenToUsersToggle '/HABILITAR
            Call HandleServerOpenToUsersToggle(UserIndex)
        
        Case ClientPacketID.TurnOffServer           '/APAGAR
            Call HandleTurnOffServer(UserIndex)
        
        Case ClientPacketID.TurnCriminal            '/CONDEN
            Call HandleTurnCriminal(UserIndex)
        
        Case ClientPacketID.ResetFactions           '/RAJAR
            Call HandleResetFactions(UserIndex)
        
        Case ClientPacketID.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(UserIndex)
        
        Case ClientPacketID.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(UserIndex)
        
        Case ClientPacketID.AlterName               '/ANAME
            Call HandleAlterName(UserIndex)
        
        Case ClientPacketID.ToggleCentinelActivated '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(UserIndex)
        
        Case ClientPacketID.DoBackUp                '/DOBACKUP
            Call HandleDoBackUp(UserIndex)
        
        Case ClientPacketID.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(UserIndex)
        
        Case ClientPacketID.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(UserIndex)
        
        Case ClientPacketID.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(UserIndex)
        
        Case ClientPacketID.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(UserIndex)
    
        Case ClientPacketID.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(UserIndex)
            
        Case ClientPacketID.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(UserIndex)
        
        Case ClientPacketID.SaveChars               '/GRABAR
            Call HandleSaveChars(UserIndex)
        
        Case ClientPacketID.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(UserIndex)
        
        Case ClientPacketID.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(UserIndex)
            
        Case ClientPacketID.night                   '/NOCHE
            Call HandleNight(UserIndex)
        
        Case ClientPacketID.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(UserIndex)
        
        Case ClientPacketID.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(UserIndex)
        
        Case ClientPacketID.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(UserIndex)
        
        Case ClientPacketID.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(UserIndex)
        
        Case ClientPacketID.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(UserIndex)
        
        Case ClientPacketID.Restart                 '/REINICIAR
            Call HandleRestart(UserIndex)
        
        Case ClientPacketID.ChatColor               '/CHATCOLOR
            Call HandleChatColor(UserIndex)
        
        Case ClientPacketID.Ignored                 '/IGNORADO
            Call HandleIgnored(UserIndex)
        
        Case ClientPacketID.CheckSlot               '/SLOT
            Call HandleCheckSlot(UserIndex)
            
        Case ClientPacketID.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
            Call HandleSetIniVar(UserIndex)
            
        Case ClientPacketID.Noche                   'NOCHE
            Call HandleNoche(UserIndex)
            
        Case ClientPacketID.LoginCuenta
            Call HandleLoginCuenta(UserIndex)
            
        Case ClientPacketID.BorrarPJ
            Call HandleBorrarPJ(UserIndex)
            
        Case ClientPacketID.DueloSet
            Call HandleDueloSet(UserIndex)
            
        Case ClientPacketID.MoveItem                ' 0.13.3
            Call HandleMoveItem(UserIndex)
            
        Case ClientPacketID.DropObjTo               ' Drop to pos.
            Call HandleDropObj(UserIndex)
            
        Case ClientPacketID.SolInfMonturaClient
            Call HandleSolInfMonturaClient(UserIndex)
            
        Case ClientPacketID.SolParty
            Call HandleSolParty(UserIndex)
            
        Case ClientPacketID.QuieroParty
            'Call HandleQuieroParty(UserIndex)
            
        Case ClientPacketID.RespuestaParty
            'Call HandleRespuestaParty(UserIndex)
            
        Case ClientPacketID.Premios
            Call HandlePremios(UserIndex)
            
        Case ClientPacketID.rCaptions
          Call HandleRequieredCaptions(UserIndex)
         
        Case ClientPacketID.SCAPTIONS
          Call HandleSendCaptions(UserIndex)
          
        Case ClientPacketID.AdvChar
            Call HandleAdvertirPJ(UserIndex)
            
        Case ClientPacketID.GMCommands              'GM Messages
            Call HandleGMCommands(UserIndex)
            
        Case ClientPacketID.TPCastle
            Call HandleTPCastle(UserIndex)
            
        Case ClientPacketID.SolicitarRank
            Call HandleSolicitarRank(UserIndex)
        
        Case ClientPacketID.Torneo
            Call HandleTorneo(UserIndex)
                
        Case ClientPacketID.ReleasePet              '/LIBERAR
            Call HandleReleasePet(UserIndex)
            
        Case ClientPacketID.MotdCastillos
            Call HandleMotdCastillos(UserIndex)
            
        Case ClientPacketID.OlvidaSpell
            Call HandleOlvidaSpell(UserIndex)
            
        Case ClientPacketID.SalirSubasta
            Call HandleSalirSubasta(UserIndex)
            
        Case ClientPacketID.NuevaSubasta
            Call HandleNuevaSubasta(UserIndex)

        Case Else
            'ERROR : Abort!
            Call CloseSocket(UserIndex)
    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.length > 0 And Err.Number = 0 Then
        Err.Clear
        HandleIncomingData = True
    
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        
        'Call LogError("Error: " & Err.Number & " [" & Err.Description & "] " & " Source: " & Err.source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & packetID)
                        
         Debug.Print "Number: " & Err.Number & " packetID: " & packetID & " Description: " & Err.Description
        HandleIncomingData = False
        
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(UserIndex)
        HandleIncomingData = False
    End If
End Function

Private Sub HandleGMCommands(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: Lorwik
'COMANDOS GM
'***************************************************

On Error GoTo errHandler

Dim Command As Byte

With UserList(UserIndex)
    Call .incomingData.ReadByte
    
    Command = .incomingData.PeekByte
    
    Select Case Command
    
        Case eGMCommands.VerHD                   '/VERHD NICKUSUARIO
            Call HandleVerHD(UserIndex)
               
        Case eGMCommands.BanHD                   '/BANHD NICKUSUARIO
            Call HandleBanHD(UserIndex)
       
        Case eGMCommands.UnBanHD                 '/UNBANHD NICKUSUARIO
            Call HandleUnbanHD(UserIndex)
            
        Case eGMCommands.BanChar                 '/BAN
            Call HandleBanChar(UserIndex)
            
        Case eGMCommands.UnbanChar               '/UNBAN
            Call HandleUnbanChar(UserIndex)
            
        Case eGMCommands.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(UserIndex)
        
        Case eGMCommands.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(UserIndex)
            
        Case eGMCommands.BanACC                  '/BANACC
            Call HandleBanAcc(UserIndex)
            
        Case eGMCommands.UnbanACC                '/UNBANACC
            Call HandleUnbanAcc(UserIndex)
            
        Case eGMCommands.SearchObj               '/BUSCAR [OBJETO]
            Call HandleSearchObj(UserIndex)
            
        Case eGMCommands.Searchnpc               '/BUSCARNPC [ONPC]
            Call HandleSearchNPC(UserIndex)
            
        Case eGMCommands.EntregarDC              '/ENTREGARDC
            Call HandleEntregarDC(UserIndex)
            
        Case eGMCommands.GMMessage               '/GMSG
            Call HandleGMMessage(UserIndex)
        
        Case eGMCommands.showName                '/SHOWNAME
            Call HandleShowName(UserIndex)
        
        Case eGMCommands.OnlineRoyalArmy         '/ONLINEREAL
            Call HandleOnlineRoyalArmy(UserIndex)
        
        Case eGMCommands.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(UserIndex)
        
        Case eGMCommands.GoNearby                '/IRCERCA
            Call HandleGoNearby(UserIndex)
        
        Case eGMCommands.comment                 '/REM
            Call HandleComment(UserIndex)
        
        Case eGMCommands.serverTime              '/HORA
            Call HandleServerTime(UserIndex)
        
        Case eGMCommands.Where                   '/DONDE
            Call HandleWhere(UserIndex)
        
        Case eGMCommands.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(UserIndex)
        
        Case eGMCommands.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(UserIndex)
        
        Case eGMCommands.WarpChar                '/TELEP
            Call HandleWarpChar(UserIndex)
        
        Case eGMCommands.Silence                 '/SILENCIAR
            Call HandleSilence(UserIndex)
        
        Case eGMCommands.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(UserIndex)
        
        Case eGMCommands.SOSRemove               'SOSDONE
            Call HandleSOSRemove(UserIndex)
        
        Case eGMCommands.GoToChar                '/IRA
            Call HandleGoToChar(UserIndex)
        
        Case eGMCommands.invisible               '/INVISIBLE
            Call HandleInvisible(UserIndex)
        
        Case eGMCommands.GMPanel                 '/PANELGM
            Call HandleGMPanel(UserIndex)
            
        Case eGMCommands.HacerTorneo
            Call HandleHacerTorneo(UserIndex)
               
        Case eGMCommands.Cerrartorneo
            Call HandleCerrarTorneo(UserIndex)
            
        Case eGMCommands.GlobalOn                '/GLOBALON
            Call HandleGlobalOn(UserIndex)
           
        Case eGMCommands.GlobalOff               '/GLOBALOFF
            Call HandleGlobalOff(UserIndex)
            
        Case eGMCommands.Jail                    '/CARCEL
            Call HandleJail(UserIndex)
            
    End Select
End With

Exit Sub

errHandler:
    Call LogError("Error en GMCommands. Error: " & Err.Number & " - " & Err.Description & ". Paquete: " & Command)

End Sub
''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim PJ As String
    Dim nombre As String
    Dim accName As String
    Dim Password As String
    Dim version As String
    Dim SerialHD As Long
    Dim CuentaId As Integer
    
    accName = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()
    PJ = buffer.ReadASCIIString()
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    'Serial del HD
    SerialHD = val(SDesencriptar(buffer.ReadASCIIString()))
    
    '02/11/2015 Irongete: Cargo los datos de la cuenta.
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id,password,mail FROM cuenta WHERE mail = '" & accName & "' AND password = '" & Password & "'")
    
    '02/11/2015 Irongete: Compruebo si existe la cuenta.
    If RS.RecordCount = 0 Then
        Call WriteErrorMsg(UserIndex, "La cuenta no existe o la contraseña es incorrecta.")
        Exit Sub
    End If
        
    CuentaId = CInt(RS!id)
    
    If Not CuentaId > 0 Then
        Call WriteErrorMsg(UserIndex, "La cuenta no existe o la contraseña es incorrecta.")
        Exit Sub
    End If
        
    '02/11/2015 Irongete: Compruebo si la contraseña es correcta.
    If Not Password = CStr(RS!Password) Then
        Call WriteErrorMsg(UserIndex, "La cuenta no existe o la contraseña es incorrecta.")
        Exit Sub
    End If
    
    RS.Close
    
    '02/11/2015 Irongete: Cargo la informacion del personaje
    Dim i As Integer
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id_cuenta,nombre FROM personaje WHERE nombre = '" & PJ & "'")
     
    nombre = RS!nombre
    
    If Not AsciiValidos(nombre) Then
        Call WriteErrorMsg(UserIndex, "Nombre invalido.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If RS.EOF = True Or RS.BOF = True Or Len(nombre) = 0 Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        RS.Close
        Exit Sub
    End If
    
    '¿El personaje que intenta borrar ya esta borrado?
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("Select borrado FROM personaje WHERE nombre = '" & nombre & "'")
    
    If CByte(RS!borrado) = 1 Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        RS.Close
        Exit Sub
    End If
    
    UserList(UserIndex).CuentaId = CuentaId
    RS.Close
    
    If Not CuentaId = UserList(UserIndex).CuentaId Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    If BANCheck(nombre) Or BanHD_find(SerialHD) > 0 Then
        Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a AODrag a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.aodrag.es")
    ElseIf Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.aodrag.es")
    Else
        Call ConnectUser(UserIndex, nombre, SerialHD, UserList(UserIndex).CuentaId)
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 15 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    Dim version As String
    Dim race As eRaza
    Dim gender As eGenero
    Dim Class As eClass
    Dim CuentaName As String
    Dim Head As Integer
    Dim SerialHD As String
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    UserName = buffer.ReadASCIIString()
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    ' Serial del HD
    SerialHD = val(SDesencriptar(buffer.ReadASCIIString())) ' GSZAO
    
    race = buffer.ReadByte()
    gender = buffer.ReadByte()
    Class = buffer.ReadByte()
    CuentaName = buffer.ReadASCIIString()
    Head = buffer.ReadInteger()
    
   
    If BanHD_find(SerialHD) > 0 Then ' GSZAO
        Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a AODrag debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.aodrag.es")
    ElseIf Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en la web oficial del juego")
    Else
        Call ConnectNewUser(UserIndex, UserName, race, gender, Class, CuentaName, Head, SerialHD)
    End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 23/09/2009
'15/07/2009: ZaMa - Now invisible admins talk by console.
'23/09/2009: ZaMa - Now invisible admins can't send empty chat.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Dijo: " & Chat)
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.invisible = 0 Then
                Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If Not (.flags.AdminInvisible = 1) Then
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor))
                End If
            Else
                If RTrim(Chat) <> "" Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 15/07/2009
'15/07/2009: ZaMa - Now invisible admins yell by console.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        

        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Grito: " & Chat)
        End If
            
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.invisible = 0 Then
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
            
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
                
            If .flags.Privilegios And PlayerType.User Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, vbRed))
                End If
            Else
                If Not (.flags.AdminInvisible = 1) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                End If
            End If
        End If
        
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 15/07/2009
'28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
'15/07/2009: ZaMa - Now invisible admins wisper by console.
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        Dim targetCharIndex As String
        Dim targetUserIndex As Integer
        Dim targetPriv As PlayerType
        
        targetCharIndex = buffer.ReadASCIIString()
        Chat = buffer.ReadASCIIString()
        
        If UCase(.Name) = targetCharIndex Then
            Call WriteConsoleMsg(UserIndex, "¡No puedes enviarte privados a ti mismo!", FontTypeNames.FONTTYPE_INFO)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub
        End If
        
        targetUserIndex = NameIndex(targetCharIndex)
        
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub
        End If

        If targetUserIndex <= 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario no existe o no se encuentra online.", FontTypeNames.FONTTYPE_INFO)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub
        End If
            
        targetPriv = UserList(targetUserIndex).flags.Privilegios
        'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
        If (targetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Then
            ' Controlamos que no este invisible
            If UserList(targetUserIndex).flags.AdminInvisible <> 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes enviar mensajes privados a los Dioses y Admins.", FontTypeNames.FONTTYPE_INFO)
                Call .incomingData.CopyBuffer(buffer)
                Exit Sub
                'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
            ElseIf (.flags.Privilegios And PlayerType.User) <> 0 And (Not targetPriv And PlayerType.User) <> 0 Then
                ' Controlamos que no este invisible
                If UserList(targetUserIndex).flags.AdminInvisible <> 1 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes enviar mensajes privados a los GMs.", FontTypeNames.FONTTYPE_INFO)
                    Call .incomingData.CopyBuffer(buffer)
                    Exit Sub
                End If
            End If
        End If
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Le dijo a '" & UserList(targetUserIndex).Name & "' " & Chat)
        End If
                    
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
                        
            If Not (.flags.AdminInvisible = 1) Then
            
                If Not EstaPCarea(UserIndex, targetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, .Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                    Call WriteConsoleMsg(targetUserIndex, .Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                Else
                    Call WriteChatOverHead(UserIndex, Chat, .Char.CharIndex, 16744703)
                    Call WriteChatOverHead(targetUserIndex, Chat, .Char.CharIndex, 16744703)
                End If
                
                Call FlushBuffer(targetUserIndex)
            
                '[CDT 17-02-2004]
                If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                    Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("A " & UserList(targetUserIndex).Name & "> " & Chat, .Char.CharIndex, vbYellow))
                End If
            Else
                Call WriteConsoleMsg(UserIndex, .Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                If UserIndex <> targetUserIndex Then Call WriteConsoleMsg(targetUserIndex, .Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                            
                If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                    Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg(.Name & "> " & UserList(targetUserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO))
                End If
            End If
       End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim dummy As Long
    Dim TempTick As Long
    Dim heading As eHeading
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        heading = .incomingData.ReadByte()
        
        'If UserList(UserIndex).Lac.LCaminar.Puedo = False Then Exit Sub
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 40 + (.flags.Speed * 2) Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = getInterval(TempTick, .flags.StartWalk) ' 0.13.5
            
            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If getInterval(TempTick, .flags.CountSH) > 30000 Then ' 0.13.5
                    .flags.CountSH = 0
                End If
        
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then dummy = 126000 \ dummy
         
                    Call LogHackAttemp("SpeedHack: " & .Name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de SpeedHack.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(UserIndex)
         
                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        '¿Está trabajando?
        If .flags.Makro <> 0 Then
            Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
            .flags.Makro = 0
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0
                
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Else
                'Move user
                Call MoveUserChar(UserIndex, heading)

                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteMultiMessage(UserIndex, eMessages.Paralizado)
            End If
            
            .flags.CountSH = 0
        End If
        
    End With
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call WritePosUpdate(UserIndex)
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/01/08
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If UserList(UserIndex).Lac.LPegar.Puedo = False Then Exit Sub
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        
        '¿Está trabajando?
        If .flags.Makro <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Call WriteConsoleMsg(UserIndex, "¡Estas meditando!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No podés usar así esta arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
    End With
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'02/26/2006: Marco - Agregué un checkeo por si el usuario trata de agarrar un item mientras comercia.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then Exit Sub
        
        'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
        If .flags.Comerciando Then Exit Sub
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(UserIndex, "No puedes tomar ningun objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call GetObj(UserIndex)
    End With
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Seguro = 1 Then
            Call WriteSegNoti(UserIndex, 1, False)
            .flags.Seguro = 0
        Else
            Call WriteSegNoti(UserIndex, 1, True)
            .flags.Seguro = 1
        End If
        
    End With
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'***************************************************
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        .flags.SeguroResu = Not .flags.SeguroResu
        
        If .flags.SeguroResu Then
            Call WriteConsoleMsg(UserIndex, "Seguro de resurrección activado.", FontTypeNames.FONTTYPE_INFO)
            Call WriteSegNoti(UserIndex, 0, True)
        Else
            Call WriteConsoleMsg(UserIndex, "Seguro de resurrección desactivado.", FontTypeNames.FONTTYPE_INFO)
            Call WriteSegNoti(UserIndex, 0, False)
        End If
    End With
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call modGuilds.SendGuildLeaderInfo(UserIndex)
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAttributes(UserIndex)
End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call EnviarFama(UserIndex)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteSendSkills(UserIndex)
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteMiniStats(UserIndex)
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 And UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
            Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(.ComUsu.DestUsu)
            
            'Send data in the outgoing buffer of the other user
            Call FlushBuffer(.ComUsu.DestUsu)
        End If
        
        Call FinComerciarUsu(UserIndex)
    End With
End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)
    End With
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim otherUser As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)
            End If
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
    End With
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'07/25/09: Marco - Agregué un checkeo para patear a los usuarios que tiran items mientras comercian.
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Slot As Byte
    Dim Amount As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
    
        '<Edurne> 'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or _
            .flags.QueMontura Or _
            .flags.Muerto = 1 Or _
            ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 1000 Then
                Call WriteConsoleMsg(UserIndex, "¡La cantidad minima que puedes tirar es de 1k!.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub 'Don't drop too much gold
            End If
            Call TirarOro(Amount, UserIndex)
            
            Call WriteUpdateGold(UserIndex)
        Else
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                
                If Not MapInfo(.Pos.Map).Pk And iniTirarOBJZonaSegura = False Then
                    Call WriteConsoleMsg(UserIndex, "¡¡No está permitido arrojar objetos al suelo en zonas seguras!!.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Nosetira = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes tirar ese objeto!", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
                
                Call DropObj(UserIndex, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)
            End If
        End If
    End With
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        
        Spell = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Muerto)
            Exit Sub
        End If
        
        '¿Está trabajando?
        If .flags.Makro <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        .flags.Hechizo = Spell
        
        If .flags.Hechizo < 1 Then
            .flags.Hechizo = 0
        ElseIf .flags.Hechizo > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
        End If
        
    End With
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call ClickIzquierdo(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    End With
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call ClickDerecho(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    End With
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case Robar, Magia, Domar
                Call WriteWorkRequestTarget(UserIndex, Skill)
            Case Ocultarse
                    'No usar invi mapas InviSinEfecto
                If permiso_en_zona(UserIndex) And permiso_zona.no_invisibilidad Then
                    Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
                If .flags.Navegando = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                If .clase = eClass.Hunter Or .clase = eClass.Warrior Then
                    Call DoOcultarse(UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, "Tu clase no puede ocultarse.", FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
    End With
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte()
        
        If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        End If
                
        Call UseInvItem(UserIndex, Slot)
    End With
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim cantidad As Integer
        
        Item = .ReadInteger()
        cantidad = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
        Call HerreroConstruirItem(UserIndex, Item, cantidad)
    End With
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim cantidad As Integer
        
        Item = .ReadInteger()
        cantidad = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        Call CarpinteroConstruirItem(UserIndex, Item, cantidad)
    End With
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        Dim Skill As eSkill
        Dim DummyINT As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Skill = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
        Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case eSkill.Proyectiles

                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                If UserList(UserIndex).Lac.LLanzar.Puedo = False Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes lanzar hechizos tan rápido!.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                        
                'Make sure the item is valid and there is ammo equipped.
                With .Invent
                    If .WeaponEqpObjIndex = 0 Then
                        DummyINT = 1
                    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                        DummyINT = 1
                    ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                        DummyINT = 1
                    ElseIf .MunicionEqpObjIndex = 0 Then
                        DummyINT = 1
                    ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
                        DummyINT = 2
                    ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                        DummyINT = 1
                    ElseIf .Object(.MunicionEqpSlot).Amount < 1 Then
                        DummyINT = 1
                    End If
                    
                    If DummyINT <> 0 Then
                        If DummyINT = 1 Then
                            Call WriteConsoleMsg(UserIndex, "No tenés municiones.", FontTypeNames.FONTTYPE_INFO)
                            
                            Call Desequipar(UserIndex, .WeaponEqpSlot)
                        End If
                        
                        Call Desequipar(UserIndex, .MunicionEqpSlot)
                        Exit Sub
                    End If
                End With
                
                'Quitamos stamina
                If .Stats.MinSta >= 10 Then
                    Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                    If .genero = eGenero.Hombre Then
                        Call WriteConsoleMsg(UserIndex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Estás muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    Exit Sub
                End If
                
                Call ClickIzquierdo(UserIndex, .Pos.Map, X, Y)
                
                tU = .flags.targetUser
                tN = .flags.TargetNPC
                
                'Validate target
                If tU > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                        Exit Sub
                    End If
                    
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        Call WriteConsoleMsg(UserIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Attack!
                    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    Call UsuarioAtacaUsuario(UserIndex, tU)
                    
                ElseIf tN > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(NPCList(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(NPCList(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                        Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                        Exit Sub
                    End If
                    
                    'Is it attackable???
                    If NPCList(tN).Attackable <> 0 Then
                        
                        'Attack!
                        Call UsuarioAtacaNpc(UserIndex, tN)
                    End If
                End If
                
                If tN > 0 Or tU > 0 Then
                    With .Invent
                        DummyINT = .MunicionEqpSlot
                        
                        'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                        Call QuitarUserInvItem(UserIndex, DummyINT, 1)
                        
                        If .Object(DummyINT).Amount > 0 Then
                            'QuitarUserInvItem unequipps the ammo, so we equip it again
                            .MunicionEqpSlot = DummyINT
                            .MunicionEqpObjIndex = .Object(DummyINT).ObjIndex
                            .Object(DummyINT).Equipped = 1
                        Else
                            .MunicionEqpSlot = 0
                            .MunicionEqpObjIndex = 0
                        End If
                        Call UpdateUserInv(False, UserIndex, DummyINT)
                    End With
                End If
                '-----------------------------------
            
            Case eSkill.Magia
                
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call ClickIzquierdo(UserIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub

        
                'Check Spell-Hit interval
                If Not IntervaloPermiteGolpeMagia(UserIndex) Then
                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                        Exit Sub
                    End If
                End If
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, "¡No puedes lanzar hechizos tan rápido!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.pesca
                DummyINT = .Invent.WeaponEqpObjIndex
                If DummyINT = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(.Pos.Map, X, Y) Then
                    Select Case DummyINT
                    
                        Case CAÑA_PESCA
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                                Exit Sub
                            End If
                            If EsNewbie(UserIndex) = True And MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
                                Call WriteConsoleMsg(UserIndex, "Los Newbie no pueden pescar en zona insegura.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            .flags.Makro = eMakro.PESCAR
                            Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)
                                    
                        Case RED_PESCA
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                                Exit Sub
                            End If
                            
                            If EsNewbie(UserIndex) = True And MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
                                Call WriteConsoleMsg(UserIndex, "Los Newbie no pueden pescar en zona insegura.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            .flags.Makro = eMakro.PescarRed
                            Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)
                        
                        Case Else
                            Exit Sub    'Invalid item!
                    End Select
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, río o mar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Robar
                'Does the map allow us to steal here?
                If MapInfo(.Pos.Map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call ClickIzquierdo(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    tU = .flags.targetUser
                    
                    If tU > 0 And tU <> UserIndex Then
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                 If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                     Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                                     Exit Sub
                                 End If
                                 
                                 '17/09/02
                                 'Check the trigger
                                 If MapData(UserList(tU).Pos.Map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 Call DoRobar(UserIndex, tU)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "No a quien robarle!.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "¡No puedes robar en zonas seguras!.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call ClickIzquierdo(UserIndex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If NPCList(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                            Exit Sub
                        End If
                        
                        If LenB(NPCList(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(UserIndex, "No podés domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No podés domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ninguna criatura allí!.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                
                If PuedeLingotear(UserIndex) Then
                    .flags.Makro = eMakro.Lingotear
                    Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Herreria
                'Target wehatever is in that tile
                Call ClickIzquierdo(UserIndex, .Pos.Map, X, Y)
                
                If .flags.targetObj > 0 Then
                    If ObjData(.flags.targetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                        Call WriteShowBlacksmithForm(UserIndex)
                    ElseIf ObjData(.flags.targetObj).OBJType = eOBJType.otFragua Then
                        Call WriteShowReciclarHerreria(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque o fragua.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque o fragua.", FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/11/09
'05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
'***************************************************
    If UserList(UserIndex).incomingData.length < 9 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim GuildName As String
        Dim site As String
        Dim codex() As String
        Dim errorStr As String
        
        desc = buffer.ReadASCIIString()
        GuildName = Trim$(buffer.ReadASCIIString())
        site = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        If modGuilds.CrearNuevoClan(UserIndex, desc, GuildName, site, codex, .FundandoGuildAlineacion, errorStr) Then
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.Name & " ha fundado el clan: " & GuildName & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            If UserList(UserIndex).Stats.ELV < 40 Then
                Call LogGM("Historial de Canjes", .Name & " tenia: " & .Stats.DragCredits & " Dragcredotidos, y ahora tiene: " & .Stats.DragCredits - 50 & " al fundar el clan: " & GuildName)
                .Stats.DragCredits = .Stats.DragCredits - 50
            End If
            Call WriteUpdateUserStats(UserIndex)
            
            'Update tag
             Call RefreshCharStatus(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo.!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .nombre & vbCrLf _
                                               & "Descripción:" & .desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Mana necesario: " & .ManaRequerido & vbCrLf _
                                               & "Stamina necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
        
        itemSlot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub
        
        'Validate item slot
        If itemSlot > MAX_INVENTORY_SLOTS Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)
    End With
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/28/2008
'Last Modified By: NicoNZ
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim heading As eHeading
        Dim posX As Integer
        Dim posY As Integer
                
        heading = .incomingData.ReadByte()
        
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
            Select Case heading
                Case eHeading.NORTH
                    posY = -1
                Case eHeading.EAST
                    posX = 1
                Case eHeading.SOUTH
                    posY = 1
                Case eHeading.WEST
                    posX = -1
            End Select
            
                If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                    Exit Sub
                End If
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer
        Dim PetIndex As Byte
        
        PetIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If NPCList(.flags.TargetNPC).NPCType <> eNPCType.Entrenador Then Exit Sub
        
        If NPCList(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < NPCList(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(NPCList(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, NPCList(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    NPCList(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    NPCList(.flags.TargetNPC).Mascotas = NPCList(.flags.TargetNPC).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        If UserList(UserIndex).Lac.LUsar.Puedo = False Then Exit Sub 'Le pongo este mismo intervalo para evitar la compra muy rapida.
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        '¿El NPC puede comerciar?
        If NPCList(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        If UserList(UserIndex).Lac.LUsar.Puedo = False Then Exit Sub 'Le pongo este mismo intervalo para evitar la compra muy rapida.
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿Es el banquero?
        If NPCList(.flags.TargetNPC).NPCType <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User retira el item del slot
        Call UserRetiraItem(UserIndex, Slot, Amount)
    End With
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        If UserList(UserIndex).Lac.LUsar.Puedo = False Then Exit Sub 'Le pongo este mismo intervalo para evitar la compra muy rapida.
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If NPCList(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        If UserList(UserIndex).Lac.LUsar.Puedo = False Then Exit Sub 'Le pongo este mismo intervalo para evitar la compra muy rapida.
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If NPCList(.flags.TargetNPC).NPCType <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, Slot, Amount)
    End With
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Call DesplazarHechizo(UserIndex, dir, .ReadByte())
    End With
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        Dim Slot As Byte
        Dim TempItem As Obj
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Slot = .ReadByte()
    End With
        
    With UserList(UserIndex)
        TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount
        
        If dir = 1 Then 'Mover arriba
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            .BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
        Else 'mover abajo
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            .BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount
        End If
    End With
    
    Call UpdateBanUserInv(True, UserIndex, 0)
    Call UpdateVentanaBanco(UserIndex)

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim codex() As String
        
        desc = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(desc, codex, .GuildIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        'If amount is invalid, or slot is invalid and it's not gold, then ignore it.
        If ((Slot < 1 Or Slot > MAX_INVENTORY_SLOTS) And Slot <> FLAGORO) _
                        Or Amount <= 0 Then Exit Sub
        
        'Is the other player valid??
        If tUser < 1 Or tUser > MaxUsers Then Exit Sub
        
        'Is the commerce attempt valid??
        If UserList(tUser).ComUsu.DestUsu <> UserIndex Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        End If
        
        'Is he still logged??
        If Not UserList(tUser).flags.UserLogged Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else
            'Is he alive??
            If UserList(tUser).flags.Muerto = 1 Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            End If
            
            'Has he got enough??
            If Slot = FLAGORO Then
                'gold
                If Amount > .Stats.GLD Then
                    Call WriteMultiMessage(tUser, eMessages.NoCantidad)
                    Exit Sub
                End If
            Else
                'inventory
                If Amount > .Invent.Object(Slot).Amount Then
                    Call WriteMultiMessage(UserIndex, eMessages.NoCantidad)
                    Exit Sub
                End If
            End If
            
            'Prevent offer changes (otherwise people would ripp off other players)
            If .ComUsu.Objeto > 0 Then
                Call WriteMultiMessage(tUser, eMessages.NoCantidad)
                Exit Sub
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteConsoleMsg(UserIndex, "No podés vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            .ComUsu.Objeto = Slot
            .ComUsu.cant = Amount
            
            'If the other one had accepted, we turn that back and inform of the new offer (just to be cautious).
            If UserList(tUser).ComUsu.Acepto = True Then
                UserList(tUser).ComUsu.Acepto = False
                Call WriteConsoleMsg(tUser, .Name & " ha cambiado su oferta.", FontTypeNames.FONTTYPE_TALK)
            End If
            
            Call EnviarObjetoTransaccion(tUser)
        End If
    End With
End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim User As String
        Dim details As String
        
        User = buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherGuildIndex As Integer
        
        guild = buffer.ReadASCIIString()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()

        tUser = NameIndex(UserName)
        
        Set RS = New ADODB.Recordset
        
        '14/11/2015 Irongete: Actualizo el clan del jugador
        Set RS = SQL.Execute("UPDATE personaje SET id_clan = '" & .GuildIndex & "' WHERE nombre = '" & UserName & "'")
        
        '14/11/2015 Irongete: Borro las solicitudes de este jugador
        Set RS = SQL.Execute("SELECT id FROM personaje WHERE nombre = '" & UserName & "'")
        
        Set RS = SQL.Execute("DELETE FROM rel_clan_solicitud WHERE id_personaje = '" & RS!id & "'")

        'Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
        'Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim reason As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        Set RS = New ADODB.Recordset
        
        '14/11/2015 Irongete: Borro las solicitudes de este jugador
        Set RS = SQL.Execute("SELECT id FROM personaje WHERE nombre = '" & UserName & "'")
        
        Set RS = SQL.Execute("DELETE FROM rel_clan_solicitud WHERE id_personaje = '" & RS!id & "'")
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        Set RS = New ADODB.Recordset
        
        GuildIndex = UserList(UserIndex).GuildIndex
        UserName = buffer.ReadASCIIString()
        
        '14/11/2015 Irongete: Compruebo si es el lider del clan
        Set RS = SQL.Execute("SELECT nombre FROM clan WHERE id = '" & GuildIndex & "' AND lider = '" & UserList(UserIndex).id & "'")
        If RS.RecordCount > 0 And (Not UserName = UserList(UserIndex).Name) Then
        
           
            '14/11/2015 Irongete: Echo al personaje del clan
            Set RS = SQL.Execute("UPDATE personaje SET id_clan = '0' WHERE nombre = '" & UserName & "'")
            
            'Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            'Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        
        
        Else
        
         Call WriteConsoleMsg(UserIndex, "No puedes expulsar a este jugador.", FontTypeNames.FONTTYPE_GUILD)
        
        End If
     
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim error As String
        
        If Not modGuilds.v_AbrirElecciones(UserIndex, error) Then
            Call WriteConsoleMsg(UserIndex, error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim application As String
        Dim errorStr As String
        Dim clanid As Integer
        Dim Ahora As String
        Ahora = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & " " & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now)
        
        guild = buffer.ReadASCIIString()
        application = buffer.ReadASCIIString()
        
        Set RS = New ADODB.Recordset
        
        '14/11/2015 Irongete: Obtener id del clan
        Set RS = SQL.Execute("SELECT id FROM clan WHERE nombre = '" & guild & "'")
        clanid = RS!id
        Set RS = SQL.Execute("INSERT INTO rel_clan_solicitud (id_clan, id_personaje, fecha, mensaje) VALUES ('" & clanid & "', '" & .id & "' , '" & Ahora & "', '" & application & "')")
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        For i = 1 To LastUser
            If LenB(UserList(i).Name) <> 0 Then
                If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
                    count = count + 1
            End If
        Next i
        
        Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(count), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'***************************************************
    Dim tUser As Integer
    Dim isNotVisible As Boolean
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        If .flags.Morph = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando transformado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado. ", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)
        End If
        
        Call Cerrar_Usuario(UserIndex)
    End With
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim GuildIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .Name)
        
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(UserIndex, "Tu no puedes salir de ningún clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim earnings As Integer
    Dim percentage As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(NPCList(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        Select Case NPCList(.flags.TargetNPC).NPCType
            Case eNPCType.Banquero
                Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
        End Select
    End With
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(NPCList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If NPCList(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        NPCList(.flags.TargetNPC).Movement = TipoAI.Estatico
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre ál.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(NPCList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If NPCList(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(NPCList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        'Make sure it's the trainer
        If NPCList(.flags.TargetNPC).NPCType <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
    End With
End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comenzás a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/08 (NicoNZ)
'Arreglé un bug que mandaba un index de la meditacion diferente
'al que decia el server.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo podés usar meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If ManaMaxima(UserIndex) = 0 Then
             Call WriteConsoleMsg(UserIndex, "Sólo las clases mágicas conocen el arte de la meditación.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        '¿Está trabajando?
        If .flags.Makro <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = ManaMaxima(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Mana restaurado", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteUpdateMana(UserIndex)
            Exit Sub
        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            
            'Call WriteConsoleMsg(UserIndex, "Comienzas a meditar.", FontTypeNames.FONTTYPE_INFOBOLD)
            
            .Char.loops = INFINITE_LOOPS
            
            'Show proper FX according to level
            If .Stats.ELV < 10 Then
                .Char.FX = 46
            
            ElseIf .Stats.ELV < 20 Then
                .Char.FX = 3
            
            ElseIf .Stats.ELV < 30 Then
                .Char.FX = 49
                
            ElseIf .Stats.ELV < 40 Then
                .Char.FX = 40
            
            ElseIf .Stats.ELV < 45 Then
                .Char.FX = 11
           
            Else
                If criminal(UserIndex) = False Then
                    .Char.FX = 34
                Else
                    .Char.FX = 35
                End If
            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            'Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFOBOLD)
        End If
    End With
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (NPCList(.flags.TargetNPC).NPCType <> eNPCType.Revividor _
            And (NPCList(.flags.TargetNPC).NPCType <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) _
            Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, NPCList(.flags.TargetNPC).Pos) > 10 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        Call RevivirUsuario(UserIndex)
        Call WriteConsoleMsg(UserIndex, "El Sacerdote levanta las manos, pronuncia unas palabras sagradas y tu cuerpo comienza a tomar forma. ¡Has resucitado!", FontTypeNames.FONTTYPE_INFOBOLD)
    End With
End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (NPCList(.flags.TargetNPC).NPCType <> eNPCType.Revividor _
            And NPCList(.flags.TargetNPC).NPCType <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, NPCList(.flags.TargetNPC).Pos) > 10 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        If iniSacerdoteCuraVeneno = True Then ' GSZAO
            If .flags.Envenenado <> 0 Then
                .flags.Envenenado = 0
            End If
        End If
        
        .Stats.MinHP = VidaMaxima(UserIndex)
        
        Call WriteUpdateHP(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "El Sacerdote levanta las manos, pronuncia unas palabras sagradas y tus heridas comienzan a sanar rapidamente. ¡Has sido curado! ", FontTypeNames.FONTTYPE_INFOBOLD)
    End With
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendUserStatsTxt(UserIndex, UserIndex)
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendHelp(UserIndex)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            'Does the NPC want to trade??
            If NPCList(.flags.TargetNPC).Comercia = 0 Then
                If LenB(NPCList(.flags.TargetNPC).desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
            End If
            
            If Distancia(NPCList(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                Exit Sub
            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
        '[Alejo]
        ElseIf .flags.targetUser > 0 Then
            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Is the other one dead??
            If UserList(.flags.targetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it me??
            If .flags.targetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con vos mismo...", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Check distance
            If Distancia(UserList(.flags.targetUser).Pos, .Pos) > 3 Then
                Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                Exit Sub
            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.targetUser).flags.Comerciando = True And _
                UserList(.flags.targetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.targetUser
            .ComUsu.DestNick = UserList(.flags.targetUser).Name
            .ComUsu.cant = 0
            .ComUsu.Objeto = 0
            .ComUsu.Acepto = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.targetUser)
           
          ' Call WriteConsoleMsg(UserIndex, "El comercio entre usuarios se encuentra deshabilitado en estos momentos.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(NPCList(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                Exit Sub
            End If
            
            'If it's the banker....
            If NPCList(.flags.TargetNPC).NPCType = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If NPCList(.flags.TargetNPC).NPCType <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, NPCList(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If NPCList(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(UserIndex)
        Else
            Call EnlistarCaos(UserIndex)
        End If
    End With
End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If NPCList(.flags.TargetNPC).NPCType <> eNPCType.Noble _
                Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, NPCList(.flags.TargetNPC).Pos) > 4 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        If NPCList(.flags.TargetNPC).flags.Faccion = 0 Then
             If .Faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te daré una recompensa.", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             If .Faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te daré una recompensa.", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If NPCList(.flags.TargetNPC).NPCType <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub

        If Distancia(.Pos, NPCList(.flags.TargetNPC).Pos) > 4 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If

        If NPCList(.flags.TargetNPC).flags.Faccion = 0 Then
             If .Faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call ModFacciones.RecompensaArmadaReal(UserIndex)
        Else
             If .Faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call ModFacciones.RecompensaCaos(UserIndex)
        End If
    End With
End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendMOTD(UserIndex)
End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Dim time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    time = getInterval((GetTickCount() And &H7FFFFFFF), tInicioServer) \ 1000 ' 0.13.5
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    If time = 1 Then
        UpTimeStr = time & " día, " & UpTimeStr
    Else
        UpTimeStr = time & " días, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(UserIndex, "Servidor Online hace " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

''
' Handles the "PartyLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyLeave(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Call mdParty.SalirDeParty(UserIndex)
End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    ConsultaPopular.SendInfoEncuesta (UserIndex)
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 15/07/2009
'02/03/2009: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
'15/07/2009: ZaMa - Now invisible admins only speak by console
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & Chat))
                
                If Not (.flags.AdminInvisible = 1) Then _
                    Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, vbYellow, True))
                    Call BroadcastClan(UserList(UserIndex).GuildIndex, UserList(UserIndex).Name & ": " & Chat)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        Dim PartyId As Integer
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            'Compruebo que está en party
            PartyId = GetPersonajePartyId(UserIndex)
            If PartyId > 0 Then
            
                Call Drag_Party.BroadcastParty(PartyId, UserList(UserIndex).Name & ": " & Chat)
                
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call CentinelaCheckClave(UserIndex, .incomingData.ReadInteger())
    End With
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
        If .GuildIndex <> 0 Then
            Set RS = New ADODB.Recordset
            Set RS = SQL.Execute("SELECT nombre FROM personaje WHERE id_clan = '" & UserList(UserIndex).GuildIndex & "' AND logged = '1'")
            
            While Not RS.EOF
            
                onlineList = RS!nombre & ", " & onlineList
                RS.MoveNext
            Wend
            Call WriteConsoleMsg(UserIndex, "Compañeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(UserIndex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)
        End If
        
        
    End With
End Sub

''
' Handles the "PartyOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Call mdParty.OnlineParty(UserIndex)
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim request As String
        
        request = buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
Dim Categoria As String
Dim Mensaje As String
Dim cant As String

    Dim Ahora As String
    Ahora = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & " " & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Categoria = .incomingData.ReadASCIIString
        Mensaje = .incomingData.ReadASCIIString

        If Categoria = "Reporte de Bugs" Then
        
            Set RS = New ADODB.Recordset
            Set RS = SQL.Execute("INSERT INTO feedback (personaje, fecha, mensaje) VALUES ('" & UserList(UserIndex).id & "', '" & Ahora & "', '" & "BUG> " & Mensaje & "')")
        
            Call WriteConsoleMsg(UserIndex, "Gracias por tu mensaje.", FontTypeNames.FONTTYPE_INFO)
            
            Call WriteConsoleMsg(SendTarget.ToAdmins, "REPORTE DE BUG> " & Mensaje, FONTTYPE_CONSEJOCAOS)
            
        ElseIf Categoria = "Sugerencias" Then
            Set RS = New ADODB.Recordset
            Set RS = SQL.Execute("INSERT INTO feedback (personaje, fecha, mensaje) VALUES ('" & UserList(UserIndex).id & "', '" & Ahora & "', '" & "SUGERENCIA> " & Mensaje & "')")
        
            Call WriteConsoleMsg(UserIndex, "Gracias por tu mensaje.", FontTypeNames.FONTTYPE_INFO)
            
            Call WriteConsoleMsg(SendTarget.ToAdmins, "SUGERENCIAS> " & Mensaje, FONTTYPE_CONSEJO)
        End If
        
        'If Not Ayuda.Existe(.name) Then
            Call Ayuda.Push(.Name, Categoria, Mensaje)
        'Else
        '    Call Ayuda.Quitar(.name)
        '    Call Ayuda.Push(.name, Categoria, Mensaje)
        '    Call WriteConsoleMsg(UserIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        'End If
    End With
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Description As String
        
        Description = buffer.ReadASCIIString()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedés cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Else
            If Not AsciiValidos(Description) Then
                Call WriteConsoleMsg(UserIndex, "La descripción tiene caractéres inválidos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .desc = Trim$(Description)
                Call WriteConsoleMsg(UserIndex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If NPCList(.flags.TargetNPC).NPCType <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, NPCList(.flags.TargetNPC).Pos) > 10 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        If Amount > 0 And Amount <= .Stats.Banco Then
             .Stats.Banco = .Stats.Banco - Amount
             .Stats.GLD = .Stats.GLD + Amount
             Call WriteUpdateGLDBank(UserIndex, .Stats.Banco)
        Else
             Call WriteMultiMessage(UserIndex, eMessages.NoCantidad)
        End If
        
        Call WriteUpdateGold(UserIndex)
    End With
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If NPCList(.flags.TargetNPC).NPCType = eNPCType.Noble Then
           'Quit the Royal Army?
           If .Faccion.ArmadaReal = 1 Then
               If NPCList(.flags.TargetNPC).flags.Faccion = 0 Then
                   Call ExpulsarFaccionReal(UserIndex)
                   Call WriteChatOverHead(UserIndex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
               Else
                   Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
               End If
            'Quit the Chaos Legion??
           ElseIf .Faccion.FuerzasCaos = 1 Then
               If NPCList(.flags.TargetNPC).flags.Faccion = 1 Then
                   Call ExpulsarFaccionCaos(UserIndex, False)
                   Call WriteChatOverHead(UserIndex, "Ya volverás arrastrandote.", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
               Else
                   Call WriteChatOverHead(UserIndex, "Sal de aquí maldito criminal", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
               End If
           Else
               Call WriteChatOverHead(UserIndex, "¡No perteneces a ninguna facción!", NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite)
           End If
        End If
    End With
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(NPCList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        If NPCList(.flags.TargetNPC).NPCType <> eNPCType.Banquero Then Exit Sub
        
        If Amount > 0 And Amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.GLD = .Stats.GLD - Amount
            Call WriteUpdateGLDBank(UserIndex, .Stats.Banco)
            
            Call WriteUpdateGold(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.NoCantidad)
        End If
    End With
End Sub


''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim clanType As eClanType
        Dim error As String
        
        clanType = .incomingData.ReadByte()
        
        .FundandoGuildAlineacion = ALINEACION_NEUTRO
        
        'Select Case UCase$(Trim(clanType))
         '   Case eClanType.ct_RoyalArmy
          '      .FundandoGuildAlineacion = ALINEACION_ARMADA
           ' Case eClanType.ct_Evil
            '    .FundandoGuildAlineacion = ALINEACION_LEGION
            'Case eClanType.ct_Neutral
            '    .FundandoGuildAlineacion = ALINEACION_NEUTRO
            'Case eClanType.ct_GM
            '    .FundandoGuildAlineacion = ALINEACION_MASTER
            'Case eClanType.ct_Legal
            '    .FundandoGuildAlineacion = ALINEACION_CIUDA
            'Case eClanType.ct_Criminal
            '    .FundandoGuildAlineacion = ALINEACION_CRIMINAL
            'Case Else
            '    Call WriteConsoleMsg(userIndex, "Alineación inválida.", FontTypeNames.FONTTYPE_GUILD)
            '    Exit Sub
        'End Select
        
        If modGuilds.PuedeFundarUnClan(UserIndex, .FundandoGuildAlineacion, error) Then
            Call WriteShowGuildFundationForm(UserIndex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(UserIndex, error, FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub



''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim memberCount As Integer
        Dim i As Long
        Dim UserName As String
        
        guild = buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")
            End If
            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")
            End If
            
            If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(UserIndex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        
        message = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & message)
        
            If LenB(message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(message)
            
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)
        End If
    End With
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.ArmadaReal = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Armadas conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Armadas conectados", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.FuerzasCaos = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/07
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        Dim tIndex As Integer
        Dim X As Long
        Dim Y As Long
        Dim i As Long
        Dim found As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                                If MapData(UserList(tIndex).Pos.Map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                        Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                        found = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            
                            If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim comment As String
        comment = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.Name, "Hora.")
    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/Donde " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/07/06
'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1() As String
        Dim List2() As String
        
        Map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(Map) Then
            For i = 1 To LastNPC
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If NPCList(i).Pos.Map = Map Then
                    '¿esta vivo?
                    If NPCList(i).flags.Active And NPCList(i).Hostile = 1 And NPCList(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = NPCList(i).Name & ": (" & NPCList(i).Pos.X & "," & NPCList(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else
                            For j = 0 To NPCcount1 - 1
                                If Left$(List1(j), Len(NPCList(i).Name)) = NPCList(i).Name Then
                                    List1(j) = List1(j) & ", (" & NPCList(i).Pos.X & "," & NPCList(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = NPCList(i).Name & ": (" & NPCList(i).Pos.X & "," & NPCList(i).Pos.Y & ")"
                                NPCcant1(j) = 1
                            End If
                        End If
                    Else
                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = NPCList(i).Name & ": (" & NPCList(i).Pos.X & "," & NPCList(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else
                            For j = 0 To NPCcount2 - 1
                                If Left$(List2(j), Len(NPCList(i).Name)) = NPCList(i).Name Then
                                    List2(j) = List2(j) & ", (" & NPCList(i).Pos.X & "," & NPCList(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = NPCList(i).Name & ": (" & NPCList(i).Pos.X & "," & NPCList(i).Pos.Y & ")"
                                NPCcant2(j) = 1
                            End If
                        End If
                    End If
                End If
            Next i
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay más NPCS", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call LogGM(.Name, "Numero enemigos en mapa " & Map)
        End If
    End With
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/09
'26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Integer
        Dim Y As Integer
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        X = .flags.TargetX
        Y = .flags.TargetY
        
        Call FindLegalPos(UserIndex, .flags.TargetMap, X, Y)
        Call WarpUserChar(UserIndex, .flags.TargetMap, X, Y, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)
    End With
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Map As Integer
        Dim X As Integer
        Dim Y As Integer
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        Map = buffer.ReadInteger()
        X = buffer.ReadByte()
        Y = buffer.ReadByte()
        
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(Map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = UserIndex
                End If
            
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(Map, X, Y) Then
                    Call FindLegalPos(tUser, Map, X, Y)
                    Call WarpUserChar(tUser, Map, X, Y, True)
                    Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "ESTIMADO USUARIO, ud ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)
                
                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(UserIndex)
    End With
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then _
            Call Ayuda.Quitar(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
                    
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        Call FlushBuffer(tUser)
                    End If
                    
                    Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)
        Call LogGM(.Name, "/INVISIBLE")
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex)
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'***************************************************
    Dim i As Long
    Dim names() As String
    Dim count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        count = 1
        
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(count) = UserList(i).Name
                    count = count + 1
                End If
            End If
        Next i
        
        If count > 1 Then Call WriteUserNameList(UserIndex, names(), count - 1)
    End With
End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim Users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).Name
                
                ' Display the user being checked by the centinel
                If UserList(i).flags.CentinelaIndex <> 0 Then _
                    Users = Users & " (*)"
            End If
        Next i
        
        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim Users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).Name & ", "
            End If
        Next i
        
        If LenB(Users) <> 0 Then
            Users = Left$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim jailTime As Byte
        Dim count As Byte
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        jailTime = buffer.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(UserIndex, "No podés encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 60 Then
                        Call WriteConsoleMsg(UserIndex, "No podés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                        End If
                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                        End If
                        
                        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                            count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", count + 1)
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(reason) & " " & Date & " " & time)
                        End If
                        
                        Call Encarcelar(tUser, jailTime, .Name)
                        Call LogGM(.Name, " encarcelo a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/22/08 (NicoNZ)
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC As Integer
        Dim auxNPC As npc
        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & NPCList(tNPC).Name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = NPCList(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre el NPC antes", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim privs As PlayerType
        Dim count As Byte
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If Not privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No podés advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                    End If
                    If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                    End If
                    
                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(reason) & " " & Date & " " & time)
                        
                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName), FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/06/2009
'02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
'11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
'***************************************************
    If UserList(UserIndex).incomingData.length < 8 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim opcion As Byte
        Dim Arg1 As String
        Dim Arg2 As String
        Dim valido As Boolean
        Dim LoopC As Byte
        Dim CommandString As String
        Dim n As Byte
        Dim UserCharPath As String
        Dim Var As Long
        
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)
        End If
        
        opcion = buffer.ReadByte()
        Arg1 = buffer.ReadASCIIString()
        Arg2 = buffer.ReadASCIIString()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)
                Case PlayerType.Consejero
                    ' Los RMs consejeros sólo se pueden editar su head, body y level
                    valido = tUser = UserIndex And _
                            (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level)
                
                Case PlayerType.SemiDios
                    ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) _
                            Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level sólo lo puede hacer sobre sí mismo
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) Or _
                            opcion = eEditOptions.eo_Body Or _
                            opcion = eEditOptions.eo_Head Or _
                            opcion = eEditOptions.eo_CiticensKilled Or _
                            opcion = eEditOptions.eo_CriminalsKilled Or _
                            opcion = eEditOptions.eo_Class Or _
                            opcion = eEditOptions.eo_Skills Or _
                            opcion = eEditOptions.eo_addGold
            End Select
            
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then   'Si no es RM debe ser dios para poder usar este comando
            valido = True
        End If

        If valido Then
            UserCharPath = CharPath & UserName & ".chr"
            If tUser <= 0 And Not FileExist(UserCharPath) Then
                Call WriteConsoleMsg(UserIndex, "Esta intentando editar un usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.Name, "Intento editar un usuario inexistente.")
            Else
                'For making the Log
                CommandString = "/MOD "
                
                Select Case opcion
                    
                    Case eEditOptions.eo_Gold
                        If val(Arg1) < MAX_ORO_EDIT Then
                            If tUser <= 0 Then ' Esta offline?
                                Call WriteVar(UserCharPath, "STATS", "GLD", val(Arg1))
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.GLD = val(Arg1)
                                Call WriteUpdateGold(tUser)
                            End If
                        Else
                            If .flags.Privilegios = PlayerType.Admin Then
                                If tUser <= 0 Then ' Esta offline?
                                    Call WriteVar(UserCharPath, "STATS", "GLD", val(Arg1))
                                    Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                                Else ' Online
                                    UserList(tUser).Stats.GLD = val(Arg1)
                                    Call WriteUpdateGold(tUser)
                                End If
                            Else
                                Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores. Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "ORO "
                
                    Case eEditOptions.eo_Experience
                        If val(Arg1) > 20000000 Then
                                Arg1 = 20000000
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Var = GetVar(UserCharPath, "STATS", "EXP")
                            Call WriteVar(UserCharPath, "STATS", "EXP", Var + val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                            Call CheckUserLevel(tUser)
                            Call WriteUpdateExp(tUser)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "EXP "
                    
                    Case eEditOptions.eo_Body
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Body", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "BODY "
                    
                    Case eEditOptions.eo_Head
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Head", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, UserList(tUser).Char.body, val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "HEAD "
                    
                    Case eEditOptions.eo_CriminalsKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "FACCIONES", "CrimMatados", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.CriminalesMatados = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CRI "
                    
                    Case eEditOptions.eo_CiticensKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "FACCIONES", "CiudMatados", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.CiudadanosMatados = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CIU "
                    
                    Case eEditOptions.eo_Level
                        If val(Arg1) > STAT_MAXELV Then
                            Arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
                        End If
                        
                        ' Chequeamos si puede permanecer en el clan
                        If val(Arg1) >= 25 Then
                            
                            Dim GI As Integer
                            If tUser <= 0 Then
                                GI = GetVar(UserCharPath, "GUILD", "GUILDINDEX")
                            Else
                                GI = UserList(tUser).GuildIndex
                            End If
                            
                            If GI > 0 Then
                                If modGuilds.GuildAlignment(GI) = "Legión oscura" Or modGuilds.GuildAlignment(GI) = "Armada Real" Then
                                    'We get here, so guild has factionary alignment, we have to expulse the user
                                    Call modGuilds.m_EcharMiembroDeClan(-1, UserName)
                                    
                                    Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(UserName & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                                    ' Si esta online le avisamos
                                    If tUser > 0 Then _
                                        Call WriteConsoleMsg(tUser, "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la Facción bajo la cual tu clan está alineado, estarás excluído del mismo.", FontTypeNames.FONTTYPE_GUILD)
                                End If
                            End If
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "STATS", "ELV", val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.ELV = val(Arg1)
                            Call WriteUpdateUserStats(tUser)
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "LEVEL "
                    
                    Case eEditOptions.eo_Class
                        For LoopC = 1 To NUMCLASES
                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                            
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then ' Offline
                                Call WriteVar(UserCharPath, "INIT", "Clase", LoopC)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).clase = LoopC
                            End If
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "CLASE "
                        
                    Case eEditOptions.eo_Skills
                        For LoopC = 1 To NUMSKILLS
                            If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                        Next LoopC
                        
                        If LoopC > NUMSKILLS Then
                            Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then ' Offline
                                Call WriteVar(UserCharPath, "Skills", "SK" & LoopC, Arg2)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLS "
                    

                    Case eEditOptions.eo_Nobleza
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "REP", "Nobles", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Reputacion.NobleRep = Var
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "NOB "
                        
                    Case eEditOptions.eo_Asesino
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "REP", "Asesino", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Reputacion.AsesinoRep = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "ASE "
                    
                    Case eEditOptions.eo_Sex
                        Dim Sex As Byte
                        Sex = IIf(UCase(Arg1) = "MUJER", eGenero.Mujer, 0) ' Mujer?
                        Sex = IIf(UCase(Arg1) = "HOMBRE", eGenero.Hombre, Sex) ' Hombre?
                        
                        If Sex <> 0 Then ' Es Hombre o mujer?
                            If tUser <= 0 Then ' OffLine
                                Call WriteVar(UserCharPath, "INIT", "Genero", Sex)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).genero = Sex
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SEX "
                    
                    Case eEditOptions.eo_Raza
                        Dim raza As Byte
                        
                        Arg1 = UCase$(Arg1)
                        Select Case Arg1
                            Case "HUMANO"
                                raza = eRaza.Humano
                            Case "ELFO"
                                raza = eRaza.Elfo
                            Case "DROW"
                                raza = eRaza.Drow
                            Case "ENANO"
                                raza = eRaza.Enano
                            Case "GNOMO"
                                raza = eRaza.Gnomo
                            Case "ORCO"
                                raza = eRaza.Orco
                            Case "NO-MUERTO"
                                raza = eRaza.NoMuerto
                            Case Else
                                raza = 0
                        End Select
                        
                            
                        If raza = 0 Then
                            Call WriteConsoleMsg(UserIndex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "INIT", "Raza", raza)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).raza = raza
                            End If
                        End If
                            
                        ' Log it
                        CommandString = CommandString & "RAZA "
                        
                    Case eEditOptions.eo_addGold
                    
                        Dim bankGold As Long
                        
                        If Abs(Arg1) > MAX_ORO_EDIT Then
                            Call WriteConsoleMsg(UserIndex, "No está permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                bankGold = GetVar(CharPath & UserName & ".chr", "STATS", "BANCO")
                                Call WriteVar(UserCharPath, "STATS", "BANCO", IIf(bankGold + val(Arg1) <= 0, 0, bankGold + val(Arg1)))
                                Call WriteConsoleMsg(UserIndex, "Se le ha agregado " & Arg1 & " monedas de oro a " & UserName & ".", FONTTYPE_TALK)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "AGREGAR "
                        
                    Case eEditOptions.eo_Speed
                        If val(Arg1) > 255 Then
                                Arg1 = 255
                        End If
                        
                        UserList(UserIndex).flags.Speed = Arg1
                        Call WriteChangeSpeed(UserIndex, UserList(UserIndex).Char.CharIndex, UserList(UserIndex).flags.Speed)
                        
                    Case Else
                        Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                        CommandString = CommandString & "UNKOWN "
                        
                End Select
                
                CommandString = CommandString & Arg1 & " " & Arg2
                Call LogGM(.Name, CommandString & " " & UserName)
                
            End If
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Last Modification by: (liquid).. alto bug zapallo..
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
                
        Dim targetName As String
        Dim targetIndex As Integer
        
        targetName = Replace$(buffer.ReadASCIIString(), "+", " ")
        targetIndex = NameIndex(targetName)
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            'is the player offline?
            If targetIndex <= 0 Then
                'don't allow to retrieve administrator's info
                If Not (EsDios(targetName) Or EsAdmin(targetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, Buscando en Charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(UserIndex, targetName)
                End If
            Else
                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, targetIndex)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call LogGM(.Name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo Charfile... ", FontTypeNames.FONTTYPE_INFO)
                
                Call SendUserMiniStatsTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserMiniStatsTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BAL " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserOROTxtFromChar(UserIndex, UserName)
            Else
                Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserInvTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserInvTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BOV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserBovedaTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserBovedaTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim message As String
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)
            
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                For LoopC = 1 To NUMSKILLS
                    message = message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC

            Else
                Call SendUserSkillsTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex
            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        Call DarCuerpoDesnudo(tUser)
                        
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    .Stats.MinHP = VidaMaxima(UserIndex)
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call FlushBuffer(tUser)
                
                Call LogGM(.Name, "Resucito a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 12/28/06
'
'***************************************************
    Dim i As Long
    Dim list As String
    Dim priv As PlayerType
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then _
                    list = list & UserList(i).Name & ", "
            End If
        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 23/03/2009
'23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Map As Integer
        Map = .incomingData.ReadInteger
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Dim LoopC As Long
        Dim list As String
        Dim priv As PlayerType
        
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = Map Then
                If UserList(LoopC).flags.Privilegios And priv Then _
                    list = list & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.Name, "Intento perdonar un personaje de nivel avanzado.")
                    Call WriteConsoleMsg(UserIndex, "Solo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Rank As Integer
        
        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(UserIndex, "No podes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " echo a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Echo a " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "Estás loco?? como vas a piñatear un gm!!!! :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName, FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & UserName)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No está online", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, UserName, reason)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +)", FontTypeNames.FONTTYPE_INFO)
            Else
                If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
                    Call UnBan(UserName)
                
                    'penas
                    cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & time)
                
                    Call LogGM(.Name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no esta baneado. Imposible unbanear", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            NPCList(.flags.TargetNPC).flags.Inmovilizado = 0
            NPCList(.flags.TargetNPC).flags.Paralizado = 0
            NPCList(.flags.TargetNPC).Contadores.Paralisis = 0
        End If
    End With
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                  (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True)
                    Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call EnviarSpawnList(UserIndex)
    End With
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim npc As Integer
        npc = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then _
              Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            
            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
        End If
    End With
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & NPCList(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LimpiarMundo
    End With
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & "> " & message, FontTypeNames.FONTTYPE_TALK))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 24/07/07
'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim priv As PlayerType
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User
            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                    Dim ip As String
                    Dim lista As String
                    Dim LoopC As Long
                    ip = UserList(tUser).ip
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).ip = ip Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).Name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No hay ningun personaje con ese nick", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim ip As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv As PlayerType
        
        ip = .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & ip)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User
        End If

        For LoopC = 1 To LastUser
            If UserList(LoopC).ip = ip Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).Name & ", "
                    End If
                End If
            End If
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim tGuild As Integer
        
        GuildName = buffer.ReadASCIIString()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & _
                  modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.Name, "/CT " & Mapa & "," & X & "," & Y)
        
        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then _
            Exit Sub
        
        If MapData(Mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ET As Obj
        ET.Amount = 1
        ET.ObjIndex = 378
        
        Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)
        
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = Mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With
    End With
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        Dim Mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        With MapData(Mapa, X, Y)
            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                Call LogGM(UserList(UserIndex).Name, "/DT: " & Mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)
                End If
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim tUser As Integer
        Dim desc As String
        
        desc = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.targetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim midiID As Byte
        Dim Mapa As Integer
        
        midiID = .incomingData.ReadByte
        Mapa = .incomingData.ReadInteger
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.Map
            End If
        
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))
        End If
    End With
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim waveID As Byte
        Dim Mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        waveID = .incomingData.ReadByte()
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
        'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, X, Y) Then
                Mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y
            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, Y))
        End If
    End With
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("ARMADA REAL> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, NPCList(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < XMaxMapSize + 1 And Y < YMaxMapSize + 1 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).Name, "/MASSDEST")
    End With
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim tObj As Integer
        Dim lista As String
        Dim X As Long
        Dim Y As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
    End With
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SecurityIp.DumpTables
    End With
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, Echando de los consejos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte
        Dim tLog As String
        
        tTrigger = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'
'***************************************************
    Dim tTrigger As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger
        
        Call LogGM(.Name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, _
            "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y _
            , FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim lista As String
        Dim LoopC As Long
        
        Call LogGM(.Name, "/BANIPLIST")
        
        For LoopC = 1 To BanIPs.count
            lista = lista & BanIPs.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar
    End With
End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim cantMembers As Integer
        Dim LoopC As Long
        Dim member As String
        Dim count As Byte
        Dim tIndex As Integer
        Dim tFile As String
        
        GuildName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " banned al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)
                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)
                    End If
                    
                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & member & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
                    count = val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", count + 1)
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & count + 1, LCase$(.Name) & ": BAN AL CLAN: " & GuildName & " " & Date & " " & time)
                Next LoopC
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/02/09
'Agregado un CopyBuffer porque se producia un bucle
'inifito al intentar banear una ip ya baneada. (NicoNZ)
'07/02/09 Pato - Ahora no es posible saber si un gm está o no online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bannedIP As String
        Dim tUser As Integer
        Dim reason As String
        Dim i As Long
        
        ' Is it by ip??
        If buffer.ReadBoolean() Then
            bannedIP = buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte()
        Else
            tUser = NameIndex(buffer.ReadASCIIString())
            
            If tUser > 0 Then bannedIP = UserList(tUser).ip
        End If
        
        reason = buffer.ReadASCIIString()
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP & " por " & reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).ip = bannedIP Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "IP POR " & reason)
                            End If
                        End If
                    Next i
                End If
            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tObj As Integer
        tObj = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.Name, "/CI: " & tObj)
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map > 0 Then _
            Exit Sub
        
        If tObj < 1 Or tObj > NumObjDatas Then _
            Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(tObj).Name) = 0 Then Exit Sub
        
        Dim Objeto As Obj
        Call WriteConsoleMsg(UserIndex, "ATENCION: FUERON CREADOS ***100*** ITEMS!, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        
        Objeto.Amount = 100
        Objeto.ObjIndex = tObj
        Call MakeObj(Objeto, .Pos.Map, .Pos.X, .Pos.Y - 1)
    End With
End Sub
Private Sub HandleMovePj(ByVal UserIndex As Integer)
'***************************************************
'Author: Edurne
'Last Modification: 10/09/2015
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tStr As String, tIndex As Integer
        tStr = .incomingData.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        tIndex = NameIndex(tStr)
        If tIndex > 0 Then
            Call WarpUserChar(tIndex, 1, 30, 70, True)
            Call WriteConsoleMsg(UserIndex, "Personaje " & tStr & " movido a Nix.", FontTypeNames.FONTTYPE_GUILD)
        Else
            If FileExist(CharPath & UCase$(tStr) & ".chr") Then
                Call WriteVar(CharPath & UCase$(tStr) & ".chr", "INIT", "Position", CStr("1-30-70"))
                Call WriteConsoleMsg(UserIndex, "Personaje " & tStr & " movido a Nix desde ficha.", FontTypeNames.FONTTYPE_GUILD)
            Else
                Call WriteConsoleMsg(UserIndex, "Personaje " & tStr & " inexistente.", FontTypeNames.FONTTYPE_GUILD)
            End If
        End If
        
        Call LogGM(.Name, "/MOVER : " & tStr)

    End With
End Sub


''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.Name, "/DEST")
        
        If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call EraseObj(10000, .Pos.Map, .Pos.X, .Pos.Y)
    End With
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)
    
            If tUser > 0 Then
                UserList(tUser).Faccion.FuerzasCaos = 0
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECHO DE LA REAL A: " & UserName)
            
            If tUser > 0 Then
                UserList(tUser).Faccion.ArmadaReal = 0
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte
        midiID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast musica: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
    End With
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte
        waveID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 1/05/07
'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim punishment As Byte
        Dim NewText As String
        
        UserName = buffer.ReadASCIIString()
        punishment = buffer.ReadByte
        NewText = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call LogGM(.Name, " borro la pena: " & punishment & "-" & _
                      GetVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment) _
                      & " de " & UserName & " y la cambió por: " & NewText)
                    
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, LCase$(.Name) & ": <" & NewText & "> " & Date & " " & time)
                    
                    Call WriteConsoleMsg(UserIndex, "Pena Modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.Name, "/BLOQ")
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0
        End If
        
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)
    End With
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & NPCList(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < XMaxMapSize + 1 And Y < YMaxMapSize + 1 Then
                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.Name, "/MASSKILL")
    End With
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim lista As String
        Dim LoopC As Byte
        Dim priv As Integer
        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            End If
            
            If validCheck Then
                Call LogGM(.Name, "/LASTIP " & UserName)
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"
                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC
                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the user`s chat color
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim color As Long
        
        color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Ignore the user
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 09/09/2008 (NicoNZ)
'Check one Users Slot in Particular from Inventory
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer
        
        UserName = buffer.ReadASCIIString() 'Que UserName?
        Slot = buffer.ReadByte() 'Que Slot?
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
            tIndex = NameIndex(UserName)  'Que user index?
            
            Call LogGM(.Name, .Name & " Checkeo el slot " & Slot & " de " & UserName)
               
            If tIndex > 0 Then
                If Slot > 0 And Slot <= MAX_INVENTORY_SLOTS Then
                    If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                        Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Restart the game
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        
        'time and Time BUG!
        Call LogGM(.Name, .Name & " reinicio el mundo")
        
        Call ReiniciarServidor(True)
    End With
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the objects
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado a los objetos.")
        
        Call LoadOBJData
    End With
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the spells
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        
        Call CargarHechizos
    End With
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s INI
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los INITs.")
        
        If Not LoadSini Then
            MsgBox "Error al cargar el archivo Server.ini"
            Exit Sub
        End If
    End With
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s NPC
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
         
        Call LogGM(.Name, .Name & " ha recargado los NPCs.")
    
        Call CargaNpcsDat
    
        Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Kick all the chars that are online
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados
    End With
End Sub

''
' Handle the "Night" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNight(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        
        DeNoche = Not DeNoche
        
        Dim i As Long
        
        For i = 1 To NumUsers
            If UserList(i).flags.UserLogged And UserList(i).ConnID > -1 Then
                Call EnviarNoche(i)
            End If
        Next i
    End With
End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Show the server form
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Clean the SOS
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha borrado los SOS")
        
        Call Ayuda.Reset
    End With
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Save the characters
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado todos los chars")
        
        'Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios
    End With
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la información sobre el BackUp")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.Map).BackUp = 1
        Else
            MapInfo(.Pos.Map).BackUp = 0
        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).BackUp)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si es PK el mapa.")
        
        MapInfo(.Pos.Map).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si es Restringido el mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Restringir = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & MapInfo(.Pos.Map).Restringir, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar la Magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo(.Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'InviSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noinvi = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar Invisibilidad el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapInfo(.Pos.Map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'ResuSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noresu = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar Resucitar el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapInfo(.Pos.Map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion del Terreno del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Terreno = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & MapInfo(.Pos.Map).Terreno, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el Mapa", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion de la Zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Saves the map
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        
        Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Allows admins to read guild messages
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        
        guild = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, guild)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha hecho un backup")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete
    End With
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 10/08/2011 - ^[GS]^
'Activate or desactivate the Centinel
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        centinelaActivado = Not centinelaActivado
        
        Call ResetCentinelas ' 0.13.3
        
        If centinelaActivado Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
        Else
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
        End If
    End With
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user name
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        newName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj esta online, debe salir para el cambio", FontTypeNames.FONTTYPE_WARNING)
                Else
                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " es inexistente ", FontTypeNames.FONTTYPE_INFO)
                    Else
                        GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))
                        
                        If GuildIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")
                                
                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa", FontTypeNames.FONTTYPE_INFO)
                                
                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                                
                                Dim cantPenas As Byte
                                
                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.Name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & time)
                                
                                Call LogGM(.Name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneo a " & NPCList(NpcIndex).Name & " en mapa " & .Pos.Map)
        End If
    End With
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneo con respawn " & NPCList(NpcIndex).Name & " en mapa " & .Pos.Map)
        End If
    End With
End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/12/07
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)
    End With
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
        End If
    End With
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'Turns off the server
'***************************************************
    Dim handle As Integer
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & time & " server apagado por " & .Name & ". "
        
        Close #handle
        
        Unload frmMain
    End With
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/CONDEN " & UserName)
            
            tUser = NameIndex(UserName)
            If tUser > 0 Then _
                Call VolverCriminal(tUser)
        End If
                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/09/09
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Char As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call ResetFacciones(tUser)
            Else
                Char = CharPath & UserName & ".chr"
                
                If FileExist(Char, vbNormal) Then
                    Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(Char, "FACCIONES", "CiudMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "CrimMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingresó a ninguna Facción")
                    Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rArReal", 0)
                    Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rExReal", 0)
                    Call WriteVar(Char, "FACCIONES", "recCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "recReal", 0)
                    Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
                    Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJARCLAN " & UserName)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Request user mail
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim mail As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")
                
                Call WriteConsoleMsg(UserIndex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/29/06
'Send a message to all the users
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "SetMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 03/31/07
'Set the MOTD
'Modified by: Juan Martín Sotuyo Dodero (Maraxus)
'   - Fixed a bug that prevented from properly setting the new number of lines.
'   - Fixed a bug that caused the player to be kicked.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim newMOTD As String
        Dim auxiliaryString() As String
        Dim LoopC As Long
        
        newMOTD = buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con exito", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín sotuyo Dodero (Maraxus)
'Last Modification: 12/29/06
'Change the MOTD
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub
        End If
        
        Dim auxiliaryString As String
        Dim LoopC As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)
    End With
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 06/08/2012 - ^[GS]^
'Show guilds messages
'***************************************************

    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim TActual As Long
        TActual = GetTickCount() And &H7FFFFFFF
        If getInterval(TActual, .Counters.TimerPuedeSendPing) < TIEMPO_SEND_PING Then ' GSZAO
            ' Es un cheater!!
            'Call Cerrar_Usuario(UserIndex)
            'Exit Sub
        End If
        .Counters.TimerPuedeSendPing = GetTickCount() And &H7FFFFFFF
        
        Call WritePong(UserIndex)
        Call FlushBuffer(UserIndex)
    End With
End Sub
Private Sub HandlePuntosClanes(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        
        Call WriteShowPuntosClanes(UserIndex)
    End With
 
End Sub

Private Sub HandleIrAlCementerio(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        With UserList(UserIndex)
        
            If .flags.Muerto = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡¡No estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.ArenaRinkel = True Then Call modArenaRinkel.SalirArenaRinkel(UserIndex)
        
        Select Case .Pos.Map
            Case 1, 2, 3, 4, 6, 9, 12, 13, 14
                Call WarpUserChar(UserIndex, 1, 31, 88, True) 'Cementerio de NIX
            Case 7, 5, 8, 11, 19
                Call WarpUserChar(UserIndex, 7, 66, 28, True) 'Cementerio de BANDER
            Case Else
                Call WarpUserChar(UserIndex, 1, 31, 88, True) 'mapa no programado = NIX
        End Select
        
        End With
    End With
 
End Sub

Public Sub HandleCastillosCMD(ByVal UserIndex As Integer)

     With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
                
    End With
    
    '15/03/2016 Lorwik
    Call DueñosCastillos(UserIndex)
       
End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal UserIndex As Integer)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/09
'Modify server.ini
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim sLlave As String
        Dim sClave As String
        Dim sValor As String
        
        'Obtengo los parámetros
        sLlave = buffer.ReadASCIIString()
        sClave = buffer.ReadASCIIString()
        sValor = buffer.ReadASCIIString()
    
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim sTmp As String
            'Obtengo el valor según llave y clave
            sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)
            
            'Si obtengo un valor escribo en el server.ini
            If LenB(sTmp) Then
                Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                Call LogGM(.Name, "Modificó en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                Call WriteConsoleMsg(UserIndex, "Modificó " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Logged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.logged)
        Call .outgoingData.WriteASCIIString(modGuilds.GuildName(.GuildIndex))
        Call .outgoingData.WriteASCIIString(.Name)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageCreateRenderValue(ByVal X As Byte, ByVal Y As Byte, ByVal rValue As String, ByVal rType As Byte)
'***************************************************
'Author: maTih.-
'Last Modification: 09/06/2012 - ^[GS]^
'***************************************************

    ' @ Envia el paquete para crear un valor en el render
     
    With auxiliarBuffer
         .WriteByte ServerPacketID.CreateRenderText
         .WriteByte X
         .WriteByte Y
         .WriteASCIIString (rValue)
         .WriteByte rType
         
         PrepareMessageCreateRenderValue = .ReadASCIIStringFixed(.length)
         
    End With
     
End Function

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call WriteUpdateGLDBank(UserIndex, UserList(UserIndex).Stats.Banco)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowReciclarHerreria(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 04/03/2016
'Manda a mostrar el form de reciclaje de herreria.
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowReciclarHerreria)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCSwing(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NPCSwing)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCKillUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NPCKillUser)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldUser)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldOther)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserSwing(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserSwing)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NobilityLost" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNobilityLost(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NobilityLost" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NobilityLost)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CantUseWhileMeditating)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHP)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal version As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
        Call .WriteInteger(version)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal damage As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCHitUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.NPCHitUser)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHitNPC(ByVal UserIndex As Integer, ByVal damage As Long, ByVal LeQuda As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHitNPC" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHitNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call .WriteLong(damage)
        Call .WriteASCIIString(LeQuda)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserAttackedSwing(ByVal UserIndex As Integer, ByVal attackerIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserAttackedSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserAttackedSwing)
        Call .WriteInteger(UserList(attackerIndex).Char.CharIndex)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long, Optional ByVal NoConsole = False)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatOverHead" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(Chat, CharIndex, color, NoConsole))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal Chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildChat" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(Chat))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal klan As String, ByVal criminal As Byte, _
                                ByVal privileges As Byte, ByVal bType As Byte, ByVal Ataque As Integer, ByVal Speed As Byte, ByVal PartyId As Integer, ByVal NPCType As Integer, ByVal NpcId As Integer, ByVal Quest As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                            helmet, Name, klan, criminal, privileges, bType, Ataque, Speed, PartyId, NPCType, NpcId, Quest))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterRemove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Writes the "ForceCharMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterChange" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteChangeSpeed(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Speed As Byte)
'***************************************************
'Author: Lorwik
'Last Modification: 16/12/2018
'Writes the "ChangeSpeed" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChangeSpeed(Speed, CharIndex))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal GrhIndex As Long, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectDelete" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockPosition" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayMidi" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim Tmp As String
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT nombre FROM clan ORDER BY id ASC")
        
        While Not RS.EOF
             Tmp = Tmp & RS!nombre & SEPARATOR
            RS.MoveNext
        Wend
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AreaChanged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PauseToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateFX" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(VidaMaxima(UserIndex))
        Call .WriteInteger(UserList(UserIndex).Stats.MinHP)
        Call .WriteInteger(ManaMaxima(UserIndex))
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
        Call .WriteInteger(UserList(UserIndex).Stats.DragCredits)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        'If ObjIndex = 0 Then Exit Sub
        
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(Slot).Amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(Slot).Equipped)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteSingle(SalePrice(obData.valor))
        Call .WriteByte(IIf(SexoPuedeUsarItem(UserIndex, ObjIndex) = True And FaccionPuedeUsarItem(UserIndex, ObjIndex) = True And ClasePuedeUsarItem(UserIndex, ObjIndex) = True, 1, 0))
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(Slot).Amount)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteLong(obData.valor)
        Call .WriteByte(IIf(SexoPuedeUsarItem(UserIndex, ObjIndex) = True And FaccionPuedeUsarItem(UserIndex, ObjIndex) = True And ClasePuedeUsarItem(UserIndex, ObjIndex) = True, 1, 0))
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 09/07/2012 - ^[GS]^
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot) ' slot
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(Slot)) ' nro de hechizo
        
        If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
            Call .WriteASCIIString(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).nombre) ' nombre
            Call .WriteInteger(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).GrhIndex) ' grhindex
        Else
            Call .WriteASCIIString("(None)")
            Call .WriteInteger(0)
        End If
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Atributes" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.atributes)
        Call .WriteByte(FuerzaMaxima(UserIndex))
        Call .WriteByte(AgilidadMaxima(UserIndex))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Energia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) Then
                count = count + 1
                validIndexes(count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) Then
                count = count + 1
                validIndexes(count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase) Then
                count = count + 1
                validIndexes(count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
        Next i
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RestOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSetPartyId(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal PartyId As Integer)
    On Error GoTo errHandler
    
    If CharIndex > 0 Then
        CharIndex = UserList(CharIndex).Char.CharIndex
    End If
    
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SetPartyId)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(PartyId)
    End With

    Exit Sub
    
errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ErrorMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Blind" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Dumb" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSignal" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(ObjIndex).texto)
        Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Obj As Obj, ByVal price As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/13/08
'Last Modified by: Nicolas Ezequiel Bouhid (NicoNZ)
'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim ObjInfo As ObjData
    
    If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.ObjIndex)
    End If

    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(ObjInfo.Name)
        Call .WriteInteger(Obj.Amount)
        Call .WriteSingle(price)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteInteger(Obj.ObjIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.MaxHIT)
        Call .WriteInteger(ObjInfo.MinHIT)
        Call .WriteInteger(ObjInfo.MaxDef)
        Call .WriteInteger(ObjInfo.MinDef)
        Call .WriteByte(IIf(SexoPuedeUsarItem(UserIndex, Obj.ObjIndex) = True And FaccionPuedeUsarItem(UserIndex, Obj.ObjIndex) = True And ClasePuedeUsarItem(UserIndex, Obj.ObjIndex) = True, 1, 0))
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Fame" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Fame)
        
        Call .WriteLong(UserList(UserIndex).Reputacion.AsesinoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BandidoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BurguesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.LadronesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.NobleRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.PlebeRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.Promedio)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MiniStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(UserIndex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.CriminalesMatados)
        
'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        Call .WriteLong(UserList(UserIndex).Stats.UsuariosMatados)
        
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(UserIndex).clase)
        Call .WriteLong(UserList(UserIndex).Counters.Pena)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowForumForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowForumForm)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MeditateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlindNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumbNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SendSkills" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendSkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserList(UserIndex).Stats.UserSkills(i))
        Next i
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To NPCList(NpcIndex).NroCriaturas
            str = str & NPCList(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then _
            str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal UserIndex As Integer, ByVal guildNews As String, ByRef enemies() As String, ByRef allies() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNews" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        
        Call .WriteASCIIString(guildNews)
        
        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Tmp = vbNullString
        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OfferDetails" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                            ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal reputation As Long, _
                            ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
                            ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(gold)
        Call .WriteLong(bank)
        Call .WriteLong(reputation)
        
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, ByRef guildList() As String, ByRef MemberList() As String, _
                            ByVal guildNews As String, ByRef joinRequests() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, _
                            ByVal leader As String, ByVal URL As String, ByVal memberCount As Integer, ByVal electionsOpen As Boolean, _
                            ByVal alignment As String, ByVal enemiesCount As Integer, ByVal AlliesCount As Integer, _
                            ByVal antifactionPoints As String, ByRef codex() As String, ByVal guildDesc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDetails" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim temp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)
        
        Call .WriteInteger(memberCount)
        Call .WriteBoolean(electionsOpen)
        
        Call .WriteASCIIString(alignment)
        
        Call .WriteInteger(enemiesCount)
        Call .WriteInteger(AlliesCount)
        
        Call .WriteASCIIString(antifactionPoints)
        
        For i = LBound(codex()) To UBound(codex())
            temp = temp & codex(i) & SEPARATOR
        Next i
        
        If Len(temp) > 1 Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
        
        Call .WriteASCIIString(guildDesc)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/12/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Writes the "ParalizeOK" message to the given user's outgoing data buffer
'And updates user position
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    Call WritePosUpdate(UserIndex)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TradeOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.TradeOK)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankOK)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(ObjData(ObjIndex).Name)
        Call .WriteLong(Amount)
        Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
        Call .WriteByte(ObjData(ObjIndex).OBJType)
        Call .WriteInteger(ObjData(ObjIndex).MaxHIT)
        Call .WriteInteger(ObjData(ObjIndex).MinHIT)
        Call .WriteInteger(ObjData(ObjIndex).MaxDef)
        Call .WriteInteger(ObjData(ObjIndex).MinDef)
        Call .WriteLong(SalePrice(ObjData(ObjIndex).valor))
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal UserIndex As Integer, ByVal night As Boolean)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Writes the "SendNight" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendNight)
        Call .WriteBoolean(night)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub




Public Sub WriteShowPuntosClanes(ByVal UserIndex As Integer)
    Dim Tmp As String
    Dim TotalClanes As Integer
    
    Dim GuildName As String
    Dim Puntos As Integer
    
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPuntosClanes)
        
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT j1.nombre AS nombre, rel_clan_puntos.puntoscastillos AS puntoscastillos FROM rel_clan_puntos JOIN clan j1 ON j1.id = rel_clan_puntos.id_clan ORDER BY puntoscastillos DESC")
        
        TotalClanes = RS.RecordCount
        
        While Not RS.EOF
             Tmp = Tmp & RS!nombre & "@" & RS!puntoscastillos & SEPARATOR
            
            RS.MoveNext
        Wend
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
    Exit Sub

End Sub


''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06 NIGO:
'Writes the "UserNameList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Pong" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
    Call FlushBuffer(UserIndex)
Exit Sub

errHandler:
    Debug.Print "error en ping"
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With UserList(UserIndex).outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call EnviarDatosASlot(UserIndex, sndData)
    End With
    'Debug.Print "flush buffer"
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, ByVal newNick As String) As String
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'Prepares the "Change Nick" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChangeNick)
        
        Call .WriteInteger(CharIndex)
        Call .WriteASCIIString(newNick)
        
        PrepareMessageCharacterChangeNick = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long, Optional ByVal NoConsole As Boolean = False) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ChatOverHead" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)
        
        Call .WriteBoolean(NoConsole)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal Chat As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal Chat As String) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)
    End With
End Function


''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMidi)
        Call .WriteByte(midi)
        Call .WriteInteger(loops)
        
        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PauseToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ObjectDelete" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "BlockPosition" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)
    End With
    
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Long, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'prepares the "ObjectCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteLong(GrhIndex)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterRemove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal klan As String, ByVal criminal As Byte, _
                                ByVal privileges As Byte, ByVal bType As Byte, ByVal Ataque As Integer, ByVal Speed As Byte, ByVal PartyId As Integer, ByVal NPCType As Integer, ByVal NpcId As Integer, ByVal Quest As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(Name)
        Call .WriteASCIIString(klan)
        Call .WriteInteger(PartyId)
        Call .WriteByte(criminal)
        Call .WriteInteger(NPCType)
        Call .WriteInteger(NpcId)
        Call .WriteByte(Quest)
        Call .WriteByte(privileges)
        Call .WriteByte(bType)
        Call .WriteInteger(Ataque)
        Call .WriteByte(Speed)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterChange" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageChangeSpeed(ByVal Speed As Byte, ByVal CharIndex As Integer) As String
'***************************************************
'Author: Lorwik
'Last Modification: 16/12/2018
'Prepares the "ChangeSpeed" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChangeSpeed)
        
        Call .WriteInteger(CharIndex)
        Call .WriteByte(Speed)
        
        PrepareMessageChangeSpeed = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Prepares the "ForceCharMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, isCriminal As Boolean, Tag As String) As String
'***************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modification: 04/07/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'Prepares the "UpdateTagAndStatus" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteBoolean(isCriminal)
        Call .WriteASCIIString(Tag)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub HandleNoche(ByVal UserIndex As Integer)
'**************************Lorwik/Noche**************************
'Todo el codigo fue reformado para reducirlo y optmizarlo.
Dim i As Long
With UserList(UserIndex)
'Leemos el paquete
Call .incomingData.ReadByte
        
DayStatus = .incomingData.ReadByte()
        
If .flags.Privilegios And PlayerType.User Then Exit Sub
        
    Select Case DayStatus
        Case 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Las luces del alba iluniman las tierras del Dragón.", FontTypeNames.FONTTYPE_GUILD))
            'Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(67, 0, 0))
            
        Case 1
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El sol se levanta hasta alcanzar su maximo explendor.", FontTypeNames.FONTTYPE_GUILD))
        
        Case 2
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El sol empieza a esconderse.", FontTypeNames.FONTTYPE_GUILD))
        
        Case 3
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("La noche cae sobre las tierras del Dragón. Las criaturas de la noche salen a cazar.", FontTypeNames.FONTTYPE_GUILD))
            'Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(81, 0, 0))
            
    End Select
    
        For i = 1 To LastUser
            Call WriteNoche(i, DayStatus)
        Next i
End With
End Sub
Public Function WriteNoche(ByVal UserIndex As Integer, ByVal Clima As Byte)
 '**************************Lorwik/Noche**************************
'Todo el codigo fue reformado para reducirlo y optmizarlo.
With UserList(UserIndex).outgoingData
    Call .WriteByte(ServerPacketID.Noche)
    Call .WriteByte(Clima)
End With
End Function

Public Sub HandleLoginCuenta(ByVal UserIndex As Integer)
Dim CuentaPass, CuentaEmail As String
Dim SerialHD As Long
Dim version As String

With UserList(UserIndex)
    'Remove Packet ID
    Call .incomingData.ReadByte
    
    'Convert version number to string
    version = CStr(.incomingData.ReadByte()) & "." & CStr(.incomingData.ReadByte()) & "." & CStr(.incomingData.ReadByte())
    CuentaEmail = .incomingData.ReadASCIIString
    CuentaPass = .incomingData.ReadASCIIString
    SerialHD = val(SDesencriptar(.incomingData.ReadASCIIString()))
    
    If Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.aodrag.es")
        Exit Sub
    Else
        Call ConectarCuenta(UserIndex, CuentaEmail, CuentaPass, SerialHD)
        Exit Sub
    End If
End With
End Sub
 
Public Sub HandleBorrarPJ(ByVal UserIndex As Integer)
Dim CuentaPass, CuentaEmail, NamePJ As String
Dim SerialHD As Long
Dim version As String

With UserList(UserIndex)
    'Remove Packet ID
    Call .incomingData.ReadByte
    
    'Convert version number to string
    CuentaEmail = .incomingData.ReadASCIIString
    CuentaPass = .incomingData.ReadASCIIString
    SerialHD = val(SDesencriptar(.incomingData.ReadASCIIString()))
    NamePJ = .incomingData.ReadASCIIString
    
    Call Mod_Cuentas.BorrarPersonaje(UserIndex, CuentaEmail, CuentaPass, SerialHD, NamePJ)
End With
End Sub
 
Public Sub EnviarCuenta(ByVal UserIndex As Integer, ByVal PJName As String, ByVal CantPjs As Byte, ByVal Total As Byte, ByVal estado As Byte)
Dim MapaUsuario As String
Dim UserFile As clsIniReader

With UserList(UserIndex).outgoingData
    .WriteByte (ServerPacketID.CuentaSV)
    .WriteByte (estado)
    .WriteByte (CantPjs)
    
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT clase,elv,body,head,casco,escudo,arma,raza,logged,position FROM personaje WHERE nombre = '" & PJName & "'")
    
    If CantPjs > 0 Then
        Dim Poss As String
        Poss = ReadField(1, RS!Position, 45)
        If Poss = 0 Then Poss = 1
        MapaUsuario = MapInfo(Poss).MapName
    Else
        MapaUsuario = ""
    End If
    
    If MapaUsuario = "" Then MapaUsuario = "Mapa Desconocido" 'Asi no tira error
    
    If Not CantPjs = 0 Then
        .WriteASCIIString (PJName)
        .WriteByte (RS!clase)
        .WriteByte (RS!ELV)
        .WriteInteger (RS!body)
        .WriteInteger (RS!Head)
        .WriteByte (RS!casco)
        .WriteByte (RS!Escudo)
        .WriteByte (RS!Arma)
        .WriteByte (RS!raza)
        .WriteByte (RS!logged)
        .WriteASCIIString (MapaUsuario)
    End If
    
    .WriteByte (Total)
End With
End Sub

Public Sub WriteCastleAttack(ByVal UserIndex As Integer, ByVal Castillo As Byte)
'***************************
'Autor: Lorwik ----------------> Te lo cambio xD <Edurne>
'1 para el castillo norte
'2 para el castillo este
'3 para el castillo sur
'4 para el castillo oeste
'5 para fortaleza
'***************************
On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CastleAttack)
    Call UserList(UserIndex).outgoingData.WriteByte(Castillo)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleGlobalMsg(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
     
    On Error GoTo errHandler
    With UserList(UserIndex)
     
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(.incomingData)
     
     
    Call buffer.ReadByte
     
    Dim Chat As String
     
    Chat = buffer.ReadASCIIString()
     
    If LenB(Chat) <> 0 Then
        'Analize chat...
        Call Statistics.ParseChat(Chat)
        If HayGlobal Then
            
            If Chat = "" Or Chat = " " Then
                Call WriteConsoleMsg(UserIndex, "Escribe un mensaje.", FONTTYPE_INFOBOLD)
                Call .incomingData.CopyBuffer(buffer)
                Exit Sub
            End If
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & "> " & Chat, FontTypeNames.FONTTYPE_global))
            Call AddtoRichTextBox(frmMain.RecTxt, vbNewLine & Hour(Now) & ":" & Minute(Now) & " - " & .Name & "> " & Chat, 139, 248, 244, False, True)
        Else
            Call WriteConsoleMsg(UserIndex, "¡El chat global esta desactivado!", FONTTYPE_INFOBOLD)
        End If
    End If
     
    Call .incomingData.CopyBuffer(buffer)
    End With
     
errHandler:
    Dim error As Long
    error = Err.Number
    On Error GoTo 0
     
     
    Set buffer = Nothing
     
    If error <> 0 Then _
    Err.Raise error
End Sub

Public Sub HandleGlobalOn(ByVal UserIndex As Integer)
 
    With UserList(UserIndex)
     
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User) Then Exit Sub
         
        If HayGlobal = True Then
        Call WriteConsoleMsg(UserIndex, "¡El modo de habla global ya se encontraba activado!", FONTTYPE_INFO)
        Exit Sub
        End If
         
        HayGlobal = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El chat Global ha sido activado. ¡Se ruega no abusar, podria llevar a sanción!", FontTypeNames.FONTTYPE_INFOBOLD))
    End With
End Sub
 
Public Sub HandleGlobalOff(ByVal UserIndex As Integer)
 
    With UserList(UserIndex)
     
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User) Then Exit Sub
         
        If HayGlobal = False Then
        Call WriteConsoleMsg(UserIndex, "¡El modo de habla global ya se encontraba desactivado!", FONTTYPE_INFO)
        Exit Sub
        End If
         
        HayGlobal = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El chat Global ha sido desactivado.", FontTypeNames.FONTTYPE_INFOBOLD))
    End With
End Sub

Public Sub HandleDueloSet(ByVal UserIndex As Integer)
Dim TipoDuelo As Byte
Dim MapaDuelo As Byte
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        TipoDuelo = .incomingData.ReadByte
        Select Case TipoDuelo
            Case 0
                Call EsperarOponenteDuelo(UserIndex)
            Case 1 'Duelo Clasico sin ELO
                Call EsperarOponenteDueloClasico(UserIndex, False)
            Case 2 'Duelo Clasico con ELO
                Call EsperarOponenteDueloClasico(UserIndex, True)
            Case 3 'Arena de Rinkel
                Call modArenaRinkel.EntrarArenaRinkel(UserIndex)
            Case 50 'Comenzar Arena de Rinkel
                Call modArenaRinkel.Preparar(UserIndex)
            Case Else
                Exit Sub
        End Select
    End With
End Sub

Private Sub HandleDropObj(ByVal UserIndex As Integer)
'***************************************************
'Author: maTih.-
'Last Modification: 10/11/2012 - ^[GS]^
'***************************************************

With UserList(UserIndex)

    Dim selInvSlot  As Byte         ' <<< Slot.
    Dim TargetX     As Byte         ' <<< Posición X.
    Dim TargetY     As Byte         ' <<< Posición Y.
    Dim Amount      As Integer      ' <<< Cantidad.
    Dim TargetNPC   As Integer      ' <<< Npc?.
    Dim targetUser  As Integer      ' <<< Usuario?.
    Dim targetObj   As Obj          ' <<< -

    'Remove packetID.
    Call .incomingData.ReadByte
    
    'Get the incoming Data.
    selInvSlot = .incomingData.ReadByte()
    TargetX = .incomingData.ReadByte()
    TargetY = .incomingData.ReadByte()
    Amount = .incomingData.ReadInteger()
    
    'If dead
    If .flags.Muerto = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.Muerto)
        Exit Sub
    End If
    
    '¿Está trabajando?
    If .flags.Makro <> 0 Then
        Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    'If user meditates
    If .flags.Meditando Then
        Call WriteConsoleMsg(UserIndex, "¡Estas meditando!", FontTypeNames.FONTTYPE_INFOBOLD)
        Exit Sub
    End If

    ' Distance
    If Distance(.Pos.X, .Pos.Y, TargetX, TargetY) > iniDragDrop Then
        Call WriteConsoleMsg(UserIndex, "¡¡No puedes tirar objetos los tan lejos!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Legal pos
    If LegalPos(.Pos.Map, TargetX, TargetY, .flags.Navegando, Not .flags.Navegando) = False Then
        Call WriteConsoleMsg(UserIndex, "¡¡Posición invalida!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If ObjData(UserList(UserIndex).Invent.Object(selInvSlot).ObjIndex).OBJType = otPiedraResu Then
        Call WriteConsoleMsg(UserIndex, "¡No puedes tirar ese objeto!", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
    
    If ObjData(UserList(UserIndex).Invent.Object(selInvSlot).ObjIndex).Nosetira = 1 Then
        Call WriteConsoleMsg(UserIndex, "¡No puedes tirar ese objeto!", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
    
    'Not valid slot or not valid objIndex?
    If Not selInvSlot <> 0 Then Exit Sub
    If Not .Invent.Object(selInvSlot).ObjIndex <> 0 Then Exit Sub
        
    'Analize the amount.
    If Not Amount <> 0 Then Exit Sub
    If Amount > .Invent.Object(selInvSlot).Amount Then Amount = .Invent.Object(selInvSlot).Amount
        
    'There is a user in that position?.
    If MapData(.Pos.Map, TargetX, TargetY).UserIndex <> 0 Then
        Call DropToUser(UserIndex, MapData(.Pos.Map, TargetX, TargetY).UserIndex, selInvSlot, Amount)
        Exit Sub
    End If
    
    If MapData(.Pos.Map, TargetX, TargetY).NpcIndex <> 0 Then
        Call DropToNPC(UserIndex, MapData(.Pos.Map, TargetX, TargetY).NpcIndex, selInvSlot, Amount)
        Exit Sub
    End If
        
    'Set the flags.
    targetObj.ObjIndex = .Invent.Object(selInvSlot).ObjIndex
    targetObj.Amount = Amount
    
    'Prevent the replace
    If MapData(.Pos.Map, TargetX, TargetY).ObjInfo.ObjIndex <> 0 Then
        If MapData(.Pos.Map, TargetX, TargetY).ObjInfo.ObjIndex <> targetObj.ObjIndex Then
            ' Other object exists
            Call WriteConsoleMsg(UserIndex, "¡¡¡Ya hay un objeto ahi!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    'There is no user or NPC, I throw the floor.
    'Is safe area?
    If Not MapInfo(.Pos.Map).Pk And iniTirarOBJZonaSegura = False Then
        Call WriteConsoleMsg(UserIndex, "¡¡No esta permitido arrojar objetos al suelo en zonas seguras!!", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
    End If
        
    'Create the object in the position indicated.
    MakeObj targetObj, .Pos.Map, TargetX, TargetY
        
    'Quit the obj to user.
    QuitarUserInvItem UserIndex, selInvSlot, Amount
    
    'Update the userinventory.
    UpdateUserInv False, UserIndex, selInvSlot
    
    'WriteConsoleMsg UserIndex, "Has arrojado los items correctamente", FontTypeNames.FONTTYPE_CITIZEN
End With

End Sub

Public Sub HandleMoveItem(ByVal UserIndex As Integer) ' 0.13.3
'***************************************************
'Author: Ignacio Mariano Tirabasso (Budi)
'Last Modification: 27/12/2013 - GoDKeR
'
'***************************************************

With UserList(UserIndex)

    Dim originalSlot As Byte
    Dim newSlot As Byte
    Dim tipo As Byte
    
    Call .incomingData.ReadByte
    
    originalSlot = .incomingData.ReadByte
    newSlot = .incomingData.ReadByte
    tipo = .incomingData.ReadByte
    
    Select Case tipo
        
            Case 1
                Call modUsuariosInv.MoveItem(UserIndex, originalSlot, newSlot)
            
            Case 3
                Call modHechizos.MoveSpell(UserIndex, originalSlot, newSlot) '#FABULOUS
    End Select
    
End With

End Sub

Public Sub WriteMontateToggle(ByVal UserIndex As Integer)
On Error Resume Next
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MontateToggle)
End Sub

Public Sub HandleSolInfMonturaClient(ByVal UserIndex As Integer)
Dim TipoPaquete As Byte
Dim SlotMontura As Byte
Dim Skill As Byte

    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        TipoPaquete = .incomingData.ReadByte
        SlotMontura = .incomingData.ReadByte
        Skill = .incomingData.ReadByte
        
        Select Case TipoPaquete
            Case 0
                Call WriteEnvioListaMonturas(UserIndex)
            Case 1
                Call WriteInfoMontura(UserIndex, SlotMontura)
            Case 2
                Call AsignacionSkillMontura(UserIndex, SlotMontura, Skill)
        End Select
    End With
End Sub

Public Sub WriteEnvioListaMonturas(ByVal UserIndex As Integer)
Dim i As Byte
    With UserList(UserIndex)
        If .Stats.NUMMONTURAS = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes ninguna montura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call .outgoingData.WriteByte(ServerPacketID.EnvioListaMonturas)
        
        'Envio el numero de monturas
        Call .outgoingData.WriteByte(.Stats.NUMMONTURAS)
        
        'Y ahora envio las monturas
        Call .outgoingData.WriteASCIIString(.flags.Montura(1).nombre)
        Call .outgoingData.WriteASCIIString(.flags.Montura(2).nombre)
        Call .outgoingData.WriteASCIIString(.flags.Montura(3).nombre)
        Call .outgoingData.WriteASCIIString(.flags.Montura(4).nombre)
        Call .outgoingData.WriteASCIIString(.flags.Montura(5).nombre)
    End With
End Sub

Public Sub WriteInfoMontura(ByVal UserIndex As Byte, ByVal SlotMontura As Byte)
    With UserList(UserIndex)
    
        'Para prevenir bugs
        If SlotMontura = 0 Then
            Call WriteConsoleMsg(UserIndex, "No tienes ninguna montura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        Call .outgoingData.WriteByte(ServerPacketID.EnvioInfoMontura)
        
        Call .outgoingData.WriteASCIIString(.flags.Montura(SlotMontura).nombre)
        Call .outgoingData.WriteByte(.flags.Montura(SlotMontura).tipo)
        Call .outgoingData.WriteByte(.flags.Montura(SlotMontura).MonturaLevel)
        Call .outgoingData.WriteLong(.flags.Montura(SlotMontura).Exp)
        
        Call .outgoingData.WriteByte(.flags.Montura(SlotMontura).Ataque)
        Call .outgoingData.WriteByte(.flags.Montura(SlotMontura).Defensa)
        Call .outgoingData.WriteByte(.flags.Montura(SlotMontura).AtMagia)
        Call .outgoingData.WriteByte(.flags.Montura(SlotMontura).DefMagia)
        Call .outgoingData.WriteByte(.flags.Montura(SlotMontura).Evasion)
        Call .outgoingData.WriteByte(.flags.Montura(SlotMontura).Skills)
        Call .outgoingData.WriteByte(.flags.Montura(SlotMontura).Speed)
    End With
End Sub

''
' Handles the "SolParty" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSolParty(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Recibimos las intenciones de un usuario de crear party y le preguntamos al otro si se quiere unir
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        Dim tU As Integer   'Target user

        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        Call ClickIzquierdo(UserIndex, .Pos.Map, X, Y)
        
        tU = .flags.targetUser
                
        'Validate target
        If tU > 0 Then
            '¿Esta cerca?
            If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                Exit Sub
            End If
                    
            'Prevent from hitting self
            If tU = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes enviarte una solicitud a ti mismo!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(tU).PartyId > 0 Then
                Call WriteConsoleMsg(UserIndex, "¡El usuario ya parte de una Party!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
                
            If UserList(tU).PartyId = 0 Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Estás muerto!", FontTypeNames.FONTTYPE_PARTY)
                End If
                    
                Call WritePeticionInvitacionAParty(UserIndex, tU)
            End If
                
        End If
    End With
End Sub




Private Sub HandleEntregarDC(ByVal UserIndex As Integer)
 
If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Dim EntregaoQuita As Boolean
        Call buffer.CopyBuffer(.incomingData)
       
        'Remove packet ID
        Call buffer.ReadByte
       
        Dim UserName As String
        Dim DragCreditos As Long
        Dim tUser As Integer
       
        UserName = buffer.ReadASCIIString()
        DragCreditos = buffer.ReadLong()
        EntregaoQuita = buffer.ReadBoolean
        
        If UserIndex > 0 Then
             If InStr(1, UserName, " ") Then
                 UserName = Replace(UserName, " ", " ")
             End If
            
             If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
            
             If EntregaoQuita = True Then
            
                  If LenB(UserName) = 0 Then
                      Call WriteConsoleMsg(UserIndex, "Comando Incorrecto. Utilice /ENTREGARDC NICKNAME@PUNTOS.", FontTypeNames.FONTTYPE_INFO)
                  Else
                      tUser = NameIndex(UserName)
                  End If
                 
                  If tUser <= 0 Then
                      Call WriteConsoleMsg(UserIndex, "El Usuario no esta Online", FontTypeNames.FONTTYPE_INFO)
                  End If

                    Call LogGM("Historial DragCreditos Entregados", UserList(UserIndex).Name & " entrego " & DragCreditos & " a " & UserList(tUser).Name & " ahora tiene " & UserList(tUser).Stats.DragCredits + DragCreditos)
                    UserList(tUser).Stats.DragCredits = UserList(tUser).Stats.DragCredits + DragCreditos
                    Call WriteConsoleMsg(tUser, "¡Has recibido " & DragCreditos & " DragCreditos!", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(UserIndex, "Le has dado " & DragCreditos & " DragCreditos a " & UserList(tUser).Name, FontTypeNames.FONTTYPE_INFO)
                    Call WriteUpdateUserStats(UserIndex)

             Else
                 If LenB(UserName) = 0 Then
                      Call WriteConsoleMsg(UserIndex, "Comando Incorrecto. Utilice /QUITARDC NICKNAME@PUNTOS.", FontTypeNames.FONTTYPE_INFO)
                  Else
                      tUser = NameIndex(UserName)
                  End If
                 
                  If tUser <= 0 Then
                      Call WriteConsoleMsg(UserIndex, "El Usuario no esta Online", FontTypeNames.FONTTYPE_INFO)
                  End If
                 
                  If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                      UserList(tUser).Stats.DragCredits = UserList(tUser).Stats.DragCredits - DragCreditos
                      Call WriteConsoleMsg(tUser, "Se te han restado " & DragCreditos & " DragCreditos", FontTypeNames.FONTTYPE_INFO)
                      Call WriteConsoleMsg(UserIndex, "Le has quitado " & DragCreditos & " DragCreditos a " & UserList(tUser).Name, FontTypeNames.FONTTYPE_INFO)
                      Call WriteUpdateUserStats(UserIndex)
                  Else
                      Call WriteConsoleMsg(UserIndex, "El Personaje no existe", FontTypeNames.FONTTYPE_INFO)
                  End If
             End If
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
   
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set buffer = Nothing
   
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandlePremios(ByVal UserIndex As Integer)
'Segun el tipo, le mostramos los premios o le entregamos uno
Dim tipo As Byte
Dim ObjPremio As Obj
Dim PremioIndex As Integer
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        tipo = .incomingData.ReadByte
        PremioIndex = .incomingData.ReadInteger
        
        Select Case tipo
            Case 0
                Call WriteMuestroPremios(UserIndex)
            Case 1
                If PremioIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No seleccionaste ningún premio.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Set the object
                ObjPremio.ObjIndex = PremiosInfo(PremioIndex).ObjIndex
                ObjPremio.Amount = PremiosInfo(PremioIndex).cantidad
                 
                If ObjPremio.ObjIndex <= 0 Then Exit Sub
                 
                If PremiosInfo(PremioIndex).Puntos <= UserList(UserIndex).Stats.DragCredits Then
                    If Not MeterItemEnInventario(UserIndex, ObjPremio) Then
                        Call WriteConsoleMsg(UserIndex, "No tienes espacio en el inventario.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call LogGM("Historial de Canjes", UserList(UserIndex).Name & " tenia: " & UserList(UserIndex).Stats.DragCredits & " Dragcredotidos, y ahora tiene: " & UserList(UserIndex).Stats.DragCredits - PremiosInfo(PremioIndex).Puntos & " al adquirir: " & ObjData(ObjPremio.ObjIndex).Name)
                        UserList(UserIndex).Stats.DragCredits = UserList(UserIndex).Stats.DragCredits - PremiosInfo(PremioIndex).Puntos
                        Call UpdateUserInv(True, UserIndex, 0)
                        Call WriteUpdateUserStats(UserIndex)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes suficientes puntos.", FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

Private Sub WriteMuestroPremios(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If CantPremios = 0 Then
            Call WriteConsoleMsg(UserIndex, "En estos momentos no hay ningun premio para mostrar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
   
        Dim i As Integer
        Call .outgoingData.WriteByte(ServerPacketID.muestropremios)
        Call .outgoingData.WriteInteger(CantPremios)
   
        For i = 1 To CantPremios
            Call .outgoingData.WriteASCIIString(ObjData(PremiosInfo(i).ObjIndex).Name)
            Call .outgoingData.WriteInteger(PremiosInfo(i).Puntos)
            Call .outgoingData.WriteInteger(PremiosInfo(i).cantidad)
        Next i
    End With
End Sub

Public Sub WriteAtaqueNPC(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.AtaqueNPC)
    Call UserList(UserIndex).outgoingData.WriteInteger(NpcIndex)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteUpdateGLDBank(ByVal UserIndex As Integer, ByVal cantidad As Long)
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UpdateGLDBank)
    Call UserList(UserIndex).outgoingData.WriteLong(cantidad)
End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
        Call .WriteByte(FuerzaMaxima(UserIndex))
        Call .WriteByte(AgilidadMaxima(UserIndex))
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(AgilidadMaxima(UserIndex))
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(FuerzaMaxima(UserIndex))
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleRequieredCaptions(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim miBuffer As New clsByteQueue
         
        Call miBuffer.CopyBuffer(.incomingData)
         
        Call miBuffer.ReadByte
         
        Dim tU As Integer
        tU = NameIndex(miBuffer.ReadASCIIString)
         
        If .flags.Privilegios > PlayerType.User Then
            If tU > 0 Then
                WriteRequieredCAPTIONS tU
                UserList(tU).elpedidor = UserIndex
            Else
                WriteConsoleMsg UserIndex, "User Offline", FontTypeNames.FONTTYPE_GUILD
            End If
        End If
         
        Call .incomingData.CopyBuffer(miBuffer)
     
    End With
End Sub
 
Private Sub HandleSendCaptions(ByVal UserInde As Integer)
    With UserList(UserInde)
        Dim miBuffer As New clsByteQueue
         
        Call miBuffer.CopyBuffer(.incomingData)
         
        Call miBuffer.ReadByte
         
        Dim Captions As String
        Dim cCaptions As Byte
        Captions = miBuffer.ReadASCIIString()
        cCaptions = miBuffer.ReadByte()
         
        If .elpedidor > 0 Then
        WriteShowCaptions .elpedidor, Captions, cCaptions, .Name
        End If
         
        Call .incomingData.CopyBuffer(miBuffer)
     
    End With
End Sub
 
Public Sub WriteShowCaptions(ByVal UserIndex As Integer, ByVal Caps As String, ByVal cCAPS As Byte, ByVal sendIndex As String)
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowCaptions)
        Call .WriteASCIIString(sendIndex)
        Call .WriteASCIIString(Caps)
        Call .WriteByte(cCAPS)
    End With
End Sub
 
Public Sub WriteRequieredCAPTIONS(ByVal UserIndex As Integer)
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.rCaptions)
    End With
End Sub

Public Sub HandleAdvertirPJ(ByVal UserIndex As Integer)
 If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo errHandler
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
       
        Call buffer.ReadByte
       
        Dim UserAdvertido As String
        Dim UserAdvName As Integer
       
        UserAdvertido = buffer.ReadASCIIString
        UserAdvName = NameIndex(UserAdvertido)
 
        Call General.AdvertirUsuario(UserIndex, UserAdvName)
 
        Call .incomingData.CopyBuffer(buffer)
    End With
 
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
   
    Set buffer = Nothing
   
    If error <> 0 Then _
        Err.Raise error
 
End Sub

Public Sub HandleSearchObj(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
       
        'Remove packet ID
        Call buffer.ReadByte
       
        Dim UserObj As String
        Dim tUser As Integer
        Dim Rank As Integer
        Dim n As Integer
        Dim i As Integer
       
        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
       
        UserObj = buffer.ReadASCIIString()
       
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
           
            For i = 1 To UBound(ObjData)
                If InStr(1, Tilde(ObjData(i).Name), Tilde(UserObj)) Then
                    Call WriteConsoleMsg(UserIndex, i & " " & ObjData(i).Name & ".", FontTypeNames.FONTTYPE_CENTINELA)
                    n = n + 1
                End If
            Next
            If n = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hubo resultados de la busqueda: " & UserObj, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Hubo " & n & " resultados de la busqueda: " & UserObj, FontTypeNames.FONTTYPE_INFO)
            End If
           
        End If
       
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
 
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set buffer = Nothing
   
    If error <> 0 Then _
        Err.Raise error
End Sub

Public Sub HandleSearchNPC(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
       
        'Remove packet ID
        Call buffer.ReadByte
       
Dim p As Integer
Dim npc As String
Dim UserNPC As String
Dim Rank As Integer
Dim i As Integer
Dim n As Integer

    Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
       
    UserNPC = buffer.ReadASCIIString()

    If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
        p = val(GetVar(App.Path & "\Dat\Npcs.dat", "INIT", "NumNPCs"))
        For i = 1 To p
        npc = GetVar(App.Path & "\Dat\Npcs.dat", "Npc" & i, "Name")
            If InStr(1, Tilde(npc), Tilde(UserNPC)) Then
                Call WriteConsoleMsg(UserIndex, i & " " & npc & ".", FontTypeNames.FONTTYPE_CENTINELA)
                n = n + 1
            End If
        Next
            If n = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hubo resultados de la busqueda: " & UserNPC, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Hubo " & n & " resultados de la busqueda: " & UserNPC, FontTypeNames.FONTTYPE_INFO)
            End If
    End If
       
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
 
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set buffer = Nothing
   
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "VerHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleVerHD(ByVal UserIndex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification:  08/06/2012 - ^[GS]^
'Verifica el HD del usuario.
'***************************************************
 
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo errHandler
    With UserList(UserIndex)
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
               
        Dim iUsuario As Integer

        iUsuario = NameIndex(buffer.ReadASCIIString())
       
        If iUsuario = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "El usuario " & UserList(iUsuario).Name & " tiene un disco con el Serial " & UserList(iUsuario).flags.SerialHD, FONTTYPE_INFOBOLD)
        End If
       
        Call .incomingData.CopyBuffer(buffer)
       
    End With
   
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then Err.Raise error
    
End Sub
 
''
' Handles the "UnBanHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUnbanHD(ByVal UserIndex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification:  08/06/2012 - ^[GS]^
'Maneja el unbaneo del serial del HD de un usuario.
'***************************************************
 
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
 
On Error GoTo errHandler
    With UserList(UserIndex)
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
               
        Dim SerialHD As String
        SerialHD = buffer.ReadASCIIString()
       
        If (BanHD_rem(SerialHD)) Then
            Call WriteConsoleMsg(UserIndex, "El disco con el Serial " & SerialHD & " se ha quitado de la lista de baneados.", FONTTYPE_INFOBOLD)
        Else
            Call WriteConsoleMsg(UserIndex, "El disco con el Serial " & SerialHD & " no se encuentra en la lista de baneados.", FONTTYPE_INFO)
        End If
       
        Call .incomingData.CopyBuffer(buffer)
    End With
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then Err.Raise error
 
End Sub
 
''
' Handles the "BanHD" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleBanHD(ByVal UserIndex As Integer)
'***************************************************
'Author: ArzenaTh
'Last Modification: 08/06/2012 - ^[GS]^
'Maneja el baneo del serial del HD de un usuario.
'***************************************************
 
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
 
On Error GoTo errHandler
    With UserList(UserIndex)
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim i As Long
        Dim iUsuario As Integer
        Dim BannedHD As String
        
        iUsuario = NameIndex(buffer.ReadASCIIString())
        If iUsuario > 0 Then
            BannedHD = UserList(iUsuario).flags.SerialHD
        End If
        
        If .flags.Privilegios And (PlayerType.Admin And PlayerType.Dios) Then
            If LenB(BannedHD) > 0 Then
                If BanHD_find(BannedHD) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanHD_add(BannedHD)
                    Call WriteConsoleMsg(UserIndex, "Has baneado el disco duro " & BannedHD & " del usuario " & UserList(iUsuario).Name, FontTypeNames.FONTTYPE_INFO)
                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).flags.SerialHD = BannedHD Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "Ban de serial de disco duro.")
                            End If
                        End If
                    Next i
                End If
            ElseIf iUsuario <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
       
        Call .incomingData.CopyBuffer(buffer)
    End With
   
errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then Err.Raise error
End Sub

''
' Handles the "BanAcc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanAcc(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call WriteVar(App.Path & "\Cuentas\" & UserList(UserIndex).Accounted & ".chr", "CUENTA", "BAN", 1)
            Call WriteVar(App.Path & "\Cuentas\" & UserList(UserIndex).Accounted & ".chr", "CUENTA", "RAZON", reason)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UnbanAcc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanAcc(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim cantPenas As Byte
        Dim NameAccount As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
                Call WriteConsoleMsg(UserIndex, "Personaje inexistente (no use +)", FontTypeNames.FONTTYPE_INFO)
            Else
            
                NameAccount = GetVar(CharPath & UserName & ".chr", "INIT", "Accounted")
                
                If GetVar(App.Path & "\Cuentas\" & NameAccount & ".chr", "Cuenta", "BAN") = 1 Then
                    Call WriteVar(App.Path & "\Cuentas\" & NameAccount & ".chr", "Cuenta", "BAN", 0)
                    Call WriteVar(App.Path & "\Cuentas\" & NameAccount & ".chr", "Cuenta", "RAZON", "")
                
                    Call LogGM(.Name, "/UNBANACC a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "La cuenta de " & UserName & " no esta baneada. Es imposible unbanear", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

errHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleTPCastle(ByVal UserIndex As Integer)
Dim Mapa As Byte, X As Byte, Y As Byte
Dim Castillo As Byte

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        Castillo = .incomingData.ReadByte
        
        'COMPROBACIONES:
        If .Stats.MinHP < .Stats.MaxHP Then
            Call WriteConsoleMsg(UserIndex, "Tu salud debe estar completa.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .Counters.Pena > 0 Or UserList(UserIndex).Pos.Map = Prision.Map Then
            Call WriteMultiMessage(UserIndex, eMessages.Carcel)
            Exit Sub
        End If

        If .flags.Paralizado = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Paralizado)
            Exit Sub
        End If
        
        If .flags.Ceguera = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.Cegado)
            Exit Sub
        End If
        
        If .GuildIndex <> Castillos(Castillo).Dueño Then
            Call WriteConsoleMsg(UserIndex, "El castillo no te pertenece.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Pos.Map = Mapa Then
            Call WriteConsoleMsg(UserIndex, "¡Ya estas dentro!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Castillo
            Case 1 'Castillo Norte
                Mapa = 25
                X = RandomNumber(48, 55)
                Y = RandomNumber(50, 60)
            Case 2 'Castillo Este
                Mapa = 26
                X = RandomNumber(48, 55)
                Y = RandomNumber(50, 60)
            Case 3 'Castillo Sur
                Mapa = 15
                X = RandomNumber(48, 55)
                Y = RandomNumber(50, 60)
            Case 4 'Castillo Oeste
                Mapa = 27
                X = RandomNumber(48, 55)
                Y = RandomNumber(50, 60)
            Case 5 'Fortaleza
                Mapa = 28
                X = RandomNumber(40, 61)
                Y = RandomNumber(46, 47)
        End Select

        
        Call WarpUserChar(UserIndex, Mapa, X, Y, True)
    End With
End Sub

Public Sub WriteSegNoti(ByVal UserIndex As Integer, ByVal tipo As Byte, ByVal Activado As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.SegNoti)
        Call .outgoingData.WriteByte(tipo)
        Call .outgoingData.WriteBoolean(Activado)
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteEnviaManual(ByVal UserIndex As Integer)
On Error GoTo errHandler
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.EnviaManual)
    End With
    Exit Sub
errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleSolicitarRank(ByVal UserIndex As Integer)
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteEnviaRank(UserIndex)
End Sub

Private Sub WriteEnviaRank(ByVal UserIndex As Integer)
Dim i As Byte
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.EnviarRanking)
        
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT nombre,elo FROM personaje ORDER BY elo DESC LIMIT 5")
        
        i = 1
        While Not RS.EOF
            Call .outgoingData.WriteASCIIString(RS!nombre)
            Call .outgoingData.WriteDouble(RS!ELO)
            i = i + 1
            RS.MoveNext
        Wend
        
        Call .outgoingData.WriteDouble(.Stats.ELO)
    
    End With
End Sub

Private Sub HandleHacerTorneo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        MinLevel = .incomingData.ReadByte
        MaxLevel = .incomingData.ReadByte
        Cupos = .incomingData.ReadByte
        AutoSum = .incomingData.ReadByte
        If AutoSum = 1 Then
        Mapa = .incomingData.ReadInteger
        X = .incomingData.ReadByte
        Y = .incomingData.ReadByte
        End If
         
        If .flags.Privilegios And (PlayerType.User) Then Exit Sub
         
        If Hay_Torneo = False Then
            Hay_Torneo = True
           
            Dim data As String
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Torneo organizado por " & UserList(UserIndex).Name & "]", FontTypeNames.FONTTYPE_WARNING))
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Nivel máximo: " & MaxLevel & "", FontTypeNames.FONTTYPE_GM))
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Nivel Mínimo: " & MinLevel & "", FontTypeNames.FONTTYPE_GM))
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Cupo máximo: " & Cupos & "", FontTypeNames.FONTTYPE_GM))
           ' Dim i As Integer
            'For i = 1 To 8
            '    If Torneo_Clases_Validas2(i) = 1 Then
             '       Data = Data & Torneo_Clases_Validas(i) & ","
             '   End If
           ' Next
           ' Data = Left$(Data, Len(Data) - 1)
           ' Data = Data & "."
           ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Clases válidas: " & Data & "", FontTypeNames.FONTTYPE_INFO))
           ' Data = ""
           ' For i = 1 To 4
              '  If Torneo_Alineacion_Validas2(i) = 1 Then
                  '  Data = Data & Torneo_Alineacion_Validas(i) & ","
              '  End If
            'Next
            'Data = Left$(Data, Len(Data) - 1)
            'Data = Data & "."
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Alineación válidas: Todas ", FontTypeNames.FONTTYPE_GM))
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Escribe ""/TORNEO"" para participar.", FontTypeNames.FONTTYPE_GM))
           
            Else
            Call WriteConsoleMsg(UserIndex, "Ya hay un torneo", FontTypeNames.FONTTYPE_INFO)
        End If
     
    End With
End Sub
Public Sub HandleTorneo(ByVal UserIndex As Integer)
With UserList(UserIndex)
Call .incomingData.ReadByte
 
 
            If Hay_Torneo = False Then
            Call WriteConsoleMsg(UserIndex, "No hay ningún torneo disponible.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            End If
           
            If UserList(UserIndex).Stats.ELV < MinLevel Then
                Call WriteConsoleMsg(UserIndex, "Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El requerido es: " & MinLevel, FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(UserIndex).Stats.ELV > MaxLevel Then
                Call WriteConsoleMsg(UserIndex, "Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El máximo es: " & MaxLevel, FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If Torneo_Inscriptos >= Cupos Then
                Call WriteConsoleMsg(UserIndex, "El cupo ya ha sido alcanzado.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
           
'            Dim i As Integer
'            For i = 1 To 8
'                If UCase$(UserList(UserIndex).clase) = UCase$(Torneo_Clases_Validas(i)) And Torneo_Clases_Validas2(i) = 0 Then
'                Call WriteConsoleMsg(UserIndex, "Tu clase no es válida en este torneo.", FontTypeNames.FONTTYPE_INFO)
 '               Exit Sub
 '               End If
     '       Next
           
            Dim NuevaPos As worldPos
           
           
            'El que entra no sale.
            If Not UserList(UserIndex).flags.EnTorneo = True Then
                Call WriteConsoleMsg(UserIndex, "Estás en la lista de espera del torneo. Estás en el puesto nº " & Torneo_Inscriptos + 1, FontTypeNames.FONTTYPE_INFO)
                .flags.EnTorneo = True
               
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("/TORNEO [" & UserList(UserIndex).Name & "]", FontTypeNames.FONTTYPE_INFOBOLD))
                Torneo_Inscriptos = Torneo_Inscriptos + 1
                If Torneo_Inscriptos = Cupos Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Cupo alcanzado.", FontTypeNames.FONTTYPE_exp))
                End If
                If AutoSum = 1 Then
                    Call WarpUserChar(UserIndex, Mapa, X, Y, True)
                End If
            Else
'                Call Torneo.Quitar(UserList(Userindex).Name)
                Call WriteConsoleMsg(UserIndex, "Ya estás en la lista de espera del torneo.", FontTypeNames.FONTTYPE_INFO)
'                Torneo_Inscriptos = Torneo_Inscriptos - 1
'                If Torneo_SumAuto = 1 Then
'                    Call WarpUserChar(Userindex, 1, 50, 50, True)
'                End If
            End If
 
End With
End Sub
Public Sub HandleCerrarTorneo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        If Hay_Torneo = True Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Las incripciones al torneo fueron cerradas", FontTypeNames.FONTTYPE_exp))
            Hay_Torneo = False
            Dim i As Long
            For i = 1 To LastUser
                UserList(i).flags.EnTorneo = True
            Next i
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("No hay ningun torneo disponible", FontTypeNames.FONTTYPE_INFO))
        End If
    End With
End Sub

Public Sub WritePartyClanIndex(ByVal UserIndex As Integer, ByVal tipo As Byte)
On Error GoTo errHandler
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.PartyClanIndex)
        Call .outgoingData.WriteByte(tipo)
        
        'Party
        If tipo = 0 Then
            Call .outgoingData.WriteInteger(.PartyId)
            Debug.Print .PartyId
        ElseIf tipo = 1 Then 'Clanes
            Call .outgoingData.WriteASCIIString(" <" & modGuilds.GuildName(.GuildIndex) & ">")
        End If
    End With
    Exit Sub
errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Handles the "ReleasePet" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReleasePet(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(NPCList(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteMultiMessage(UserIndex, eMessages.Lejos)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If NPCList(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call QuitarPet(UserIndex, .flags.TargetNPC)
            
    End With
End Sub

Private Sub HandleMotdCastillos(ByVal UserIndex As Integer)
    Call UserList(UserIndex).incomingData.ReadByte
    
    '15/03/2016 Lorwik
    Call DueñosCastillos(UserIndex)
End Sub

Public Sub WritePeticionInvitacionAParty(ByVal LiderIndex As Integer, ByVal InvitadoIndex As Integer)
'Le preguntamos si se quiere unir a nuestra party
    With UserList(InvitadoIndex)
    
        'Para prevenir bugs
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(LiderIndex, eMessages.Muerto)
            Exit Sub
        End If
        
        Call .outgoingData.WriteByte(ServerPacketID.PeticionInvitacionAParty)
        Call .outgoingData.WriteASCIIString(UserList(LiderIndex).Name)
            
    End With
End Sub



Private Sub HandleAbandonarParty(ByVal AbandonaIndex As Integer)

    Dim PartyId As Integer
    Dim PersonajeId As Integer
    Dim NuevoLiderId As Integer
    Dim NuevoLiderIndex As Integer
    Dim PartyUserIndex As Integer
    Dim Index As Integer
    
    
    Set RS = New ADODB.Recordset
    
    With UserList(AbandonaIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '20/11/2015 Irongete: Comprobar si tiene party y obtener la id de la misma
        PartyId = GetPersonajePartyId(AbandonaIndex)
        PersonajeId = GetPersonajeId(AbandonaIndex)
        
        If PartyId > 0 Then
        
            Set RS = SQL.Execute("DELETE FROM rel_party_personaje WHERE id_personaje = '" & PersonajeId & "' AND id_party = '" & PartyId & "'")
            Call Drag_Party.BroadcastParty(PartyId, UserList(AbandonaIndex).Name & " ha abandonado la party.")
            Call WriteConsoleMsg(AbandonaIndex, "Has abandonado la party", FONTTYPE_PARTY)
            .PartyId = 0
            
            '18/02/2016 Irongete: Le pongo la PartyId a 0 a la persona que sale
            Call WriteSetPartyId(AbandonaIndex, 0, 0)
            
            '18/02/2016 Irongete: Si queda mas gente en la party, le envio el CharIndex de la persona que se ha ido a cada uno de los miembros para que pongan el PartyId del que se acaba de ir a 0
            Set RS = SQL.Execute("SELECT j1.userindex AS `Index` FROM rel_party_personaje JOIN personaje j1 ON j1.id = rel_party_personaje.id_personaje WHERE rel_party_personaje.id_party = '" & PartyId & "' AND rel_party_personaje.id_personaje <> '" & UserList(AbandonaIndex).id & "'")
            If RS.RecordCount > 0 Then
                While Not RS.EOF
                    Index = RS!Index
                    Call WriteSetPartyId(Index, AbandonaIndex, 0)
                    RS.MoveNext
                Wend
            End If

            '16/02/2016 Irongete: Si solo queda una persona en party no le damos el lider, la disolvemos.
            Set RS = SQL.Execute("SELECT j1.UserIndex As `Index` FROM rel_party_personaje JOIN personaje j1 ON j1.id = rel_party_personaje.id_personaje WHERE id_party = '" & PartyId & "'")
            If RS.RecordCount = 1 Then
            
                While Not RS.EOF
                    Index = RS!Index
                    UserList(Index).PartyId = 0
                    Call WriteSetPartyId(Index, 0, 0)
                    RS.MoveNext
                Wend
                
                '16/02/2016 Irongete: Borro la party y la persona que quedaba en ella
                Call Drag_Party.BroadcastParty(PartyId, "Se ha disuelto la party")
                Set RS = SQL.Execute("DELETE FROM party WHERE id = '" & PartyId & "'")
                Set RS = SQL.Execute("DELETE FROM rel_party_personaje WHERE id_party = '" & PartyId & "'")
        
            Else
                '20/11/2015 Irongete: Si quien abandona la party es el lider, buscamos uno nuevo
                If EsLiderDeParty(AbandonaIndex) Then
                
                    '20/11/2015 Irongete: Busco un miembro en la party al que hacer lider
                    Set RS = SQL.Execute("SELECT id_personaje FROM rel_party_personaje WHERE id_party = '" & PartyId & "' LIMIT 1")
                    
                    If RS.RecordCount = 1 Then
                        NuevoLiderId = RS!id_personaje
                        NuevoLiderIndex = GetPersonajeIndex(NuevoLiderId)
                        Call NuevoLiderParty(NuevoLiderIndex)
                    Else
                        Call BorrarParty(PartyId)
                    End If
                    
                End If
            
            
            End If
            
        
            
            
            
            
            
        
        End If

    End With
    
End Sub

Private Sub HandleRespuestaPeticionAParty(ByVal InvitadoIndex As Integer)
'El usuario al que hemos invitado acepta o rechaza la invitacion de Party
    
    Dim LiderId As Integer
    Dim LiderIndex As Integer
    Dim PartyId As Integer
    Dim Respuesta As Boolean
    Set RS = New ADODB.Recordset
    
    With UserList(InvitadoIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        Respuesta = .incomingData.ReadBoolean
    
        '18/11/2015 Irongete: Compruebo que este jugador tiene una invitación pendiente
        Set RS = SQL.Execute("SELECT id_invita FROM rel_party_invitacion WHERE id_invitado = '" & UserList(InvitadoIndex).id & "'")
        If RS.RecordCount = 1 Then
        
        
        LiderId = RS!id_invita
        LiderIndex = GetPersonajeIndex(LiderId)
        
        
        
            
            'Acepto o rechazo la oferta de party?
            If Respuesta = True Then
            
                '18/11/2015 Irongete: Obtengo el userindex de la persona que manda la invitación
                
                
                '18/11/2015 Irongete: La persona que invita está ya en party o quiere crear una nueva
                PartyId = GetPersonajePartyId(LiderIndex)
                If PartyId > 0 Then
                    Call Drag_Party.AñadirAParty(PartyId, InvitadoIndex)
                Else
                    Call Drag_Party.CrearParty(LiderIndex, InvitadoIndex)
                    Exit Sub
                End If
                
            Else
                Call WriteConsoleMsg(LiderIndex, UserList(InvitadoIndex).Name & " ha rechazado tu invitacion.", FontTypeNames.FONTTYPE_GUILD)
                Set RS = SQL.Execute("DELETE FROM rel_party_invitacion WHERE id_invitado = '" & UserList(InvitadoIndex).id & "'")
                Exit Sub
            End If
        
        Else
            Call WriteConsoleMsg(InvitadoIndex, "No tienes ninguna invitación a party.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
    
    End With
    
End Sub

Public Sub HandleNuevoLiderDeParty(ByVal UserIndex As Integer)

    Dim NuevoLiderNombre As String
    Dim NuevoLiderIndex As Integer

    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        NuevoLiderNombre = .incomingData.ReadASCIIString
        
        '20/11/2015 Irongete: Id del personaje expulsado
        Set RS = SQL.Execute("SELECT userindex FROM personaje WHERE nombre = '" & NuevoLiderNombre & "' AND logged = '1'")
        NuevoLiderIndex = RS!UserIndex
        
        Call NuevoLiderParty(NuevoLiderIndex)
    
    End With
        
        
        
        
End Sub

'17/11/2015 Irongete: Expulsa a un miembro de la party
' @param LiderIndex userindex del lider de la party
' @param ExpulsadoIndex userindex de la persona que va a ser expulsada
Public Sub HandleExpulsarDeParty(ByVal LiderIndex As Integer)
    
    Dim PartyId As Integer
    Dim ExpulsadoNombre As String
    Dim ExpulsadoIndex As Integer
    Dim PartyUserIndex As Integer
    Dim Index As Integer
    
    Set RS = New ADODB.Recordset
    
    With UserList(LiderIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ExpulsadoNombre = .incomingData.ReadASCIIString
        
        PartyId = GetPartyId(LiderIndex)
        
        '20/11/2015 Irongete: Id del personaje expulsado
        Set RS = SQL.Execute("SELECT userindex FROM personaje WHERE nombre = '" & ExpulsadoNombre & "' AND logged = '1'")
        ExpulsadoIndex = RS!UserIndex
        
    
        '17/11/2015 Irongete: Comprobar que ambos jugadores estan en la misma party
         If PartyId = GetPartyId(ExpulsadoIndex) Then
         
            '17/11/2015 Irongete: Para expulsar si hace falta ser lider de la party
            If EsLiderDeParty(LiderIndex) Then
            
                '17/11/2015 Irongete: Expulsar al miembro de la party
                Set RS = SQL.Execute("DELETE FROM rel_party_personaje WHERE id_personaje = '" & UserList(ExpulsadoIndex).id & "'")
                UserList(ExpulsadoIndex).PartyId = 0
                
                '20/11/2015 Irongete: Broadcast de la expulsion a los miembros de la party
                Call Drag_Party.BroadcastParty(PartyId, UserList(LiderIndex).Name & " ha expulsado a " & UserList(ExpulsadoIndex).Name)
                Call WriteConsoleMsg(ExpulsadoIndex, UserList(LiderIndex).Name & " te ha expulsado de la party", FontTypeNames.FONTTYPE_PARTY)
                
                '20/11/2015 Irongete: Envio al jugador cual es su PartyId
                Call WriteSetPartyId(ExpulsadoIndex, 0, 0)
                
                '18/02/2016 Irongete: Si queda mas gente en la party, le envio el CharIndex de la persona que ha sido expulsada a cada uno de los miembros para que pongan el PartyId a 0
                Set RS = SQL.Execute("SELECT j1.userindex AS `Index` FROM rel_party_personaje JOIN personaje j1 ON j1.id = rel_party_personaje.id_personaje WHERE rel_party_personaje.id_party = '" & PartyId & "'")
                If RS.RecordCount > 0 Then
                    While Not RS.EOF
                        Index = RS!Index
                        Call WriteSetPartyId(Index, ExpulsadoIndex, 0)
                        RS.MoveNext
                    Wend
                End If
            
            Else
                Call WriteConsoleMsg(LiderIndex, "No eres el lider de la party", FontTypeNames.FONTTYPE_PARTY)
            End If
        End If
        
        '16/02/2016 Irongete: Si solo queda una persona en party no le damos el lider, la disolvemos.
        Set RS = SQL.Execute("SELECT j1.UserIndex As UserIndex FROM rel_party_personaje JOIN personaje j1 ON j1.id = rel_party_personaje.id_personaje WHERE id_party = '" & PartyId & "'")
        If RS.RecordCount = 1 Then
        
            While Not RS.EOF
                PartyUserIndex = RS!UserIndex
                UserList(PartyUserIndex).PartyId = 0
                Call WriteSetPartyId(PartyUserIndex, 0, 0)
                RS.MoveNext
            Wend
            
            '16/02/2016 Irongete: Borro la party y la persona que quedaba en ella
            Call Drag_Party.BroadcastParty(PartyId, "Se ha disuelto la party")
            Set RS = SQL.Execute("DELETE FROM party WHERE id = '" & PartyId & "'")
            Set RS = SQL.Execute("DELETE FROM rel_party_personaje WHERE id_party = '" & PartyId & "'")
    
        End If

    
    End With


End Sub

Public Sub HandleActualizarParty(ByVal UserIndex As Integer)
    With UserList(UserIndex)
    'Remove packet ID
        Call .incomingData.ReadByte
            
        Call WriteActualizacionParty(UserIndex)
    End With
End Sub

Public Sub WriteActualizacionParty(ByVal UserIndex As Integer)
    Dim Tmp As String
    Dim PartyId As Integer
    Dim LiderId As Integer
    Dim JugadorNombre As String
    Dim JugadorId As Integer
    
    
    With UserList(UserIndex).outgoingData
    
        PartyId = GetPartyId(UserIndex)
        
        Call .WriteByte(ServerPacketID.ActualizacionParty)
        
        '20/11/2015 Irongete: Miembros que forman la party
        Set RS = New ADODB.Recordset
        Set RS = SQL.Execute("SELECT party.lider, personaje.nombre, personaje.id FROM party INNER JOIN rel_party_personaje ON party.id = rel_party_personaje.id_party INNER JOIN personaje ON rel_party_personaje.id_personaje = personaje.id WHERE party.id = '" & PartyId & "'")
        
        While Not RS.EOF
        
            LiderId = RS!lider
            JugadorNombre = RS!nombre
            JugadorId = RS!id
        
            If Not IsNull(LiderId) And Not IsNull(JugadorId) Then
                
                If LiderId = JugadorId Then
                    Tmp = Tmp & CStr(JugadorNombre) & " (líder)" & SEPARATOR
                    Debug.Print Tmp
                Else
                    Tmp = Tmp & CStr(JugadorNombre) & SEPARATOR
                End If
            End If
            
            RS.MoveNext
        Wend
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
End Sub

Private Sub HandleOlvidaSpell(ByVal UserIndex As Integer)
Dim Slot As Long
    With UserList(UserIndex)

        Call .incomingData.ReadByte
        
        Slot = .incomingData.ReadLong
        Call WriteConsoleMsg(UserIndex, "¡Olvidaste " & Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).nombre & "!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).Stats.UserHechizos(Slot) = 0
        Call UpdateUserHechizos(False, UserIndex, CByte(Slot))
    End With
End Sub

Private Sub HandleSalirSubasta(ByVal UserIndex As Integer)
Dim Slot As Long
    With UserList(UserIndex)

        Call .incomingData.ReadByte
        
        UserList(UserIndex).flags.Comerciando = False
    End With
End Sub

Private Sub HandleNuevaSubasta(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Call .incomingData.ReadByte
        Call modSubastasDRAG.NuevaSubasta(UserIndex, .incomingData.ReadByte, .incomingData.ReadLong, .incomingData.ReadByte, .incomingData.ReadInteger)
        
    End With
End Sub

Public Sub WriteLimpiarACC(ByVal UserIndex As Integer)
'*****************************************************************************
'Author: Lorwik
'Last Modification: 20/03/2016
'Descripción: Le decimos al cliente que limpie la lista de PJ de la cuenta.
'****************************************************************************
    On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.limpiaracc)
    End With
    Exit Sub
errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteMultiMessage(ByVal UserIndex As Integer, ByVal id As Byte)
    '*****************************************************************************
'Author: Lorwik
'Last Modification: 26/03/2016
'Descripción: Le enviamos al cliente un ID para que lo identifique y
'muestre un mensaje en consola.
'****************************************************************************
    On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(id)
    End With
    Exit Sub
errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSubInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 18/11/2018
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.SubInit)
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCrearEfecto(ByVal UserIndex As Integer, ByVal EfectoIndex As Integer)
  On Error GoTo errHandler
  With UserList(UserIndex).outgoingData
    Call .WriteByte(ServerPacketID.CrearEfecto)
    Call .WriteInteger(EfectoList(EfectoIndex).id) 'id del efecto
    Call .WriteInteger(EfectoIndex) 'indice de su efecto en el EfectoList
    Call .WriteByte(EfectoList(EfectoIndex).tipo)
    Call .WriteInteger(EfectoList(EfectoIndex).duracion)
    Call .WriteBoolean(EfectoList(EfectoIndex).beneficioso)
    Call .WriteLong(EfectoList(EfectoIndex).grh)
  End With
  Exit Sub
errHandler:
     If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteQuitarEfecto(ByVal UserIndex As Integer, ByVal EfectoIndex As Integer)
  On Error GoTo errHandler
  With UserList(UserIndex).outgoingData
    Call .WriteByte(ServerPacketID.QuitarEfecto)
    Call .WriteInteger(EfectoIndex)
  End With
  Exit Sub
errHandler:
     If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub



Public Sub WriteCrearZona(ByVal UserIndex As Integer, ByVal ZonaIndex As Integer)
  On Error GoTo errHandler
  With UserList(UserIndex).outgoingData
    Call .WriteByte(ServerPacketID.CrearZona)
    Call .WriteLong(ZonaIndex)
    Call .WriteByte(ZonaList(ZonaIndex).x1)
    Call .WriteByte(ZonaList(ZonaIndex).x2)
    Call .WriteByte(ZonaList(ZonaIndex).y1)
    Call .WriteByte(ZonaList(ZonaIndex).y2)
    Call .WriteInteger(ZonaList(ZonaIndex).permisos)
    Call .WriteLong(ZonaList(ZonaIndex).grh)
  End With
  Exit Sub
errHandler:
     If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
