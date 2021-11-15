Attribute VB_Name = "ModKits"
Option Explicit
 
'Oro por equipo ganador (se divide por la cantidad de jugadores)
Const KITS_GOLD_TO_WINNER    As Long = 5000000
'Ejemplo = 5kk dividido 5 jugadores = 1kk por player :P
 
'Constante para enviar a todos los usuarios del evento un mensaje
Const KITS_ALL_MSG           As Byte = 0
 
'Mapa donde se realiza el evento.
Const KITS_MAP               As Integer = 18
 
'Posicionex X-Y para el equipo uno.
Const KITS_X1                As Byte = 50
Const KITS_Y1                As Byte = 24
 
'Posicionex X-Y para el equipo dos.
Const KITS_X2                As Byte = 49
Const KITS_Y2                As Byte = 75
 
'Jugadores por equipo.
Const KITS_USER              As Byte = 5
 
'Segundos para revivir usuarios.
Const KITS_RESPAWN           As Byte = 3
 
'Maximas muertes por equipo (en lo que consta este evento)
Const KITS_MAX_KILLS         As Byte = 20
 
Type Kits
     Users(1 To KITS_USER)   As Integer      'Usuarios.
     Kills                   As Byte         'Cuantas muertes tiene.
     Counters                As Byte         'Para los cupos.
End Type
 
Type kitsEvent
     KitsNum(1 To 2)         As Kits         'Index para los equipos.
     QuotaMax                As Byte         'Maximos cupos (no vamos contando con esto)
     KillsMax                As Byte         'Cuantas muertes por equipo.
     EventEnabled            As Boolean      'Si hay evento.
     EventStarted            As Boolean      'Si ya empezó.
End Type
 
Public kitsEvent             As kitsEvent
 
Sub KitsClear()
 
' \ author : maTih.-
' \ Note   : Limpia todo tipo de variables y uso de este sistema.
 
Dim loopC   As Long
 
With kitsEvent
   
    .EventStarted = False
    .EventEnabled = False
    .KillsMax = 0
    .QuotaMax = 0
   
    'Limpiamos los equipos.
   
    'Equipo #1
    With .KitsNum(1)
         .Counters = 0
         .Kills = 0
         
         For loopC = LBound(.Users()) To UBound(.Users())
         .Users(loopC) = 0
         Next loopC
         
    End With
   
    'Limpio el equipo #2
    With .KitsNum(2)
         .Counters = 0
         .Kills = 0
         
         For loopC = LBound(.Users()) To UBound(.Users())
         .Users(loopC) = 0
         Next loopC
         
    End With
 
End With
 
'Limpiar mapa.
 
ModKits.KitsClearMap
End Sub
 
Sub KitsClearMap()
 
' \ author : maTih.-
' \ Note   : Limpia objetos del mapa.
 
Dim LoopX   As Long
Dim LoopY   As Long
 
For LoopX = 1 To 100
    For LoopY = 1 To 100
        'Hay un objeto?
        If MapData(KITS_MAP, LoopX, LoopY).ObjInfo.ObjIndex > 0 Then
            'Borramos.
            EraseObj MapData(KITS_MAP, LoopX, LoopY).ObjInfo.Amount, KITS_MAP, LoopX, LoopY
        End If
    Next LoopY
Next LoopX
 
End Sub
 
Sub KitsStart()
 
' \ author : maTih.-
' \ Note   : Empieza el evento é inicializa variables.
 
ModKits.KitsClear
 
With kitsEvent
     .EventEnabled = True
     
     .KillsMax = KITS_MAX_KILLS
     .QuotaMax = KITS_USER
     
End With
 
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento automático : Pelea de bandos dió inicio, para participar tipea /PARTICIPAR" & vbNewLine & " El cupo es de " & kitsEvent.QuotaMax & " jugadores.", FontTypeNames.FONTTYPE_CITIZEN)
 
End Sub
 
Sub KitsGoTo()
 
' \ author : maTih.-
' \ Note   : Empieza el evento.
 
Dim loopC       As Long
Dim iUserIndex  As Integer
 
With kitsEvent
   
    .EventStarted = True
   
    'Llevamos a los usuarios del equipo uno a su base.
    With .KitsNum(1)
        For loopC = 1 To KITS_USER
            iUserIndex = .Users(loopC)
           
            If ModKits.KitsUserValid(iUserIndex) Then
                Call WarpUserChar(iUserIndex, KITS_MAP, KITS_X1 + loopC, KITS_Y1, True)
            End If
           
        Next loopC
    End With
   
    'Llevamos a los usuarios del equipo dos a su base.
    With .KitsNum(2)
        For loopC = 1 To KITS_USER
            iUserIndex = .Users(loopC)
           
            If ModKits.KitsUserValid(iUserIndex) Then
                Call WarpUserChar(iUserIndex, KITS_MAP, KITS_X1 + loopC, KITS_Y1, True)
            End If
           
        Next loopC
    End With
   
    'Avisamos.
    ModKits.KitsMessage KITS_ALL_MSG, "Empieza la batalla!! " & vbNewLine & "Muertes máximas por equipo : " & .KillsMax
   
    'Seteamos que el evento ya inicio.
    .EventStarted = True
   
End With
 
End Sub
 
Sub KitsKillUser(ByVal userIndex As Integer)
 
' \ author : maTih.-
' \ Note   : Muere un usuario.
 
Dim userKit As Byte
 
userKit = UserList(userIndex).Event.userKit
 
With kitsEvent
     
     'Sumamos el contador.
     .KitsNum(userKit).Kills = .KitsNum(userKit).Kills + 1
     
     'Avisamos cuantas muertes le quedan al equipo.
     ModKits.KitsMessage userKit, UserList(userIndex).name & " Ha muerto!! " & vbNewLine & " Al equipo le quedan : " & (KITS_MAX_KILLS - .KitsNum(userKit).Kills) & " Muertes!"
   
     'Limite de muertes?
     If .KitsNum(userKit).Kills >= KITS_MAX_KILLS Then
        'Perdieron :P
        ModKits.KitsManageKits userKit
        Exit Sub
    End If
   
    'Si no perdieron , seteamos los segundos de resu.
   
    UserList(userIndex).Event.secondRelive = KITS_RESPAWN
   
    WriteConsoleMsg userIndex, "Has muerto! volverás a la vida en " & KITS_RESPAWN & " segundos.", FontTypeNames.FONTTYPE_CITIZEN
   
End With
 
End Sub
 
Sub KitsManageKits(ByVal kitDead As Byte)
 
' \ author : maTih.-
' \ Note   : Acciones que dan fin al evento , setea ganadores y perdedores.
 
Dim kWinner     As Byte
 
'Obtengo el equipo que ganó.
kWinner = ModKits.KitsGiveWinner(kitDead)
 
'Acciones para los ganadores.
ModKits.KitsWin kWinner
'Acciones para los perdedores.
ModKits.KitsWarpLoosers kitDead
 
'Limpio todo.
ModKits.KitsClear
 
End Sub
 
Sub KitsGoSeconds(ByVal userIndex As Integer)
 
' \ author : maTih.-
' \ Note   : Contador para revivir usuarios.
 
Dim MiPos   As WorldPos
 
With UserList(userIndex).Event
     
     If UserList(userIndex).flags.Muerto <> 1 Then Exit Sub
     
     'Encontramos a donde hay que llevarlo
     ModKits.KitsPos MiPos, .userKit
     
     'Restamos el tiempo
         
     If .secondRelive > 0 Then .secondRelive = .secondRelive - 1
     
     'Llego a 0? respawn!
     
     If .secondRelive <= 0 Then
        RevivirUsuario userIndex
        'Lleno la vida
        UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MaxHP
        WriteUpdateHP userIndex
    End If
     
    WriteConsoleMsg userIndex, "Has vuelto a la vida!", FontTypeNames.FONTTYPE_CITIZEN
   
    'Aviso al team
    ModKits.KitsMessage .userKit, UserList(userIndex).name & " Volvió a la vida!"
   
    'Telep a la base.
   
    WarpUserChar userIndex, MiPos.Map, MiPos.X, MiPos.Y, True
   
End With
 
End Sub
 
Sub KitsSubmit(ByVal userIndex As Integer)
 
' \ author : maTih.-
' \ Note   : Inscribe a un usuario al evento.
 
Dim toKit   As Byte
 
toKit = ModKits.KitsGiveTeam
 
With kitsEvent
   
    'Lleno los datos con el del nuevo usuario
    'Y sieeeempre y cuando "ToKit" sea #2
    'Y El contador llege al maximo,
    'Doy inicio con el evento =)
   
    .KitsNum(toKit).Counters = .KitsNum(toKit).Counters + 1
    .KitsNum(toKit).Users(.KitsNum(toKit).Counters) = userIndex
   
    ModKits.KitsMessage toKit, UserList(userIndex).name & " Se inscribio para nuestro lado!"
   
    If toKit = 2 Then
       If .KitsNum(2).Counters >= KITS_USER Then
          ModKits.KitsGoTo
        End If
    End If
   
End With
 
With UserList(userIndex).Event
     .inEvent = True
     .secondRelive = 0
     .userKit = toKit
End With
 
End Sub
 
Sub KitsWin(ByVal kitWinner As Byte)
 
' \ author : maTih.-
' \ Note   : Gana un equipo y termina el evento.
 
Dim loopC           As Long
Dim GoldToPlayer    As Long
Dim iUserIndex      As Integer
 
With kitsEvent
 
     GoldToPlayer = ModKits.KitsGoldToPlayer(kitWinner)
   
    For loopC = 1 To KITS_USER
   
    With .KitsNum(kitWinner)
         iUserIndex = .Users(loopC)
         
         If ModKits.KitsUserValid(iUserIndex) Then
            'Doy el oro.
            UserList(iUserIndex).Stats.GLD = UserList(iUserIndex).Stats.GLD + GoldToPlayer
            'Actualizo el cliente del usuario.
            WriteUpdateGold iUserIndex
            'Lo llevo a su hogar.
            WarpUserChar iUserIndex, 1, 41, 88, True
        End If
           
    End With
   
    Next loopC
End With
 
'Chaaau nos vemos.
 
SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento automático : Guerra de bandos , GANADOR EQUIPO #" & kitWinner, FontTypeNames.FONTTYPE_CITIZEN)
 
End Sub
 
Sub KitsWarpLoosers(ByVal Looser As Byte)
 
' \ author : maTih.-
' \ Note   : Se lleva a los perdedores a su hogar
 
Dim loopC   As Long
Dim iUser   As Integer
 
For loopC = 1 To KITS_USER
    iUser = kitsEvent.KitsNum(Looser).Users(loopC)
   
    'Si es un usuario válido, busco su hogar y lo teletransporto al mismo.
    If ModKits.KitsUserValid(iUser) Then
       WarpUserChar iUser, 1, 41, 88, True
       'Si está muerto lo revivo.
       If UserList(iUser).flags.Muerto <> 0 Then
          RevivirUsuario iUser
          'Lleno la vida
          UserList(iUser).Stats.MinHP = UserList(iUser).Stats.MaxHP
          WriteUpdateHP iUser
        End If
    End If
   
Next loopC
 
End Sub
 
Function KitsGoldToPlayer(ByVal giveKit As Byte) As Long
 
' \ author : maTih.-
' \ Note   : Encuentra cuantos jugadores hay y divide el oro
 
Dim loopC   As Long
Dim iUser   As Integer
Dim aUsers  As Byte
 
For loopC = 1 To KITS_USER
    iUser = kitsEvent.KitsNum(giveKit).Users(loopC)
   
    'Sumamos el contador.
    If ModKits.KitsUserValid(iUser) Then
        aUsers = aUsers + 1
    End If
   
Next loopC
 
KitsGoldToPlayer = (KITS_GOLD_TO_WINNER / aUsers)
 
End Function
 
Sub KitsMessage(ByVal toKit As Byte, ByRef sMessage As String)
 
' \ author : maTih.-
' \ Note   : Envia un mensaje a un equipo o a todos los del evento.
 
Dim loopC       As Long
Dim K_FONT      As FontTypeNames
Dim iUserIndex  As Integer
 
K_FONT = FontTypeNames.FONTTYPE_WARNING
 
Select Case toKit
   
    Case 0              'Todos los del evento.
   
        SendData SendTarget.toMap, KITS_MAP, PrepareMessageConsoleMsg(sMessage, K_FONT)
   
    Case 1 To 2         'Un equipo específico.
   
        For loopC = 1 To KITS_USER
            With kitsEvent.KitsNum(toKit)
            iUserIndex = .Users(loopC)
                    If ModKits.KitsUserValid(iUserIndex) Then
                        WriteConsoleMsg iUserIndex, "[MENSAJE AL EQUIPO #" & toKit & "] : " & sMessage, K_FONT
                    End If
            End With
        Next loopC
   
End Select
 
End Sub
 
Sub KitsPos(ByRef MiPos As WorldPos, ByVal kitNum As Byte)
 
' \ author : maTih.-
' \ Note   : Toma el numero de equipo y guarda la posicion de su base en MiPos
 
MiPos.Map = KITS_MAP
 
If kitNum = 1 Then
    MiPos.X = KITS_X1
    MiPos.Y = KITS_Y1
Else
    MiPos.X = KITS_X2
    MiPos.Y = KITS_Y2
End If
 
End Sub
 
Function KitsGiveWinner(ByVal KitLooser As Byte) As Byte
 
' \ author : maTih.-
' \ Note   : Devuelve el equipo ganador según quien pierde.
 
If KitLooser = 1 Then
   KitsGiveWinner = 2
Else
    KitsGiveWinner = 1
End If
 
End Function
 
Function KitsGiveTag(ByVal userKit As Byte) As String
 
' \ author : maTih.-
' \ Note   : Devuelve el tag del usuario segun su equipo.
 
KitsGiveTag = " <Equipo #" & userKit & ">"
 
End Function
 
Function KitsUserIngress(ByVal userIndex As Integer, ByRef refError As String) As Boolean
 
' \ author : maTih.-
' \ Note   : Comprobaciones si puede ingresar al evento
 
KitsUserIngress = False
 
'Importante el orden de los condicionales
 
With UserList(userIndex)
 
     'Si no hay evento.
     If kitsEvent.EventEnabled <> True Then
        refError = "No hay ningún evento actualmente."
        Exit Function
    End If
   
    'Si hay evento pero ya empezó.
    If kitsEvent.EventStarted <> False Then
        refError = "El evento ya ha iniciado."
        Exit Function
    End If
   
    'Si está muerto.
    If .flags.Muerto <> 0 Then
        refError = "Estás muerto!!"
        Exit Function
    End If
   
    'Si está en carcel.
    If .Counters.Pena <> 0 Then
        refError = "Estás en la carcel!!"
        Exit Function
    End If
   
    KitsUserIngress = True
   
End With
 
End Function
 
Function KitsUserValid(ByVal userIndex As Integer) As Boolean
 
' \ author : maTih.-
' \ Note   : Devuelve si UserIndex , es <> 0 y si es un usuario logeado.
 
KitsUserValid = False
 
'Si no es diferente a 0.
If Not (userIndex <> 0) Then Exit Function
 
'Si no tiene IDValida.
 
If Not (UserList(userIndex).ConnID <> -1) Then Exit Function
 
'Si no es un usuario logeado (muy improbable llegar acá y que esto de false)
 
If Not (UserList(userIndex).flags.UserLogged) Then Exit Function
 
KitsUserValid = True
End Function
 
Function KitsGiveTeam() As Byte
 
' \ author : maTih.-
' \ Note   : Devuelve el equipo con menos jugadores (para inscribir usaurios)
 
With kitsEvent
 
'Si el equipo #1 tiene mas players que el #2 entonces
'Ingresar para el equipo #2.
If .KitsNum(1).Counters > .KitsNum(2).Counters Then
        KitsGiveTeam = 2
    Else        'Si no, entra al #1.
        KitsGiveTeam = 1
End If
 
End With
 
End Function
