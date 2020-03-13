Attribute VB_Name = "Mod_TileEngine"
Option Explicit

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Public movSpeed As Single

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Posicion en un mapa
Public Type Position
    x As Long
    y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    SX As Integer
    SY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
    
    Active As Boolean
    MiniMap_color As Long
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Long
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
        '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de ataque
Type AtaqueAnimData
    AtaqueWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Apariencia del personaje
Public Type char
    Active As Byte
    Heading As E_Heading
    Pos As Position
    moved As Boolean
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    Ataque As AtaqueAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    ParticulaIndex As Integer 'se usa

    particle_count As Integer
    particle_group() As Long
    
    Criminal As Byte
    
    Nombre As String
    ClanName As String
    PartyId As Integer
    NPCtype As Integer
    NPCID As Integer
    EstadoQuest As eEstadoQuest
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    Estainvi As Byte
    priv As Byte
    
    bType As Byte
    NPCAttack As Boolean
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    light_value(3) As Long
    
    luz As Integer
    color(3) As Long
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    blocked As Byte
    
    RenderValue As RVList
    
    Trigger As Integer
    particle_group_index As Long
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
    Ambient As String
End Type

Public Type D3D8Textures
    Texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

Public dX As DirectX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX As D3DX8

Public Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Public Type TLVERTEX2
    x As Single
    y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu1 As Single
    tv1 As Single
    tu2 As Single
    tv2 As Single
End Type

Public Const PI As Single = 3.14159265358979 'Numero PI
Public base_light As Long
Public LightIluminado(3) As Long
Public LightOscurito(3) As Long
Public NoPuedeUsar(3) As Long
Public TechoColor(3) As Long
Public AmbientColor As D3DCOLORVALUE

'**********************************
'CLIMAS Y HORARIOS (LUZ AMBIENTAL)
'**********************************
Type RGBClimax
    r As Byte
    g As Byte
    b As Byte
    A As Byte
End Type
 
Public ColorClimax As RGBClimax
'***************************

Public IniPath As String
Public MapPath As String

'Bordes del mapa.
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Dim timerElapsedTime As Single
Public timerTicksPerFrame As Single

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public AtaqueData() As AtaqueAnimData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bTecho       As Boolean 'hay techo?
Public bTechoAB As Byte

Public charlist(1 To 10000) As char

'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.x + viewPortX \ 32 - frmMain.renderer.ScaleWidth \ 64 - 1
    tY = UserPos.y + viewPortY \ 32 - frmMain.renderer.ScaleHeight \ 64
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .Active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .PartyId = 0
        .pie = False
        .Pos.x = 0
        .Pos.y = 0
        .UsandoArma = False
        .bType = 0
        Engine.Char_Particle_Group_Remove_All (CharIndex)
    End With
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer, ByVal Ataque As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        .Ataque = AtaqueData(Ataque)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        '[ANIM ATAK]
        .Arma.WeaponAttack = 0
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.x = x
        .Pos.y = y
        
        'Make active
        .Active = 1
    End With
    
    'Plot on map
    MapData(x, y).CharIndex = CharIndex
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next

With charlist(CharIndex)
    Call Engine.Char_Particle_Group_Remove_All(CharIndex)
   .Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    If .Pos.x = 0 Or .Pos.y = 0 Then Exit Sub
    
    MapData(.Pos.x, .Pos.y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End With
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)
On Error Resume Next
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
If GrhIndex = 0 Then Exit Sub
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
Static TerrenoDePaso As TipoPaso

    With charlist(CharIndex)
        If Not UserNavegando Then
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                    
                    If Not Char_Big_Get(CharIndex) Then
                        TerrenoDePaso = GetTerrenoDePaso(.Pos.x, .Pos.y)
                    ElseIf UserMontando = True Then
                        TerrenoDePaso = GetTerrenoDePaso(.Pos.x, .Pos.y)
                    Else
                        TerrenoDePaso = CONST_PESADO
                    End If
                    
                    If .pie = 0 Then
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(1), , Sound.Calculate_Volume(.Pos.x, .Pos.y), Sound.Calculate_Pan(.Pos.x, .Pos.y))
                    Else
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(2), , Sound.Calculate_Volume(.Pos.x, .Pos.y), Sound.Calculate_Pan(.Pos.x, .Pos.y))
                    End If
            End If
        Else
    ' TODO : Actually we would have to check if the CharIndex char is in the water or not....
            If Opciones.FxNavega = 1 Then Call Sound.Sound_Play(SND_NAVEGANDO)
        End If
    End With
End Sub

Private Function GetTerrenoDePaso(ByVal x As Byte, ByVal y As Byte) As TipoPaso
    With MapData(x, y).Graphic(1)
        If .GrhIndex >= 6000 And .GrhIndex <= 6307 Then
            GetTerrenoDePaso = CONST_BOSQUE
            Exit Function
        ElseIf .GrhIndex >= 7501 And .GrhIndex <= 7507 Or .GrhIndex >= 7508 And .GrhIndex <= 2508 Then
            GetTerrenoDePaso = CONST_DUNGEON
            Exit Function
        'ElseIf (TerrainFileNum >= 5000 And TerrainFileNum <= 5004) Then
        '    GetTerrenoDePaso = CONST_NIEVE
        '    Exit Function
        Else
            GetTerrenoDePaso = CONST_PISO
        End If
    End With
End Function

Public Function Char_Big_Get(ByVal CharIndex As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
   On Error GoTo ErrorHandler
   
   If UserMontando = True Then Exit Function
   
   'Make sure it's a legal char_index
    If Char_Check(CharIndex) Then
        Char_Big_Get = (GrhData(charlist(CharIndex).Body.Walk(charlist(CharIndex).Heading).GrhIndex).TileWidth > 4)
    End If
    
    Exit Function
    
ErrorHandler:
    
End Function

Private Function Char_Check(ByVal CharIndex As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If CharIndex > 0 And CharIndex <= LastChar Then
        Char_Check = (charlist(CharIndex).Heading > 0)
    End If
    
End Function

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim x As Integer
    Dim y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        MapData(x, y).CharIndex = 0
        
        addx = nX - x
        addy = nY - y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XXGRANDECIU Or .FxIndex = FxMeditar.XXGRANDECRI Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            y = -1
        
        Case E_Heading.EAST
            x = 1
        
        Case E_Heading.SOUTH
            y = 1
        
        Case E_Heading.WEST
            x = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.x - 8 To UserPos.x + 8
        For k = UserPos.y - 6 To UserPos.y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = 1521 Then '<Grh de la fogata 1521
                    location.x = j
                    location.y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While charlist(loopc).Active And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.


Function MoveToLegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Author: Lorwik
'Last Modify Date: 09/01/2011
'******************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, y).blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(x, y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.x, UserPos.y).blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.x, UserPos.y) Then
                    If Not HayAgua(x, y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(x, y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(x, y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Function HayUserAbajo(ByVal x As Integer, ByVal y As Integer, ByVal GrhIndex As Long) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.y <= y
    End If
End Function

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function


Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Long, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
End Sub

'**************************************************************
'MiniMapa
Public Sub ActualizarMiniMapa(ByVal tHeading As E_Heading)
'Esta es la forma mas optima que se me ha ocurrido. Solo dibuja una vez.
    frmMain.UserM.Left = UserPos.x - 1
    frmMain.UserM.Top = UserPos.y - 1
End Sub
Public Sub DibujarMiniMapa()

'Si el usuario esta en piramide, no dibujamos el minimapa
If UserMap = 32 Or UserMap = 33 Then
    frmMain.Minimap.Cls
    Exit Sub
End If

Dim map_x, map_y, Capas As Byte
    For map_y = 1 To 100
        For map_x = 1 To 100
        For Capas = 1 To 2
            If MapData(map_x, map_y).Graphic(Capas).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(Capas).GrhIndex).MiniMap_color
            End If
            If MapData(map_x, map_y).Graphic(4).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color
            End If
        Next Capas
        Next map_x
    Next map_y
   
    frmMain.Minimap.Refresh
    Call ActualizarMiniMapa(0)
End Sub
'***********************************************************

'***********************************************************
'CLIMAS Y HORARIO
'***********************************************************
Public Sub ClimaX()
'***************Lorwik/Climas*************************
'Descripción: Efectos climatologicos
'*****************************************************

    'Si el usuario esta muerto mostramos otro color
    If UserEstado = 1 Then ' Or Zona = "DUNGEON" Then 'Esto del dungeon hay que buscar otra forma, no me gusta nada.
        Call CalculateRGB(160, 160, 160, 1)
    Else
        Select Case Anocheceria
            'Mañana
            Case 0
                Call CalculateRGB(230, 200, 200, 255)
            'MedioDia
            Case 1
                Call CalculateRGB(255, 255, 255, 255)
            'Tarde
            Case 2
                Call CalculateRGB(200, 200, 200, 255)
            'Noche
            Case 3
                Call CalculateRGB(165, 165, 165, 1)
        End Select
    End If
End Sub

Public Sub CalculateRGB(r As Byte, g As Byte, b As Byte, A As Byte)
Dim i As Byte

With ColorClimax
    If .r < r Then
        .r = .r + 1
    Else
        .r = .r - 1
    End If
    If .b < b Then
        .b = .b + 1
    Else
        .b = .b - 1
    End If
    If .g < g Then
        .g = .g + 1
    Else
        .g = .g - 1
    End If
    If .A < A Then
        .A = .A + 1
    Else
        .A = .A - 1
    End If
    base_light = ARGB(.r, .g, .b, .A)
    For i = 0 To 3
        TechoColor(i) = base_light
    Next i
End With
End Sub

Public Sub InitColor()
'*******************************
'By Lorwik
'Iniciamos los colores
'*******************************
Dim i As Long
    bTechoAB = 255
    
    For i = 0 To 3
        LightIluminado(i) = RGB(255, 255, 255)
        LightOscurito(i) = RGB(150, 150, 150)
        NoPuedeUsar(i) = RGB(0, 0, 255)
    Next i
    
    AmbientColor.r = 200
    AmbientColor.g = 200
    AmbientColor.b = 200
    AmbientColor.A = 255
End Sub

