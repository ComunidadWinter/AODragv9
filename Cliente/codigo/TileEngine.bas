Attribute VB_Name = "Mod_TileEngine"
Option Explicit

Private MinimapMaxY As Byte
Private MinimapMaxX As Byte

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 200
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 200
Public Const YMinMapSize As Byte = 1

Public movSpeed As Single

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
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
Public Type grh
    GrhIndex As Long
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
    Angle As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.SOUTH To E_Heading.EAST) As grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.SOUTH To E_Heading.EAST) As grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.SOUTH To E_Heading.EAST) As grh
        '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.SOUTH To E_Heading.EAST) As grh
End Type

'Lista de las animaciones de ataque
Type AtaqueAnimData
    AtaqueWalk(E_Heading.SOUTH To E_Heading.EAST) As grh
    HeadOffset As Position
End Type

'Apariencia del personaje
Public Type char
    oldPos As WorldPos 'Esto es solo por si acaso...
    AnimTime As Byte
    Active As Byte
    Heading As E_Heading
    Pos As Position
    moved As Boolean
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As Integer
    Casco As Integer
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    Ataque As AtaqueAnimData
    UsandoArma As Boolean
    
    fX As grh
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
    Speed As Byte
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As grh
    CharIndex As Integer
    ObjGrh As grh
    
    light_value(3) As Long
    
    luz As Integer
    Color(3) As Long
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    blocked As Byte
    
    RenderValue As RVList
    
    trigger As Integer
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

'DX8 Objects
Public DirectX As New DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8

Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Public Type TLVERTEX2
    X As Single
    Y As Single
    z As Single
    rhw As Single
    Color As Long
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

'*********************************
'Particulas
'*********************************
Private base_tile_size As Integer

Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    Angle As Single
    grh As grh
    alive_counter As Long
    x1 As Long
    x2 As Long
    y1 As Long
    y2 As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    Radio As Integer
    rgb_list(0 To 3) As Long
End Type

Private Type Stream
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
End Type

Private Type particle_group
    Active As Boolean
    id As Long
    map_x As Long
    map_y As Long
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    x1 As Long
    x2 As Long
    y1 As Long
    y2 As Long
    Angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    rgb_list(0 To 3) As Long
    
    Speed As Single
    life_counter As Long
    
    Radio As Integer
    
    Bind_To_Char    As Integer
    Bind_Speed      As Single
    Now_Viaje_X     As Integer
    Now_Viaje_Y     As Integer
End Type
'Particle system
 
'Dim StreamData() As particle_group
Dim TotalStreams As Long
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long
Dim char_list() As char

Public engineBaseSpeed As Single
Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

Private timerElapsedTime As Single
Private timerTicksPerFrame As Double

Private lFrameTimer As Long
Private lFrameLimiter As Long

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

Private TileBufferSize As Byte

Private MouseTileX As Byte
Private MouseTileY As Byte
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

Dim bump_map_supported As Boolean

Private Type decoration
    grh As grh
    Render_On_Top As Boolean
    subtile_pos As Byte
End Type

Private Type Map_Tile
    grh(1 To 3) As grh
    decoration(1 To 5) As decoration
    decoration_count As Byte
    blocked As Boolean
    particle_group_index As Long
    char_index As Long
    light_base_value(0 To 3) As Long
    light_value(0 To 3) As Long
   
    exit_index As Long
    npc_index As Long
    item_index As Long
   
    trigger As Byte
End Type

Private Type Map
    map_grid() As Map_Tile
    map_x_max As Long
    map_x_min As Long
    map_y_max As Long
    map_y_min As Long
    map_description As String
    'Added by Juan Martín Sotuyo Dodero
    base_light_color As Long
End Type

Dim map_current As Map

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

Public EngineRun As Boolean

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Private MainViewWidth As Integer
Private MainViewHeight As Integer

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

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Dim Texture As Direct3DTexture8
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmMain.renderer.ScaleWidth \ 64 - 1
    tY = UserPos.Y + viewPortY \ 32 - frmMain.renderer.ScaleHeight \ 64
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
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
        .bType = 0
        Char_Particle_Group_Remove_All (CharIndex)
    End With
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer, ByVal Ataque As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    If charlist(CharIndex).Active Then Exit Sub
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        If .oldPos.X > 0 Then _
           MapData(.oldPos.X, .oldPos.Y).CharIndex = 0
        
        .iHead = Head
        .iBody = Body
        .Head = Head
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        .Ataque = AtaqueData(Ataque)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = Casco
        
        '[ANIM ATAK]
        .Arma.WeaponAttack = 0
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        .oldPos.X = .Pos.X
        .oldPos.Y = .Pos.Y
        
        'Make active
        .Active = 1
        
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
    
    'Lorwik: Parche para los TP en el mismo mapa:
    'Si el cuerpo que se ha creado es el propio actualizamos el minimapa
    If CharIndex = UserCharIndex Then Call DibujarMiniMapa
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next

With charlist(CharIndex)
    Call Char_Particle_Group_Remove_All(CharIndex)
   .Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    If .Pos.X = 0 Or .Pos.Y = 0 Then Exit Sub
    
    MapData(.Pos.X, .Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End With
End Sub

Public Sub InitGrh(ByRef grh As grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)
On Error Resume Next
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
If GrhIndex = 0 Then Exit Sub
    grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(grh.GrhIndex).NumFrames > 1 Then
            grh.Started = 1
        Else
            grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(grh.GrhIndex).NumFrames = 1 Then Started = 0
        grh.Started = Started
    End If
    
    If grh.Started Then
        grh.Loops = INFINITE_LOOPS
    Else
        grh.Loops = 0
    End If
    
    grh.FrameCounter = 1
    grh.Speed = GrhData(grh.GrhIndex).Speed
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
Static TerrenoDePaso As TipoPaso

    With charlist(CharIndex)
        If Not UserNavegando Then
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                    
                    If Not Char_Big_Get(CharIndex) Then
                        TerrenoDePaso = GetTerrenoDePaso(.Pos.X, .Pos.Y)
                    ElseIf UserMontando = True Then
                        TerrenoDePaso = GetTerrenoDePaso(.Pos.X, .Pos.Y)
                    Else
                        TerrenoDePaso = CONST_PESADO
                    End If
                    
                    If .pie = 0 Then
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(1), , Sound.Calculate_Volume(.Pos.X, .Pos.Y), Sound.Calculate_Pan(.Pos.X, .Pos.Y))
                    Else
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(2), , Sound.Calculate_Volume(.Pos.X, .Pos.Y), Sound.Calculate_Pan(.Pos.X, .Pos.Y))
                    End If
            End If
        Else
    ' TODO : Actually we would have to check if the CharIndex char is in the water or not....
            If Opciones.FxNavega = 1 Then Call Sound.Sound_Play(SND_NAVEGANDO)
        End If
    End With
End Sub

Private Function GetTerrenoDePaso(ByVal X As Byte, ByVal Y As Byte) As TipoPaso
    With MapData(X, Y).Graphic(1)
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
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDECIU Or .FxIndex = FxMeditar.XXGRANDECRI Then
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
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        Case E_Heading.EAST
            X = 1
        Case E_Heading.SOUTH
            Y = 1
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = 1521 Then  '<Grh de la fogata 1521
                    location.X = j
                    location.Y = k
                    
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


Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: Lorwik
'Last Modify Date: 09/01/2011
'******************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Long) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Private Sub DrawHead(ByVal X As Integer, ByVal Y As Integer, ByVal EsCabeza As Boolean, Light() As Long, ByVal Heading As Byte, ByVal Head As Integer)
    Dim Cabezoide As grh
    Dim textureX1 As Integer
    Dim textureX2 As Integer
    Dim textureY1 As Integer
    Dim textureY2 As Integer
    Dim offsetX As Integer
    Dim offsetY As Integer
    Dim SourceRect As RECT
    Dim Texture As Long
    
        If EsCabeza = True Then
            If heads(Head).Texture <= 0 Then Exit Sub
            Texture = heads(Head).Texture
        Else
            If Cascos(Head).Texture <= 0 Then Exit Sub
            Texture = Cascos(Head).Texture
        End If
        
        textureX2 = 27
        textureY2 = 32
 
        If EsCabeza = True Then
            textureX1 = heads(Head).startX
            textureY1 = ((Heading - 1) * textureY2) + heads(Head).startY
        Else
            textureX1 = Cascos(Head).startX
            textureY1 = ((Heading - 1) * textureY2) + Cascos(Head).startY
        End If
 
        offsetX = (textureX2) - 30
        offsetY = (textureY2) - 35
        
        With SourceRect
            .Left = textureX1
            .Top = textureY1
            .Right = (textureX2 + .Left)
            .Bottom = (textureY2 + .Top)
        End With
        
        Device_Textured_Render X - offsetX, Y - offsetY, _
        SurfaceDB.Surface(Texture), SourceRect, Light

End Sub


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
Public Sub ActualizarMiniMapa()
'*******************************************
'Autor: Lorwik
'Ultima modificacion: 13/12/2018
'Añadido soporte para mapas de 200x200
'*******************************************

'Esta es la forma mas optima que se me ha ocurrido. Solo dibuja una vez.
    
    If UserPos.X < 100 Then
        frmMain.UserM.Left = UserPos.X - 1
        MinimapMaxX = XMaxMapSize - 100
    ElseIf UserPos.X > 100 Then
        frmMain.UserM.Left = UserPos.X - 101
        MinimapMaxX = XMaxMapSize
    End If
    
    If UserPos.Y < 100 Then
        frmMain.UserM.Top = UserPos.Y - 1
        MinimapMaxY = YMaxMapSize - 100
    ElseIf UserPos.Y > 100 Then
        frmMain.UserM.Top = UserPos.Y - 101
        MinimapMaxY = YMaxMapSize
    End If
    
End Sub
Public Sub DibujarMiniMapa()
'*******************************************
'Autor: Lorwik
'Ultima modificacion: 13/12/2018
'Añadido soporte para mapas de 200x200
'*******************************************

    Dim MinimapMinX, MinimapMinY, Capas, map_y, map_x, X, Y As Byte
    
    'Calculamos el minimo del trozo de mapa a renderizar
    If MinimapMaxY = 200 Then
        MinimapMinY = YMinMapSize + 99
    ElseIf MinimapMaxY = 100 Then
        MinimapMinY = YMinMapSize
    End If
    
    If MinimapMaxX = 200 Then
        MinimapMinX = YMinMapSize + 99
    ElseIf MinimapMaxX = 100 Then
        MinimapMinX = XMinMapSize
    End If
    
    'Si el usuario esta en piramide, no dibujamos el minimapa
    If UserMap = 32 Or UserMap = 33 Then
        frmMain.Minimap.Cls
        Exit Sub
    End If
    
    For map_y = MinimapMinY To MinimapMaxY
        For map_x = MinimapMinX To MinimapMaxX
        For Capas = 1 To 2
        
            'Si la X y la Y es mayor a 100 tenemos que restarle 100 par aque no se salga del control
            If map_x > 100 Then
                X = map_x - 100
            ElseIf map_x < 100 Then
                X = map_x
            End If
            
            If map_y > 100 Then
                Y = map_y - 100
            ElseIf map_y - 100 Then
                Y = map_y
            End If
            
            If MapData(map_x, map_y).Graphic(Capas).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, X - 1, Y - 1, GrhData(MapData(map_x, map_y).Graphic(Capas).GrhIndex).MiniMap_color
            End If
            If MapData(map_x, map_y).Graphic(4).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, X - 1, Y - 1, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color
            End If
        Next Capas
        Next map_x
    Next map_y
   
    frmMain.Minimap.Refresh
    Call ActualizarMiniMapa
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
        Select Case DayStatus
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

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    Dim Speed As Byte
    
    Static re As RECT
    With re
        .Left = 0
        .Top = 0
        .Bottom = frmMain.renderer.ScaleHeight
        .Right = frmMain.renderer.ScaleWidth
    End With
    
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    DirectDevice.BeginScene
        
    If UserMoving Then
    
        Speed = charlist(UserCharIndex).Speed
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = OffsetCounterX - (ScrollPixelsPerFrameX + Speed) * AddtoUserPos.X * timerTicksPerFrame
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = False
            End If
        End If
            
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - (ScrollPixelsPerFrameY + Speed) * AddtoUserPos.Y * timerTicksPerFrame
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = False
            End If
        End If
    End If
        
    'Update mouse position within view area
    Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
    '****** Update screen ******
    If UserCiego Then
        'Call CleanViewPort
    Else
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
    End If
        
        '*********************FPS**********************************
    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
    If FPSFLAG Then texto.Engine_Text_Draw 645, 5, "FPS: " & FPS, vbWhite
    '**********************************************************
    
    '**********************Nombre del mapa***********************************
    If Not MapInfo.name = LastMapName Then
        If TransMapAB > 0 Then
            TransMapAB = TransMapAB - 1
        End If

        texto.Engine_Text_Draw 360, 50, MapInfo.name, vbWhite, TransMapAB, True, 2
        If MapDat.battle_mode = True Then
            texto.Engine_Text_Draw 350, 120, "¡Estas en zona insegura!", vbBlue, TransMapAB, True
        Else
            texto.Engine_Text_Draw 350, 120, "Estas en zona segura", vbGreen, TransMapAB, True
        End If
    End If
    '*************************************************************************
    
    Call Dialogos.Render
    Call DibujarCartel
        
    'FPS update
    If fpsLastCheck + 1000 < GetTickCount Then
        'FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        fpsLastCheck = GetTickCount
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
        
    'Get timing info
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    FPS = 1000 / timerElapsedTime
        
    DirectDevice.EndScene
    DirectDevice.Present re, ByVal 0, 0, ByVal 0
End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Dibuja todo el decorado, PJ, Bichos y lo que le metamos...
'**************************************************************
    Dim Y                   As Integer     'Keeps track of where on map we are
    Dim X                   As Integer     'Keeps track of where on map we are
    Dim screenminY          As Integer  'Start Y pos on current screen
    Dim screenmaxY          As Integer  'End Y pos on current screen
    Dim screenminX          As Integer  'Start X pos on current screen
    Dim screenmaxX          As Integer  'End X pos on current screen
    Dim minY                As Integer  'Start Y pos on current map
    Dim maxY                As Integer  'End Y pos on current map
    Dim minX                As Integer  'Start X pos on current map
    Dim maxX                As Integer  'End X pos on current map
    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset          As Integer
    Dim minYOffset          As Integer
    Dim PixelOffsetXTemp    As Integer 'For centering grhs
    Dim PixelOffsetYTemp    As Integer 'For centering grhs
    Dim CurrentGrhIndex     As Integer
    Dim offx                As Integer
    Dim offy                As Integer
    Dim i As Long
    
On Error Resume Next
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1

    'Lorwik> Con esto evitamos que tire error al entrar en el borde del mapa:
    If screenmaxX > XMaxMapSize Then screenmaxX = XMaxMapSize
    If screenmaxY > YMaxMapSize Then screenmaxY = YMaxMapSize
    If screenminX < XMinMapSize Then screenmaxX = XMinMapSize
    If screenminY < YMinMapSize Then screenmaxY = YMinMapSize
    
    '=============================================================================================================
    'Comenzamos a dibujar las capas. Aquellas capas que esten por debajo de las demas, se van a mostrar arriba.
    'Ejemplo: La capa 2 se mostrara por debajo de la 3 y la 4 se mostrara por encima de la 3.
    '=============================================================================================================
    
    'Draw floor layer
    
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            'Capa 1 **********************************
            
            Call DDrawGrhtoSurface(MapData(X, Y).Graphic(1), _
                (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                0, 1, MapData(X, Y).light_value)
            '******************************************
            

            'Capa 2 ************************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then _
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), _
                    (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                    (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                    1, 1, MapData(X, Y).light_value)
            '*******************************************
            
             '15/12/2018 Irongete: Pintar el Grh de la zona
          Dim ZonaIndex As Long
          For ZonaIndex = 0 To UBound(ZonaList)
            '15/12/2018 Irongete: Está este tile dentro de la zona?
            If X >= ZonaList(ZonaIndex).x1 And X <= ZonaList(ZonaIndex).x2 Then
              If Y >= ZonaList(ZonaIndex).y1 And Y <= ZonaList(ZonaIndex).y2 Then
                Call DDrawTransGrhtoSurface(ZonaList(ZonaIndex).grh, (ScreenX - 1) * TilePixelWidth + PixelOffsetX, (ScreenY - 1) * TilePixelHeight + PixelOffsetY, 1, 1, MapData(X, Y).light_value)
              End If
            End If
          Next
            
            
            
            ScreenX = ScreenX + 1
        Next X
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(X, Y)
                '******************************************

                '**************************************************
                'Capa de objetos
                'Dibujara solo los objetos
                '**************************************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value)
                End If
                
                '***************************************************
                'Capa de Char
                'Dibuja todo aquello que tenga cuerpo (PJ y Bichos)
                '***************************************************
                If .CharIndex <> 0 Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, X, Y, MapData(X, Y).light_value)
                End If
                '*************************************************
                
                'Dibujamos el valor en el render.
                If .RenderValue.Activated Then
                    modRenderValue.Draw X, Y, PixelOffsetXTemp + 20, PixelOffsetYTemp - 30
                End If
                
                'Capa 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value)
                End If
                '************************************************

            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5
    
    'Particulas ************************************************
    ScreenY = minYOffset - TileBufferSize
        For Y = minY To maxY
            ScreenX = minXOffset - TileBufferSize
                For X = minX To maxX
                    'Particulas**************************************
                        If MapData(X, Y).particle_group_index Then _
                            Particle_Group_Render MapData(X, Y).particle_group_index, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY
                    '************************************************
                ScreenX = ScreenX + 1
            Next X
        ScreenY = ScreenY + 1
    Next Y
    '***********************************************************

    'Capa4 - Techos******************************
    
    '18/02/2016 Irongete: Si está debajo de techo bajamos la transparencia a cero, si está fuera de techo la bajamos hasta lo que tenga en la configuración
    If bTecho And bTechoAB > 0 Then
        bTechoAB = bTechoAB - 1
    Else
        If bTechoAB < Opciones.BaseTecho Then
            bTechoAB = bTechoAB + 1
        ElseIf bTechoAB > 0 Then
            bTechoAB = bTechoAB - 1
        End If
    End If
    
    
    'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            If Not MapData(X, Y).Graphic(4).GrhIndex Then
                'Draw
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                    ScreenX * TilePixelWidth + PixelOffsetX, _
                    ScreenY * TilePixelHeight + PixelOffsetY, _
                    1, 1, TechoColor(), bTechoAB) 'Techo normal
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    '*******************************************
        
    'Clima**************************************
    Call ClimaX
    '*******************************************
    
    '15/12/2018 Irongete: Dibujar efectos
    Dim EfectoIndex As Long
    Dim offset_beneficioso As Integer
    Dim offset_perjudicial As Integer
    offset_beneficioso = 0
    offset_perjudicial = 0
    
    For EfectoIndex = 0 To UBound(EfectoList)
      If EfectoList(EfectoIndex).EfectoIndex > 0 Then
      
        If EfectoList(EfectoIndex).beneficioso = True Then
          Call DDrawGrhtoSurface(EfectoList(EfectoIndex).grh, 5 + offset_beneficioso, 5, 0, 1, MapData(X, Y).light_value)
          offset_beneficioso = i + 50
        End If
        
        If EfectoList(EfectoIndex).beneficioso = False Then
          Call DDrawGrhtoSurface(EfectoList(EfectoIndex).grh, 650 - offset_perjudicial, 5, 0, 1, MapData(X, Y).light_value)
          offset_perjudicial = offset_perjudicial + 50
        End If

      End If
    Next
 
End Sub

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal X As Byte, ByVal Y As Byte, Light() As Long)
'*******************************************************
'Esto forma parte del RenderScreen.
'Dibuja todo aquello que tenga cuerpo (por asi decirlo)
'Bichos y PJ
'*******************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim Color As Long
    Dim TempName As String
    Dim TempClanName As String
    Dim PartyConfirmed As Boolean, ClanConfirmed As Boolean
    Dim PixelOffsetYTemp As Integer
    
    
    With charlist(CharIndex)
    
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + (ScrollPixelsPerFrameX + .Speed) * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Inicia la animacion
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Movemos el Char
                moved = True
                .AnimTime = 10
                
                'Comprueba si ya llegamos
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + (ScrollPixelsPerFrameY + .Speed) * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Inicia la animacion
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!! <<<-- Lorwik> Esto no lo entiendo
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Movemos el Char
                moved = True
                .AnimTime = 10
                
                'Comprueba si ya llegamos
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'Si se detiene, paramos la animacion
        If Not moved Then
            'Stop animations
            If .AnimTime = 0 Then
                .Body.Walk(.Heading).Started = 0
                .Body.Walk(.Heading).FrameCounter = 1
                
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
                
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 1
                
                If .NPCAttack = False Then
                    .Ataque.AtaqueWalk(.Heading).Started = 0
                    .Ataque.AtaqueWalk(.Heading).FrameCounter = 1
                End If
                
                .Moving = False
            Else
                .AnimTime = .AnimTime - 1
            End If
        Else
            IsAttacking = False
            .NPCAttack = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        '¿Tiene cabeza?
        If charlist(CharIndex).Head Then
            movSpeed = 0.8
            'Dibujamos el cuerpo
            If .Body.Walk(.Heading).GrhIndex Then _
                Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, .Estainvi)
                
            'Dibujamos la Cabeza
            If .Head Then
                Call DrawHead(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, True, Light, charlist(CharIndex).Heading, charlist(CharIndex).Head)

                'Draw Helmet
                If .Casco Then _
                    Call DrawHead(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, False, Light, charlist(CharIndex).Heading, charlist(CharIndex).Casco)
                             
                If UserMontando = False Then
                    'Dibujamos el arma
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, .Estainvi)
                    'Dibujamos el escudo
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, .Estainvi)
                End If
            End If
        Else
            If .NPCAttack = True And .Ataque.AtaqueWalk(.Heading).GrhIndex > 0 Then
                If .Ataque.AtaqueWalk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Ataque.AtaqueWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, .Estainvi)
            Else
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, .Estainvi)
            End If
        End If
        
        If .invisible = False Then
            'Dibujamos el nombre debajo de la cabeza.
            If Opciones.NamePlayers Then
                If Len(.Nombre) > 0 Then
                    Pos = InStr(.Nombre, "<")
                    If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                    If .priv = 0 Then
                        Select Case .Criminal
                            Case 0
                                Color = ColoresPJ(49)
                            Case 1
                                Color = ColoresPJ(50)
                            Case 8
                                Color = ColoresPJ(4)
                            Case 9
                                Color = ColoresPJ(9)
                            Case 10
                                Color = ColoresPJ(3)
                        End Select
                    Else
                        Color = ColoresPJ(.priv)
                    End If
                            
                    'Nick
                    line = .Nombre
                    Call texto.Engine_Text_Draw(PixelOffsetX + 15, PixelOffsetY + 30, line, Color, , True)
                            
                    If Not .ClanName = "" Then
                        'Clan
                        line = "<" & .ClanName & ">"
                        Call texto.Engine_Text_Draw(PixelOffsetX + 15, PixelOffsetY + 45, line, RGB(157, 202, 231), , True)
                    End If
                    If .priv > 0 Then Call texto.Engine_Text_Draw(PixelOffsetX + 15, PixelOffsetY + 45, "<Administrador>", RGB(255, 255, 255), , True)
                End If
            End If
        End If
        
        PartyConfirmed = False
        ClanConfirmed = False
            
        TempClanName = .ClanName  '60 = <
            
        If .invisible And UCase$(UserNameClan) <> "" And UCase$(TempClanName) <> "" Then   'asi prevenimos que funcione por no tener clan
            If UCase$(UserNameClan) = UCase$(TempClanName) Then ClanConfirmed = True
        ElseIf .invisible And UCase$(UserNameClan) = vbNullString Then
            ClanConfirmed = False
        End If
        
        If .invisible And UserPartyId > 0 And .PartyId > 0 Then   'asi prevenimos que funcione por tener index=0
            If UserPartyId = .PartyId Then
                PartyConfirmed = True
            End If
        ElseIf .invisible And UserPartyId = 0 Then
            PartyConfirmed = False
        End If
            
           
        If .invisible And CharIndex = UserCharIndex Or ClanConfirmed = True Or PartyConfirmed = True Then
            .Estainvi = 130
        ElseIf Not .invisible Then
            .Estainvi = 255
        Else
            .Estainvi = 1
        End If
        
        'Actualizamos los dialogos
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex)   '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        movSpeed = 1
        
        '************Particulas************
        Dim i As Integer
            If .particle_count > 0 Then
                For i = 1 To .particle_count
                    If .particle_group(i) > 0 Then _
                        Particle_Group_Render .particle_group(i), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY
                Next i
            End If
        
        '*******Dibujamos los FX********
        If .FxIndex <> 0 Then
            Call DDrawTransGrhtoSurface(.fX, PixelOffsetX + FxData(.FxIndex).offsetX, PixelOffsetY + FxData(.FxIndex).offsetY, 1, 1, Light, 255, 1)
            
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
        
        '14/11/2018 Pinto el icono encima del NPC segun el estado en el que se encuentre el jugador con las quests
        If .EstadoQuest > 2 Then
          Select Case .EstadoQuest
            Case eEstadoQuest.NoAceptada
              'Call Draw_GrhIndex(2176, PixelOffsetX + 7, PixelOffsetY - 37, LightIluminado())
              
            Case eEstadoQuest.CompletadaFaltaEntregar
              'Call Draw_GrhIndex(2177, PixelOffsetX + 7, PixelOffsetY - 37, LightIluminado())
              
          End Select
        End If
        
    End With
End Sub

Public Sub RenderConnect()

    Dim X As Byte, Y As Byte
    
    Static re As RECT
    
    With re
        .Left = 0
        .Top = 0
        .Bottom = 768
        .Right = 1024
    End With
    
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    DirectDevice.BeginScene

    'Capa 1
    For X = 1 To 32
        For Y = 1 To 24
            With MapData(X + MapaConnect.X, Y + MapaConnect.Y)
                Call DDrawGrhtoSurface(.Graphic(1), _
                        (X - 1) * 32, (Y - 1) * 32, _
                         0, 1, MapData(X, Y).light_value)
                End With
        Next Y
    Next X
        
    'Capa 2
    For X = 1 To 32
        For Y = 1 To 24
            With MapData(X + MapaConnect.X, Y + MapaConnect.Y)
                Call DDrawGrhtoSurface(.Graphic(2), _
                                (X - 1) * 32, (Y - 1) * 32, _
                                1, 1, MapData(X, Y).light_value)
                End With
        Next Y
    Next X
        
    'Capa 3
    For X = 1 To 32
        For Y = 1 To 24
            With MapData(X + MapaConnect.X, Y + MapaConnect.Y)
                Call DDrawGrhtoSurface(.Graphic(3), _
                                (X - 1) * 32, (Y - 1) * 32, _
                                1, 1, MapData(X, Y).light_value)
                End With
        Next Y
    Next X
    
    If Not frmCuenta.ListPJ.ListIndex < 0 Then 'Con este If evitamos que tire error en caso de que no tengamos ningun pj seleccionado
        With Cuenta.pjs(frmCuenta.ListPJ.ListIndex + 1)
            Call DrawHead(180, 180 + BodyData(.Acuerpo).HeadOffset.Y, True, LightIluminado(), 1, .rcvHead)
            Call DDrawGrhtoSurface(BodyData(.Acuerpo).Walk(1), 180, 180, 1, 1, LightIluminado())
            texto.Engine_Text_Draw 195, 210, .NamePJ, vbWhite, , True
        End With
    End If
    
    'Capa 4
    For X = 1 To 32
        For Y = 1 To 24
            With MapData(X + MapaConnect.X, Y + MapaConnect.Y)
                Call DDrawGrhtoSurface(.Graphic(4), _
                                (X - 1) * 32, (Y - 1) * 32, _
                                1, 1, MapData(X, Y).light_value)
                End With
        Next Y
    Next X
        
    texto.Engine_Text_Draw 5, 750, "Version: " & App.Major & "." & App.Minor & "." & App.Revision, vbCyan
    
    If frmCuenta.Visible = True Then
        texto.Engine_Text_Draw 505, 600, ConsejoSeleccionado, vbWhite, , 1
        texto.Engine_Text_Draw 505, 630, "Siguiente consejo", &HC0FFFF, , 1
    End If
    
    'Clima**************************************
    Call ClimaX
        
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    FPS = 1000 / timerElapsedTime
                           
        
    DirectDevice.EndScene
    DirectDevice.Present re, ByVal 0, frmRenderConnect.hwnd, ByVal 0
End Sub

'Sets a Grh animation to loop indefinitely.

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

Private Sub DDrawGrhtoSurface(ByRef grh As grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef Light() As Long)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
On Error GoTo Error
        
    If grh.GrhIndex = 0 Then Exit Sub
        
    If Animate Then
        If grh.Started = 1 Then
            grh.FrameCounter = grh.FrameCounter + (timerElapsedTime * GrhData(grh.GrhIndex).NumFrames / grh.Speed)
            If grh.FrameCounter > GrhData(grh.GrhIndex).NumFrames Then
                grh.FrameCounter = (grh.FrameCounter Mod GrhData(grh.GrhIndex).NumFrames) + 1
                
                If grh.Loops <> INFINITE_LOOPS Then
                    If grh.Loops > 0 Then
                        grh.Loops = grh.Loops - 1
                    Else
                        grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, Light)
    End With
Exit Sub

Error:
    If Err.number = 9 And grh.FrameCounter < 1 Then
        grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Light() As Long, Optional ByVal Alpha As Boolean = False, Optional ByVal AlphaByte As Byte = 255, Optional ByVal Angle As Byte = 0)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        'Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, Light, Alpha, AlphaByte, Angle)
    End With
End Sub

Sub DDrawTransGrhtoSurface(ByRef grh As grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef Light() As Long, Optional Transp As Byte = 255, Optional blend As Byte = 0, Optional Angle As Byte = 0)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
'On Error GoTo error
    On Error Resume Next
    If grh.GrhIndex = 0 Then Exit Sub
    
    If Animate Then
        If grh.Started = 1 Then
            grh.FrameCounter = grh.FrameCounter + (timerElapsedTime * GrhData(grh.GrhIndex).NumFrames / grh.Speed) * movSpeed
            
            If grh.FrameCounter > GrhData(grh.GrhIndex).NumFrames Then
                grh.FrameCounter = (grh.FrameCounter Mod GrhData(grh.GrhIndex).NumFrames) + 1
                
                If grh.Loops <> INFINITE_LOOPS Then
                    If grh.Loops > 0 Then
                        grh.Loops = grh.Loops - 1
                    Else
                        grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
                
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, Light(), CBool(blend), Transp, Angle)
    End With
    
    
Exit Sub

Error:
    If Err.number = 9 And grh.FrameCounter < 1 Then
        grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean

movSpeed = 1
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    'Dim SurfaceDesc As DDSURFACEDESC2
    'Dim ddck As DDCOLORKEY
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = (frmMain.renderer.Height / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.Width / 32) \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    '***********************************
    'Tamaño del mapa
    '***********************************
    MinXBorder = XMinMapSize + (Round(frmMain.renderer.Width / 32) \ 2)
    MaxXBorder = XMaxMapSize - (Round(frmMain.renderer.Width / 32) \ 2)
    MinYBorder = YMinMapSize + (Round(frmMain.renderer.Height / 32) \ 2)
    MaxYBorder = YMaxMapSize - (Round(frmMain.renderer.Height / 32) \ 2)
    '***********************************
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
On Error GoTo 0
    
    frmCargando.Status.Caption = "Cargando Graficos...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(20)
    Call LoadGrhData
    frmCargando.Status.Caption = "Cargando Particulas..."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(25)
    Call CargarParticulas
    frmCargando.Status.Caption = "Cargando Minimapa...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(30)
    Call LoadMiniMap
    frmCargando.Status.Caption = "Cargando Cuerpos...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(40)
    Call CargarCuerpos
    frmCargando.Status.Caption = "Cargando Ataques...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(45)
    Call CargarAtaques
    frmCargando.Status.Caption = "Cargando Cabezas...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(50)
    Call CargarCabezas
    frmCargando.Status.Caption = "Cargando Cascos...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(55)
    Call CargarCascos
    frmCargando.Status.Caption = "Cargando Fx's...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(60)
    Call CargarFxs
    frmCargando.Status.Caption = "Cargando Luces...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(67)
    Call InitColor 'Lorwik> OJO! Colores de luces !!
    
    frmCargando.Status.Caption = "Cargando Fuentes...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(65)
    Call texto.Engine_Init_FontSettings
    Call texto.Engine_Init_FontTextures
    
    Call SurfaceDB.Initialize(DirectD3D8, True, "", 80)
    
    InitTileEngine = True
End Function
Public Sub DirectXInit()
    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim EleccionProcessing As Long
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    
    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = IIf(Opciones.VSynC = 0, D3DSWAPEFFECT_COPY, D3DSWAPEFFECT_COPY_VSYNC)
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = 1024
        .BackBufferHeight = 768
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.renderer.hwnd
    End With
    
    If Opciones.VProcessing = 0 Then
        'Mixto
        EleccionProcessing = D3DCREATE_MIXED_VERTEXPROCESSING
    ElseIf Opciones.VProcessing = 1 Then
        'Software
        EleccionProcessing = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    Else
        'Hardware
        EleccionProcessing = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    End If
    
     Set DirectDevice = DirectD3D.CreateDevice( _
                        D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                        frmMain.renderer.hwnd, _
                        EleccionProcessing, _
                        D3DWindow)

    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
    With DirectDevice
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_TFACTOR
    End With
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectD3D. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
    
    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
End Sub

Public Sub DeinitTileEngine()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Destroys all DX objects
'***************************************************
On Error Resume Next

    Set DirectD3D = Nothing
    
    Set DirectX = Nothing
End Sub

Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal Angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If Angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.Bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = 3.1459 - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, (src.Bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub
Public Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.z = z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Public Sub Device_Textured_Render(ByVal X As Integer, ByVal Y As Integer, ByVal Texture As Direct3DTexture8, ByRef src_rect As RECT, ByRef rgb_list() As Long, Optional Alpha As Boolean = False, Optional AlphaByte As Byte = 255, Optional Angle As Byte = 0)
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim srdesc As D3DSURFACE_DESC
    Static light_value(0 To 3) As Long
    Dim i As Byte
    
    'Lorwik> Esto de las luces hay que mirarlo, asi no me convence.
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light
 
    With dest_rect
        .Bottom = Y + (src_rect.Bottom - src_rect.Top)
        .Left = X
        .Right = X + (src_rect.Right - src_rect.Left)
        .Top = Y
    End With
    
    Dim texwidth As Long, texheight As Long
    Texture.GetLevelDesc 0, srdesc
    texwidth = srdesc.Width
    texheight = srdesc.Height
    
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), texwidth, texheight, Angle
    
    DirectDevice.SetTexture 0, Texture
    
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, 3
        DirectDevice.SetRenderState D3DRS_DESTBLEND, 2
    End If
    
    DirectDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(AlphaByte, 0, 0, 0)
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = X + addx
        nY = Y + addy
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    Call ActualizarMiniMapa
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

'******************************************
'HECHIZOS
'******************************************
Public Sub DrawSpells()
    Static re As RECT
    re.Left = 0
    re.Top = 0
    re.Bottom = frmMain.picSpell.ScaleHeight
    re.Right = frmMain.picSpell.ScaleWidth
    
    With DirectDevice
        .Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
        .BeginScene
        Spells.DrawSpells
        .EndScene
        .Present re, ByVal 0, frmMain.picSpell.hwnd, ByVal 0
    End With
End Sub

Public Function Map_Item_Grh_In_Current_Area(ByVal grh_index As Long, ByRef x_pos As Integer, ByRef y_pos As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    On Error GoTo ErrorHandler
    
    Dim map_x As Integer
    Dim map_y As Integer
    Dim X As Integer, Y As Integer
    
    Call Char_Pos_Get(UserCharIndex, map_x, map_y)
    
    If Map_In_Bounds(map_x, map_y) Then
        For Y = map_y - MinYBorder + 1 To map_y + MinYBorder - 1
          For X = map_x - MinXBorder + 1 To map_x + MinXBorder - 1
                If Y < 1 Then Y = 1
                If X < 1 Then X = 1
                If MapData(X, Y).ObjGrh.GrhIndex = grh_index Then
                    x_pos = X
                    y_pos = Y
                    Map_Item_Grh_In_Current_Area = True
                    Exit Function
                End If
          Next X
        Next Y
    End If
    
    Exit Function
    
ErrorHandler:
    Map_Item_Grh_In_Current_Area = False
    
End Function

Public Function Char_Pos_Get(ByVal char_index As Integer, ByRef map_x As Integer, ByRef map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        map_x = charlist(char_index).Pos.X
        map_y = charlist(char_index).Pos.Y
        Char_Pos_Get = True
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional ByVal Radio As Integer) As Long
                                        
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Returns the particle_group_index if successful, else 0
'**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, Radio
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, Radio
    End If
    End If
End Function

Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional Radio As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, Radio
    End If

End Function
 
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function
 
Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_index As Integer
    
    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If

End Function
 
Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) And Not charlist(char_index).particle_count = 0 Then
        For i = 1 To UBound(charlist(char_index).particle_group)
            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i
        Erase charlist(char_index).particle_group
        charlist(char_index).particle_count = 0
    End If
    
End Function
 
Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
 
Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************
On Error GoTo ErrorHandler:
Dim i As Integer

For i = 1 To charlist(char_index).particle_count
    If particle_group_list(charlist(char_index).particle_group(i)).stream_type = stream_type Then
        Char_Particle_Group_Find = charlist(char_index).particle_group(i)
        Exit Function
    End If
Next i

Char_Particle_Group_Find = -1
ErrorHandler:
Debug.Print "Char_Particle_Group_Find Error"
End Function
Public Function Particle_Get_Type(ByVal particle_group_index As Long) As Byte
On Error GoTo ErrorHandler:
    Particle_Get_Type = particle_group_list(particle_group_index).stream_type
Exit Function
ErrorHandler:
    Particle_Get_Type = 0
End Function
Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next
    Dim temp As particle_group
    Dim i As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            For i = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
                If charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
                    Exit For
                End If
            Next i
        End If
    End If
    
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).Active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        Debug.Print particle_group_last & "," & UBound(particle_group_list)
        ReDim Preserve particle_group_list(1 To particle_group_last) As particle_group
    End If
    particle_group_count = particle_group_count - 1
End Sub

 
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional Radio As Integer)
                               
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).Active = True
   
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
   
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    particle_group_list(particle_group_index).Radio = Radio
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = 1
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).Angle = Angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
   'Color > el R y el B esta intercambiados.
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(3)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(1)
   
    'handle
    particle_group_list(particle_group_index).id = id
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index
    End If
   
End Sub

Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional Radio As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).Active = True
    
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    particle_group_list(particle_group_index).Radio = Radio
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).Angle = Angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    'handle
    particle_group_list(particle_group_index).id = id
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on char
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
   
End Sub

Public Function Particle_Type_Get(ByVal particle_Index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_Index) Then
        Particle_Type_Get = particle_group_list(particle_Index).stream_type
    Else
        Particle_Type_Get = 0
    End If
End Function
Public Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Long, ByVal screen_y As Long)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    If particle_group_index = 0 Then Exit Sub
    
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    'Set colors
    If UserMinHP = 0 And frmRenderConnect.Visible = False Then
        temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    Else
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    End If
    
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
                
        'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).Angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
                            particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).Radio, _
                            particle_group_list(particle_group_index).particle_count, loopc
        Next loopc
        
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    
    Else
        'If it's dead destroy it
        Particle_Group_Destroy particle_group_index
    End If
End Sub
 
Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Long, ByVal screen_y As Long, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, _
                            Optional ByVal Radio As Integer, Optional ByVal count As Integer, Optional ByVal Index As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.grh, grh_index, alpha_blend
            If Radio = 0 Then
                temp_particle.X = RandomNumber(x1, x2)
                temp_particle.Y = RandomNumber(y1, y2)
            Else
                temp_particle.X = (RandomNumber(x1, x2) + Radio) + Radio * Cos(PI * 2 * Index / count)
                temp_particle.Y = (RandomNumber(y1, y2) + Radio) + Radio * Sin(PI * 2 * Index / count)
            End If
            temp_particle.X = RandomNumber(x1, x2) - (base_tile_size \ 2)
            temp_particle.Y = RandomNumber(y1, y2) - (base_tile_size \ 2)
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.Angle = Angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.grh.Angle = temp_particle.grh.Angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.Angle >= 360 Then
                temp_particle.Angle = 0
            End If
            
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
        End If
        
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
'Draw it
    If temp_particle.grh.GrhIndex Then
        Call DDrawTransGrhIndextoSurface(grh_index, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, rgb_list(), alpha_blend, , temp_particle.grh.Angle)
    End If
End Sub

Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1
    Do Until particle_group_list(loopc).Active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:

End Function
 
Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If charlist(char_index).particle_count = 0 Then
        Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
        charlist(char_index).particle_count = Char_Particle_Group_Next_Open
        ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
        Exit Function
    End If
    
    loopc = 1
    Do Until charlist(char_index).particle_group(loopc) = 0
        If loopc = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function
 
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).Active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Public Function Map_Particle_Group_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Particle_Group_Get = map_current.map_grid(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

Public Function Map_In_Bounds(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If map_x < map_current.map_x_min Or map_x > map_current.map_x_max Or map_y < map_current.map_y_min Or map_y > map_current.map_y_max Then
        Map_In_Bounds = False
        Exit Function
    End If
   
    Map_In_Bounds = True
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''/[FIN PARTICULAS]'''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

