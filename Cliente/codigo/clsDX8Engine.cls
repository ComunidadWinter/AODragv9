VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MOTOR GRÁFICO ESCRITO(mayormente) POR MENDUZ@NOICODER.COM
Option Explicit

'*********************************
'Particulas
'*********************************
Private base_tile_size As Integer

Private Type Particle
    friction As Single
    x As Single
    y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    X1 As Long
    X2 As Long
    Y1 As Long
    Y2 As Long
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
    ID As Long
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
    angle As Long
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
    ID As Long
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
    
    X1 As Long
    X2 As Long
    Y1 As Long
    Y2 As Long
    angle As Long
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

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Dim Texture As Direct3DTexture8

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public engineBaseSpeed As Single
Public FPS As Integer
Private FramesPerSecCounter As Integer
Private timerElapsedTime As Single
Private timerTicksPerFrame As Double
Private lFrameTimer As Long
Private lFrameLimiter As Long
Private ScrollPixelsPerFrameX As Byte
Private ScrollPixelsPerFrameY As Byte

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

Dim bump_map_supported As Boolean

Private Type decoration
    Grh As Grh
    Render_On_Top As Boolean
    subtile_pos As Byte
End Type

Private Type Map_Tile
    Grh(1 To 3) As Grh
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
   
    Trigger As Byte
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

Public Function Engine_Init() As Boolean
'*****************************************************
'Inicia el motor grafico
'*****************************************************
On Local Error GoTo ErrorHandler

    Dim aD3dai As D3DADAPTER_IDENTIFIER8
    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
    Dim EleccionProcessing As Long
    
    Set SurfaceDB = New clsTexManager
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    Call D3D.GetAdapterIdentifier(D3DADAPTER_DEFAULT, 0, aD3dai)
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
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
    
    Set D3DDevice = D3D.CreateDevice( _
                        D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                        frmMain.renderer.hwnd, _
                        EleccionProcessing, _
                        D3DWindow)
                                                            
    D3DDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    With D3DDevice
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_TFACTOR
    End With
                                                            
    '************************************************************************************************************************************
    
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
    TileBufferSize = 11
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    TilePixelWidth = 32 'Tamaño de tile
    TilePixelHeight = 32 'Tamaño de tile
    engineBaseSpeed = 0.019  'Velocidad a la que va a correr el engine (modifica la velocidad de caminata)
    
    '***********************************
    'Tamaño del mapa
    '***********************************
    MinXBorder = XMinMapSize + (Round(frmMain.renderer.Width / 32) \ 2)
    MaxXBorder = XMaxMapSize - (Round(frmMain.renderer.Width / 32) \ 2)
    MinYBorder = YMinMapSize + (Round(frmMain.renderer.Height / 32) \ 2)
    MaxYBorder = YMaxMapSize - (Round(frmMain.renderer.Height / 32) \ 2)
    '***********************************
    
    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    UserPos.x = 50
    UserPos.y = 50
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectD3D. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
    If D3DDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
    Call InitColor 'Lorwik> OJO! Colores de luces !!
    
    '***********************
    'Cargamos graficos
    '***********************
    'Lorwik> Si alguien tiene un error en esta parte, lo puedo saber mejor de este modo.
    
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
    frmCargando.Status.Caption = "Cargando Fuentes...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(65)
    texto.Engine_Init_FontSettings
    texto.Engine_Init_FontTextures
    movSpeed = 1
    
    Engine_Init = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error Number Returned: " & Err.number
    MsgBox "Error in engine initialization: " & Err.number & ": " & Err.Description & " Dispositivo: " & Trim$(StrConv(aD3dai.Description, vbUnicode)), vbCritical, "Direct3D Initialization"
    Engine_Init = False
End Function

Public Sub Engine_Deinit()
'************************************
'Termina con el motor grafico
'************************************
    Erase MapData
    Erase charlist
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing
    End
End Sub

Private Function CreateTLVertex(x As Single, y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    CreateTLVertex.x = x
    CreateTLVertex.y = y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

Public Sub Engine_ActFPS()
    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
End Sub

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Public Sub Draw_GrhIndex(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer, ByRef Light() As Long)
'********************************************
'Dibuja desde un GrhIndex que le indiquemos
'********************************************
    If grh_index <= 0 Then Exit Sub
    
    Device_Box_Textured_Render grh_index, _
        x, y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        Light, _
        GrhData(grh_index).SX, GrhData(grh_index).SY
End Sub

Private Sub Draw_Grh(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef Light() As Long, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional Transp As Byte = 255)
'**********************************************************
'Dibuja desde un Grh, pero primero hay que iniciar el Grh
'**********************************************************
    Dim CurrentGrhIndex As Long
    If Grh.GrhIndex = 0 Then Exit Sub
    
    If Grh.GrhIndex > GrhCount Or GrhData(Grh.GrhIndex).NumFrames = 0 And GrhData(Grh.GrhIndex).FileNum = 0 Then
        Call InitGrh(Grh, 20299)
    End If
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * (movSpeed / 1.5)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If

    Device_Box_Textured_Render CurrentGrhIndex, _
        x, y, _
        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
        Light, _
        GrhData(CurrentGrhIndex).SX, GrhData(CurrentGrhIndex).SY, _
        Alpha _
        , angle, Transp
'exits:

End Sub

Public Sub Start()
'**********************************
'Lorwik
'Bucle principal
'**********************************

    Do While prgRun
        If EngineRun Then
            If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
                Engine.Render
                If Not pausa Then Call CheckKeys
                Call FlushBuffer
            End If
            
            If frmRenderConnect.Visible = True Then Engine.RenderConnect 'Para el conectar
            If (Opciones.Audio = 1 Or Opciones.sMusica <> CONST_DESHABILITADA) Then Call Sound.Sound_Render
        End If
        
        'Call FlushBuffer
0        DoEvents
    Loop
End Sub

Public Sub Render()
On Error Resume Next
'******************************************************************************************************
'Dibujamos en el render
'De aqui salen los textos, el RenderScreen, el CharRender y todo lo demas que se dibuje en el render
'*******************************************************************************************************

    Static re As RECT
        re.Left = 0
        re.Top = 0
        re.Bottom = frmMain.renderer.ScaleHeight
        re.Right = frmMain.renderer.ScaleWidth

    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    ShowNextFrame
    
    '*********************FPS**********************************
    Engine.Engine_ActFPS
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
  
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, 0, ByVal 0
    
    lFrameLimiter = GetTickCount
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
End Sub

Sub RenderConnect()
'**************************************************************************************
'Dibujamos un mapa en el frmConnect, desde aqui podemos controlar que queremos dibujar
'**************************************************************************************

Dim x As Byte, y As Byte

    Static re As RECT
        re.Left = 0
        re.Top = 0
        re.Bottom = 768
        re.Right = 1024

        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
        D3DDevice.BeginScene
        
        'Capa 1
        For x = 1 To 32
            For y = 1 To 24
                With MapData(x + MapaConnect.x, y + MapaConnect.y)
                    Call Draw_Grh(.Graphic(1), _
                        (x - 1) * 32, (y - 1) * 32, _
                         0, 1, MapData(x, y).light_value)
                 End With
            Next y
        Next x
        
        'Capa 2
        For x = 1 To 32
            For y = 1 To 24
                With MapData(x + MapaConnect.x, y + MapaConnect.y)
                    If .Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(.Graphic(2), _
                                (x - 1) * 32, (y - 1) * 32, _
                                1, 1, MapData(x, y).light_value)
                  End If
                End With
            Next y
        Next x
        
    
        'Capa 3
        For x = 1 To 32
            For y = 1 To 24
                With MapData(x + MapaConnect.x, y + MapaConnect.y)
                    If .Graphic(3).GrhIndex <> 0 Then
                        Call Draw_Grh(.Graphic(3), _
                                (x - 1) * 32, (y - 1) * 32, _
                                1, 1, MapData(x, y).light_value)
                    End If
                End With
            Next y
        Next x
        
        'Capa 4
        For x = 1 To 32
            For y = 1 To 24
                With MapData(x + MapaConnect.x, y + MapaConnect.y)
                    If .Graphic(4).GrhIndex <> 0 Then
                        Call Draw_Grh(.Graphic(4), _
                                (x - 1) * 32, (y - 1) * 32, _
                                1, 1, MapData(x, y).light_value)
                    End If
                End With
            Next y
        Next x
                
        texto.Engine_Text_Draw 5, 750, "Version: " & App.Major & "." & App.Minor & "." & App.Revision, vbCyan
        
        If frmCuenta.Visible = True Then
            texto.Engine_Text_Draw 505, 600, ConsejoSeleccionado, vbWhite, , 1
            texto.Engine_Text_Draw 505, 630, "Siguiente consejo", &HC0FFFF, , 1
        End If
        
        'Clima**************************************
        Call ClimaX
        '*******************************************
    
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
                               
        D3DDevice.EndScene
        D3DDevice.Present re, ByVal 0, frmRenderConnect.hwnd, ByVal 0
End Sub

Sub RenderPJCuenta()
On Error GoTo errhandler
    Static re As RECT
        re.Left = 0
        re.Top = 0
        re.Bottom = 94
        re.Right = 59
        
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
        D3DDevice.BeginScene
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
        
        If Not frmCuenta.ListPJ.ListIndex < 0 Then 'Con este If evitamos que tire error en caso de que no tengamos ningun pj seleccionado
            With Cuenta.pjs(frmCuenta.ListPJ.ListIndex + 1)
                Call Draw_Grh(HeadData(.rcvHead).Head(3), 15, 45 + BodyData(.Acuerpo).HeadOffset.y, 1, 1, LightIluminado)
                Call Draw_Grh(BodyData(.Acuerpo).Walk(3), 15, 45, 1, 1, LightIluminado)
            End With
        End If
        D3DDevice.EndScene
        D3DDevice.Present re, ByVal 0, frmCuenta.RenderPJPic.hwnd, ByVal 0
errhandler:
    Debug.Print "Error en RenderPJCuenta."
End Sub

Sub ShowNextFrame()
'**********************************************************************
'De aqui pasamos a dibujar todo aquello que se va a mover con el PJ
'(PJ, Bichos, Decorados, etc...
'**********************************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.x <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        If Not UserCiego Then Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
        
End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Dibuja todo el decorado, PJ, Bichos y lo que le metamos...
'**************************************************************
    Dim y                   As Integer     'Keeps track of where on map we are
    Dim x                   As Integer     'Keeps track of where on map we are
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
    If screenmaxX > XMaxMapSize Then screenmaxX = 100
    If screenmaxY > YMaxMapSize Then screenmaxY = 100
    If screenminX < XMinMapSize Then screenmaxX = 1
    If screenminY < YMinMapSize Then screenmaxY = 1
    
    '=============================================================================================================
    'Comenzamos a dibujar las capas. Aquellas capas que esten por debajo de las demas, se van a mostrar arriba.
    'Ejemplo: La capa 2 se mostrara por debajo de la 3 y la 4 se mostrara por encima de la 3.
    '=============================================================================================================
    
    'Draw floor layer
    
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            'Capa 1 **********************************
            
                Call Draw_Grh(MapData(x, y).Graphic(1), _
                        (ScreenX - 1) * 32 + PixelOffsetX, _
                        (ScreenY - 1) * 32 + PixelOffsetY, _
                        0, 1, MapData(x, y).light_value, , x, y)
            '******************************************
            
            'Capa 2 ************************************
                If MapData(x, y).Graphic(2).GrhIndex <> 0 Then _
                    Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, MapData(x, y).light_value, , x, y)
            '*******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y

    
    ScreenY = minYOffset - TileBufferSize
    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(x, y)
                '******************************************

                '**************************************************
                'Capa de objetos
                'Dibujara solo los objetos
                '**************************************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, , x, y)
                End If
                
                '***************************************************
                'Capa de Char
                'Dibuja todo aquello que tenga cuerpo (PJ y Bichos)
                '***************************************************
                If .CharIndex <> 0 Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, x, y, MapData(x, y).light_value)
                End If
                '*************************************************
                
                'Dibujamos el valor en el render.
                If .RenderValue.Activated Then
                    modRenderValue.Draw x, y, PixelOffsetXTemp + 20, PixelOffsetYTemp - 30
                End If
                
                'Capa 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, , x, y)
                End If
                '************************************************

            End With
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    ScreenY = minYOffset - 5
    
    'Particulas ************************************************
    ScreenY = minYOffset - TileBufferSize
        For y = minY To maxY
            ScreenX = minXOffset - TileBufferSize
                For x = minX To maxX
                    'Particulas**************************************
                        If MapData(x, y).particle_group_index Then _
                            Particle_Group_Render MapData(x, y).particle_group_index, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY
                    '************************************************
                ScreenX = ScreenX + 1
            Next x
        ScreenY = ScreenY + 1
    Next y
    '***********************************************************
                         
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
    '===========================================================
    'La capa 4 es utilizada mayormente para los techos
    'si el usuario esta en cierto Trigger, esta capa
    'desaparecera hasta que el usuario salga de ese trigger
    '===========================================================

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
    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
            If Not MapData(x, y).Graphic(4).GrhIndex Then
                'Draw
                Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, 1, 1, TechoColor(), , , , , bTechoAB)
            End If
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    '*******************************************
        
    'Clima**************************************
    Call ClimaX
    '*******************************************
 
End Sub

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte, Light() As Long)
'*******************************************************
'Esto forma parte del RenderScreen.
'Dibuja todo aquello que tenga cuerpo (por asi decirlo)
'Bichos y PJ
'*******************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    Dim TempName As String
    Dim TempClanName As String
    Dim PartyConfirmed As Boolean, ClanConfirmed As Boolean
    Dim PixelOffsetYTemp As Integer
    
    
    With charlist(CharIndex)
    
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Inicia la animacion
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Movemos el Char
                moved = True
                
                'Comprueba si ya llegamos
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Inicia la animacion
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!! <<<-- Lorwik> Esto no lo entiendo
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Movemos el Char
                moved = True
                
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
            'Paramos la animacion
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            If IsAttacking = False Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
                
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            End If
            
            If .NPCAttack = False Then
                .Ataque.AtaqueWalk(.Heading).Started = 0
                .Ataque.AtaqueWalk(.Heading).FrameCounter = 1
            End If
            
            .Moving = False
        Else
            IsAttacking = False
            .NPCAttack = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        
        If .invisible = False Then
            'Dibujamos el nombre debajo de la cabeza.
            If Opciones.NamePlayers Then
                If Len(.Nombre) > 0 Then
                    Pos = InStr(.Nombre, "<")
                    If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                    If .priv = 0 Then
                        Select Case .Criminal
                            Case 0
                                color = ColoresPJ(49)
                            Case 1
                                color = ColoresPJ(50)
                            Case 8
                                color = ColoresPJ(4)
                            Case 9
                                color = ColoresPJ(9)
                            Case 10
                                color = ColoresPJ(3)
                        End Select
                    Else
                        color = ColoresPJ(.priv)
                    End If
                            
                    'Nick
                    line = .Nombre
                    Call texto.Engine_Text_Draw(PixelOffsetX + 15, PixelOffsetY + 30, line, color, , True)
                            
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
        
        '¿Tiene cabeza?
        If .Head.Head(.Heading).GrhIndex Then
            movSpeed = 1.3
            'Dibujamos el cuerpo
            If .Body.Walk(.Heading).GrhIndex Then _
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , x, y, , .Estainvi)
            'Dibujamos la Cabeza
            If .Head.Head(.Heading).GrhIndex Then
                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, Light, , x, y, , .Estainvi)
                'Dibujamos el Casco
                If .Casco.Head(.Heading).GrhIndex Then _
                    Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, Light, , x, y, , .Estainvi)
                             
                If UserMontando = True Then
                    PixelOffsetYTemp = PixelOffsetY + .Body.HeadOffset.y + 35
                Else
                    PixelOffsetYTemp = PixelOffsetY + .Body.HeadOffset.y + 35
                End If
                          
                'Dibujamos el arma
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                    Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetYTemp, 1, 1, Light, , x, y, , .Estainvi)
                'Dibujamos el escudo
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                    Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetYTemp, 1, 1, Light, , x, y, , .Estainvi)
            
            End If
        Else
            If .NPCAttack = True And .Ataque.AtaqueWalk(.Heading).GrhIndex > 0 Then
                If .Ataque.AtaqueWalk(.Heading).GrhIndex Then _
                    Call Draw_Grh(.Ataque.AtaqueWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , x, y, , .Estainvi)
            Else
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , x, y, , .Estainvi)
            End If
        End If
        
        'Actualizamos los dialogos
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, CharIndex)
        movSpeed = 1
        
        '************Particulas************
        Dim i As Integer
            If .particle_count > 0 Then
                For i = 1 To .particle_count
                    If .particle_group(i) > 0 Then _
                        Particle_Group_Render .particle_group(i), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY
                Next i
            End If
        
        '*******Dibujamos los FX********
        If .FxIndex <> 0 Then

            Call Draw_Grh(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1, Light, FxData(.FxIndex).FXTransparente)
            
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
        
        '14/11/2018 Pinto el icono encima del NPC segun el estado en el que se encuentre el jugador con las quests
        If .EstadoQuest > 2 Then
          Select Case .EstadoQuest
            Case eEstadoQuest.NoAceptada
              Call Engine.Draw_GrhIndex(2176, PixelOffsetX + 7, PixelOffsetY - 37, LightIluminado())
              
            Case eEstadoQuest.CompletadaFaltaEntregar
              Call Engine.Draw_GrhIndex(2177, PixelOffsetX + 7, PixelOffsetY - 37, LightIluminado())
              
          End Select
        End If
        
    End With
End Sub


Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function


Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)
'**************************************************************
'Crea el plano donde se desarrolla el juego y todo se dibuja.
'Si jugamos con la configuracion de este sub, podremos provocar
'ciertos efecto. (ejem: montañas, reflejos y sombras)
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
   
    If angle > 0 Then
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
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
   
   
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, (src.Bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
   
   
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
   
   
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
   
   
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If
 
 
End Sub

Public Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, Optional alphabyte As Byte = 255)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    'Lorwik> Esto de las luces hay que mirarlo, asi no me convence.
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light
        
    'Set up the source rectangle
    With src_rect
        .Bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(alphabyte, 0, 0, 0)
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
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

Public Sub Char_Move_by_Head(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Movemos el Char en una direccion
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim x As Integer
    Dim y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
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
        
        nX = x + addx
        nY = y + addy
        
        'Para que no explote esto
        If nX <= 0 Then nX = 1
        If nY <= 0 Then nY = 1
        If nY >= 100 Then nY = 99
        If nY >= 100 Then nY = 99
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.x = nX
        .Pos.y = nY
        MapData(x, y).CharIndex = 0
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado <> 1 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Public Sub Char_Move_by_Pos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
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
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XXGRANDECIU Or .FxIndex = FxMeditar.XXGRANDECRI Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then
    'Call Dialogos.RemoveDialog(CharIndex)
    'Call Hits.RemoveHit(CharIndex)
    End If
    'If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    '    Call EraseChar(CharIndex)
    'End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder
    End With
End Function

Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
Dim r As Integer, g As Integer, b As Integer, A As Integer
    r = 127 * Vec.x + 128
    g = 127 * Vec.y + 128
    b = 127 * Vec.Z + 128
    A = 255 * fHeight
    VectorToRGBA = D3DColorARGB(A, r, g, b)
End Function

'******************************************
'HECHIZOS
'******************************************
Public Sub DrawSpells()
    Static re As RECT
    re.Left = 0
    re.Top = 0
    re.Bottom = frmMain.picSpell.ScaleHeight
    re.Right = frmMain.picSpell.ScaleWidth
    
    With D3DDevice
        .Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
        .BeginScene
        Spells.DrawSpells
        .EndScene
        .Present re, ByVal 0, frmMain.picSpell.hwnd, ByVal 0
    End With
End Sub
'******************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                        Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
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
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, Radio
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, Radio
    End If
    End If
End Function

Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                        Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
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
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, Radio
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
 
Public Function Particle_Group_Find(ByVal ID As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).ID = ID
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
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
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
   
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
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
    particle_group_list(particle_group_index).ID = ID
   
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
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
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
   
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
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
    particle_group_list(particle_group_index).ID = ID
   
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
                            particle_group_list(particle_group_index).X1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).X2, _
                            particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
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
                            Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
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
            InitGrh temp_particle.Grh, grh_index, alpha_blend
            If Radio = 0 Then
                temp_particle.x = RandomNumber(X1, X2)
                temp_particle.y = RandomNumber(Y1, Y2)
            Else
                temp_particle.x = (RandomNumber(X1, X2) + Radio) + Radio * Cos(PI * 2 * Index / count)
                temp_particle.y = (RandomNumber(Y1, Y2) + Radio) + Radio * Sin(PI * 2 * Index / count)
            End If
            temp_particle.x = RandomNumber(X1, X2) - (base_tile_size \ 2)
            temp_particle.y = RandomNumber(Y1, Y2) - (base_tile_size \ 2)
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
            
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
        End If
        
        'Add in vector
        temp_particle.x = temp_particle.x + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.y = temp_particle.y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
'Draw it
    If temp_particle.Grh.GrhIndex Then
        Draw_Grh temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, 1, 1, rgb_list(), alpha_blend, , , temp_particle.Grh.angle
    End If
    
    If temp_particle.Grh.GrhIndex Then
        Draw_Grh temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, 1, 1, rgb_list(), alpha_blend, , , temp_particle.Grh.angle
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

Private Function Char_Check(ByVal char_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)
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
''''''''''''''''''''''''/[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Map_Item_Grh_In_Current_Area(ByVal grh_index As Long, ByRef x_pos As Integer, ByRef y_pos As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    On Error GoTo ErrorHandler
    
    Dim map_x As Integer
    Dim map_y As Integer
    Dim x As Integer, y As Integer
    
    Call Char_Pos_Get(UserCharIndex, map_x, map_y)
    
    If Map_In_Bounds(map_x, map_y) Then
        For y = map_y - MinYBorder + 1 To map_y + MinYBorder - 1
          For x = map_x - MinXBorder + 1 To map_x + MinXBorder - 1
                If y < 1 Then y = 1
                If x < 1 Then x = 1
                If MapData(x, y).ObjGrh.GrhIndex = grh_index Then
                    x_pos = x
                    y_pos = y
                    Map_Item_Grh_In_Current_Area = True
                    Exit Function
                End If
          Next x
        Next y
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
        map_x = charlist(char_index).Pos.x
        map_y = charlist(char_index).Pos.y
        Char_Pos_Get = True
    End If
End Function
