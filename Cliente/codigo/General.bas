Attribute VB_Name = "Mod_General"
Option Explicit

'*****MOD Application*******
Private Declare Function GetActiveWindow Lib "user32" () As Long
'***************************

'***************************************
'Para obtener bytes libres en la unidad
'***************************************
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, bytesTotal As Currency, FreeBytesTotal As Currency) As Long

'***************************************
'Para obetener memoria libre en la RAM
'***************************************
Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'**************************************

'*************************************
'Lorwik - Nuevo formato de Mapas .CSM
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
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    OBJIndex As Integer
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
    RoboNPcsPermitido As Boolean
    ResuSinEfecto As Boolean
    MagiaSinEfecto As Boolean
    InviSinEfecto As Boolean
    NoEncriptarMP As Boolean
    version As Long
End Type

Private MapSize As tMapSize
Public MapDat As tMapDat
'********************************

'***************************
'MOVER VENTANAS
'***************************
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'***************************

Public bFogata As Boolean

Private lFrameTimer As Long
Public MinEleccion As Integer, MaxEleccion As Integer
Public Actual As Integer

'***********************
'Forms transparentes
'***********************
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

'*********************
'OBTENER SERIAL HD
'*********************
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'********************
'CONOCER S.O
'********************
'To get OS version
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      ' Maintenance string for PSS usage
End Type

Private Const VER_PLATFORM_WIN32s As Long = 0&
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
Private Const VER_PLATFORM_WIN32_NT As Long = 2&

Private Declare Function GetOSVersion Lib "kernel32" _
Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private OSInfo As OSVERSIONINFO

Sub DameOpciones()
 
Dim i As Integer
 
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = 1
                MinEleccion = 1
                MaxEleccion = 43
            Case "Elfo"
                Actual = 101
                MinEleccion = 101
                MaxEleccion = 132
            Case "Elfo Oscuro"
                Actual = 201
                MinEleccion = 201
                MaxEleccion = 230
            Case "Enano"
                Actual = 301
                MinEleccion = 301
                MaxEleccion = 330
            Case "Gnomo"
                Actual = 401
                MinEleccion = 401
                MaxEleccion = 430
            Case "Orco"
                Actual = 501
                MinEleccion = 501
                MaxEleccion = 530
            Case "No-Muerto"
                Actual = 507
                MinEleccion = 625
                MaxEleccion = 626
            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = 70
                MinEleccion = 70
                MaxEleccion = 100
            Case "Elfo"
                Actual = 170
                MinEleccion = 170
                MaxEleccion = 200
            Case "Elfo Oscuro"
                Actual = 270
                MinEleccion = 270
                MaxEleccion = 300
            Case "Gnomo"
                Actual = 399
                MinEleccion = 399
                MaxEleccion = 470
            Case "Enano"
                Actual = 370
                MinEleccion = 370
                MaxEleccion = 399
            Case "Orco"
                Actual = 570
                MinEleccion = 570
                MaxEleccion = 599
            Case "No-Muerto"
                Actual = 560
                MinEleccion = 650
                MaxEleccion = 651
            Case Else
                Actual = 70
                MaxEleccion = 70
                MinEleccion = 70
        End Select
End Select
End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
    EnableUrlDetect
    With RichTextBox
        If Len(.Text) > 10000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf)
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
            .Text = ""
        End If
       
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
       
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
       
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
       
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).Active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean, Optional ByVal acc As Boolean = False) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If acc = True Then
        UserPassword = Cuenta.Pass
        UserName = Cuenta.name
    End If
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

#If SeguridadAlkon Then
    Call UnprotectForm
#End If

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmCuenta
    Unload frmCrearPersonaje
    Unload frmConnect
    Unload frmRenderConnect
    
    frmMain.lblName.Caption = UserName
    frmMain.lblName.Refresh
    'Load main form
    frmMain.Visible = True

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    'Bloqueo de movimiento en escritura
    If frmMain.SendTxt.Visible = True And Opciones.MovEscritura = 1 Then Exit Sub
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    charlist(UserCharIndex).moved = True
    
    If LegalOk And Not UserParalizado Then
        If Not UserDescansar And Not UserMeditar Then
            Call WriteWalk(Direccion)
            MoveScreen Direccion
            MoveCharbyHead UserCharIndex, Direccion
            Call ActualizarMiniMapa
        Else
            If UserDescansar And Not UserAvisado Then
                UserAvisado = True
                Call WriteRest
            End If
            If UserMeditar And Not UserAvisado Then
                UserAvisado = True
                Call WriteMeditate
            End If
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    'Lorwik> Cambio de Zonas
    If UserPos.Y = 101 Or UserPos.Y = 99 Then Call DibujarMiniMapa
    If UserPos.X = 101 Or UserPos.X = 99 Then Call DibujarMiniMapa
    
    Call FlushBuffer
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(SOUTH, EAST))
End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************

    'No input allowed while Argentum is not the active window
    If Not Mod_General.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If input_key_get(CustomKeys.BindedKey(eKeyType.mKeyUp)) Then
                Call MoveTo(NORTH)
            'Move Right
            ElseIf input_key_get(CustomKeys.BindedKey(eKeyType.mKeyRight)) Then
                Call MoveTo(EAST)
            'Move down
            ElseIf input_key_get(CustomKeys.BindedKey(eKeyType.mKeyDown)) Then
                Call MoveTo(SOUTH)
            'Move left
            ElseIf input_key_get(CustomKeys.BindedKey(eKeyType.mKeyLeft)) Then
                Call MoveTo(WEST)
            End If
            
        Else
            If input_key_get(CustomKeys.BindedKey(eKeyType.mKeyUp)) Or input_key_get(CustomKeys.BindedKey(eKeyType.mKeyRight)) _
            Or input_key_get(CustomKeys.BindedKey(eKeyType.mKeyDown)) Or input_key_get(CustomKeys.BindedKey(eKeyType.mKeyLeft)) Then
                Call RandomMove 'Si presiona cualquier tecla y es estupido se mueve para cualquier lado.
            End If
        End If

        frmMain.lblMapCoord.Caption = UserMap & "," & UserPos.X & "," & UserPos.Y
    End If
End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer, ByVal Dir_Map As String)
'**************************************************************
'Sistema de mapas adaptado de IAO 1.4 y mejorado por Lorwik
'Este sistema comprueba y carga solo la informacion necesaria sin hacer consultas inutiles
'**************************************************************
    'Limpiamos Cuerpos y Objetos antes de cargar el mapa.
    Call Char_Clean
    
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
    
    Dim i As Long
    Dim j As Long
        fh = FreeFile
        'Abrimos el mapa
        Open Dir_Map For Binary Access Read As fh
        Get #fh, , MH
        Get #fh, , MapSize
        Get #fh, , MapDat
        
        ReDim MapData(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As MapBlock
        ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
        
        Get #fh, , L1
        
        With MH
            'Cargamos los tiles que tengan bloqueos
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs
                For i = 1 To .NumeroBloqueados
                    MapData(Blqs(i).X, Blqs(i).Y).blocked = 1
                Next i
            End If
            
            'Cargamos los tiles que tengan Capa 2
            If .NumeroLayers(2) > 0 Then
                ReDim L2(1 To .NumeroLayers(2))
                Get #fh, , L2
                For i = 1 To .NumeroLayers(2)
                    InitGrh MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex
                Next i
            End If
            
            'Cargamos los tiles que tengan Capa 3
            If .NumeroLayers(3) > 0 Then
                ReDim L3(1 To .NumeroLayers(3))
                Get #fh, , L3
                For i = 1 To .NumeroLayers(3)
                    InitGrh MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex
                Next i
            End If
            
            'Cargamos los tiles que tengan Capa 4
            If .NumeroLayers(4) > 0 Then
                ReDim L4(1 To .NumeroLayers(4))
                Get #fh, , L4
                For i = 1 To .NumeroLayers(4)
                    InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex
                Next i
            End If
            
            'Cargamos los tiles que tengan Triggers
            If .NumeroTriggers > 0 Then
                ReDim Triggers(1 To .NumeroTriggers)
                Get #fh, , Triggers
                For i = 1 To .NumeroTriggers
                    MapData(Triggers(i).X, Triggers(i).Y).trigger = Triggers(i).trigger
                Next i
            End If
            
            'Cargamos los tiles que tengan particulas
            'NOTA: Desactivado hasta implementar el sistema
            If .NumeroParticulas > 0 Then
                ReDim Particulas(1 To .NumeroParticulas)
                Get #fh, , Particulas
                For i = 1 To .NumeroParticulas
                    MapData(Particulas(i).X, Particulas(i).Y).particle_group_index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
                Next i
            End If
            
            'Cargamos los tiles que tengan luces
            'NOTA: Desactivado hasta implementar el sistema
            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
                Dim p As Byte
                Get #fh, , Luces
                For i = 1 To .NumeroLuces
                    For p = 0 To 3
                       'MapData(Luces(i).x, Luces(i).y).base_light(p) = Luces(i).base_light(p)
                        'If MapData(Luces(i).x, Luces(i).y).base_light(p) Then _
                            MapData(Luces(i).x, Luces(i).y).light_value(p) = Luces(i).light_value(p)
                    Next p
                Next i
            End If
            
        End With
    
    Close fh
    
    'Cargamos el 100 x 100 de los tiles ya que la Capa 1 tiene que estar presente en todo el mapa.
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If L1(i, j) > 0 Then
                InitGrh MapData(i, j).Graphic(1), L1(i, j)
            End If
        Next i
    Next j
    
    '******************************************
    'EXTRAS
    '******************************************
    MapInfo.name = MapDat.map_name
    MapInfo.Music = MapDat.music_number
    MapInfo.Ambient = MapDat.Ambient
    TransMapAB = 255
    
    CurMap = Map
    
    Delete_File Dir_Map
End Sub

Public Sub Char_Clean()
Dim X As Byte
Dim Y As Byte
For X = 1 To 100
    For Y = 1 To 100
        If MapData(X, Y).CharIndex Then
            EraseChar MapData(X, Y).CharIndex
        End If
        If MapData(X, Y).ObjGrh.GrhIndex Then
            MapData(X, Y).ObjGrh.GrhIndex = 0
        End If
    Next Y
Next X
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub Main()
    '**************************************
    'Iniciar Apariencia de Windows In Game!
    InitManifest
    '**************************************
    
    ReDim EfectoList(0) As EfectoInfo
    ReDim ZonaList(0) As ZonaInfo

    frmCargando.Show
    frmCargando.Refresh
    'Establecemos el 0% de la carga
    Call frmCargando.establecerProgreso(0)

#If Desarrollo = 0 Then
    OriginalClientName = "AODrag 9.0"
    ClientName = App.EXEName
    DetectName = App.EXEName
    
    'If ChangeName Then
     '   MsgBox "Se ha detectado un cambio de nombre en el ejecutable. ¡No es posible ejecutar el cliente!.", vbCritical, "AODrag"
    '    End
    'End If

    If GetVar(App.Path & "\config.ini", "general", "launcher") = 0 Then
        Call WriteVar(App.Path & "\config.ini", "general", "launcher", "0")
    Else
        MsgBox "¡Debes de ejecutar el cliente desde el Launcher!", vbInformation
        End
        Exit Sub
    End If

    If FindPreviousInstance Then
        'Call MsgBox("¡AODrag ya está corriendo! No es posible correr otra instancia del juego. Relea el reglamento. Haga click en Aceptar para salir." & vbCrLf & vbCrLf & "AODrag is already running. Game cannot be run. Click OK to quit.", vbApplicationModal + vbInformation + vbOKOnly, "Already running!")
        'End
    End If
#End If
    
    '******************Paquetes*******************************
    frmCargando.Status.Caption = "Buscando Paquetes... "
    If FileExist(App.Path & "\RECURSOS\tmp.DRAG", vbNormal) Then
        Call MsgBox("Hay Actualizaciones para los recursos. El cliente se cerrará y se abrirá el launcher para aplicar la actualización.", vbOKOnly, "Cliente Desactualizado")
        Call Shell(App.Path & "\AODrag Launcher.exe", vbNormalFocus)
        End
    End If
    
    frmCargando.Status.Caption = "Cargando Opciones..."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(2)
    
    Call CargarOpciones
    
    frmCargando.Status.Caption = "Iniciando AntiCheat..."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(5)
    
#If Desarrollo = 0 Then
    If Debugger Then
        Call AntiDebugger
        End
    End If
    
    Call ModSeguridad.BuscarEngine
    Call ModSeguridad.AntiShInitialize
    
    Call CargarNombreCheats
    Call BuscarCheats
#End If
    
    frmCargando.Status.Caption = "Cargando consejos...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(8)
    
    'Call ListarConsejos

    frmCargando.Status.Caption = "Iniciando nombres..."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(10)
    
    Call InicializarNombres
    
    frmCargando.Status.Caption = "Iniciando fuentes.."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(15)
    
    Call CargarMensajes
    
    frmCargando.Status.Caption = "Cargando Mensajes.."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(17)
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
    
    frmCargando.Status.Caption = "Instanciando clases..."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(19)
    
    Set Sound = New clsSoundEngine
    Set Dialogos = New clsDialogs
    Set incomingData = New clsByteQueue
    Set outgoingData = New clsByteQueue
    Set MainTimer = New clsTimer
    Set SurfaceDB = New clsSurfaceManDyn
    
    frmCargando.Status.Caption = "Iniciando motor grafico...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(20)
    
    Call Resolution.SetResolution
    
    Call DirectXInit
    
    'Importante que primero iniciemos el engine
    If Not InitTileEngine(frmMain.renderer.hwnd, 149, 13, 32, 32, 13, 17, 11, 8, 8, 0.019) Then
        MsgBox "¡No se ha logrado iniciar el engine gráfico! Reinstale los últimos controladores de DirectX y actualize sus controladores de video. Si el problema persiste por favor consulte los foros de soporte.", vbCritical, "Saliendo"
        Call CloseClient
    End If
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(App.Path & DirCursores & "d.ico", vbArchive) Then _
        Set picMouseIcon = LoadPicture(App.Path & DirCursores & "d.ico")
    
    UserMap = 1
    
    frmCargando.Status.Caption = "Cargando armas...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(70)
    
    Call CargarAnimArmas
    
    frmCargando.Status.Caption = "Cargando Escudos...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(72)
    
    Call CargarAnimEscudos
    
    frmCargando.Status.Caption = "Cargando Colores...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(74)
    
    Call CargarColores
    
    frmCargando.Status.Caption = "Cargando Pasos...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(70)
    
    Call CargarPasos
    
    frmCargando.Status.Caption = "Iniciando DirectSound...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(80)
    
    If Sound.Initialize_Engine(frmMain.hwnd, App.Path & "\Recursos", App.Path & "\Recursos", App.Path & "\Recursos", False, (Opciones.Audio > 0), (Opciones.sMusica <> CONST_DESHABILITADA), Opciones.FXVolume, False, Opciones.InvertirSonido) Then
        'frmCargando.picLoad.Width = 300
    Else
        MsgBox "¡No se ha logrado iniciar el engine de DirectSound! Reinstale los últimos controladores de DirectX desde www.aodrag.es. No habrá soporte de audio en el juego.", vbCritical, "Advertencia"
        frmOpciones.Frame2.Enabled = False
    End If
    
    frmCargando.Status.Caption = "Cargando Musica de Inicio...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(80)
    
    If Opciones.sMusica <> CONST_DESHABILITADA Then
        Sound.NextMusic = MUS_Inicio
        Sound.Fading = 350
        Sound.Sound_Render
    End If
    
    frmCargando.Status.Caption = "Iniciando inventario...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(85)
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectD3D8, frmMain.picInv, MAX_INVENTORY_SLOTS, , , , , , , True)
    
    frmCargando.Status.Caption = "Iniciando Hechizos...."
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(87)
    
    'Iniciamos los hechizos
    Call Spells.Initialize(frmMain.picSpell)
    
    frmCargando.Status.Caption = "¡Bienvenido a AODrag!"
    frmCargando.Refresh
    Call frmCargando.progresoConDelay(100)
    'Give the user enough time to read the welcome text
    Call Sleep(1)
    
    Unload frmCargando
    Call MostrarConnect
    frmConnect.Show vbModeless, frmRenderConnect
    
    '27/02/2016 Lorwik: Busca y pre-selecciona un server.
    Call ListarServidores
    ServIndSel = 0
    
    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
    Call LoadTimerIntervals
    
    Do While prgRun
        If EngineRun Then
            If frmMain.WindowState <> vbMinimized And frmMain.Visible Then
                Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
                
                If Not pausa Then Call CheckKeys
                ' If there is anything to be sent, we send it
                Call FlushBuffer
            End If
            
            If frmRenderConnect.Visible = True Then RenderConnect 'Para el conectar
            If (Opciones.Audio = 1 Or Opciones.sMusica <> CONST_DESHABILITADA) Then Call Sound.Sound_Render
        End If
        
        'Call FlushBuffer
       DoEvents
    Loop
End Sub

Private Sub LoadTimerIntervals()
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim Lx    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For Lx = 0 To Len(sString) - 1
            If Not (Lx = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (Lx + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next Lx
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
On Error GoTo errhandler
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaRazas(eRaza.orco) = "Orco"
    ListaRazas(eRaza.nomuerto) = "No-Muerto"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"

    SkillsNames(eSkill.magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
    SkillsNames(eSkill.Wrestling) = "Wrestling"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Energia) = "Energia"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
    
errhandler:

End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Function esGM(CharIndex As Integer) As Boolean
esGM = False
If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then _
    esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
End Function

'*******************************
'MOD APPLICATION
'*******************************
Public Function IsAppActive() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (maraxus)
'Last Modify Date: 03/03/2007
'Checks if this is the active application or not
'***************************************************
    IsAppActive = (GetActiveWindow <> 0)
End Function
'******************************
'******************************

'******************************
'Labels multi informacion
'******************************

Public Sub LabelExperiencia()

    If UserPasarNivel <= 0 Then UserPasarNivel = 1
    frmMain.ExpShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 141)
    
    If lblexpactivo = True Then
        frmMain.lblPorcLvl.Caption = UserExp & " / " & UserPasarNivel
    Else
        If Not UserLvl = MAX_LEVEL Then
            frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 0) & "%"
        Else
            frmMain.lblPorcLvl.Caption = "¡Nivel Máximo!"
        End If
   End If
End Sub

'*****************************
'*****************************

'*********************************************************************
'Funciones que manejan la memoria
'*********************************************************************

Public Function General_Bytes_To_Megabytes(Bytes As Double) As Double
Dim dblAns As Double
dblAns = (Bytes / 1024) / 1024
General_Bytes_To_Megabytes = Format(dblAns, "###,###,##0.00")
End Function

Public Function General_Get_Free_Ram() As Double
    'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    General_Get_Free_Ram = General_Bytes_To_Megabytes(dblAns)
End Function

Public Function General_Get_Free_Ram_Bytes() As Long
    GlobalMemoryStatus pUdtMemStatus
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys
End Function

'*******************************************************************
'Colores
'*******************************************************************

Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal A As Long) As Long
        
    Dim c As Long
        
    If A > 127 Then
        A = A - 128
        c = A * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    Else
        c = A * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    End If
    
    ARGB = c

End Function
'*******************************************************************

Public Sub MostrarConnect()
'Sub creado para la llamada del frmConnect con el mapa, de lo contrario no funcionaria correctamente.
Dim SelectMapa As Byte
Dim MapC As Byte
Dim XC As Byte
Dim YC As Byte

    frmRenderConnect.Visible = True
    
    MapaConnect.Map = 166
    MapaConnect.X = 27
    MapaConnect.Y = 158
    DayStatus = 1
    
    Call SwitchMap(MapaConnect.Map, Get_Extract(Map, "Mapa" & MapaConnect.Map & ".csm"))
End Sub

Public Sub Auto_Drag(ByVal hwnd As Long)
'*****************************
'MOVER VENTANAS
'*****************************
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub

Public Sub GuardarOpciones()
    Dim Arch As String
    
    Arch = App.Path & "\init\AODConfig.bnd"
    
    If Not FileExist(Arch, vbNormal) Then
        Call MsgBox("No encontro el archivo ""\init\AODCOnfig.bnd"". Reinstale el juego, si el problema persiste contacte con los administradores.")
        End
    End If
    
    '[VIDEO]
    'Cambio de resolucion
    Call WriteVar(Arch, "VIDEO", "Resolucion", Val(Opciones.NoRes))
    Call WriteVar(Arch, "VIDEO", "BaseTecho", Val(Opciones.BaseTecho))
    Call WriteVar(Arch, "VIDEO", "VSynC", Val(Opciones.VSynC))
    Call WriteVar(Arch, "VIDEO", "VProcessing", Val(Opciones.VProcessing))
    
    '[AUDIO]
    'Enable / Disable audio
    Call WriteVar(Arch, "AUDIO", "Sonido", Opciones.Audio)
    Call WriteVar(Arch, "AUDIO", "VolAudio", Opciones.FXVolume)
    Call WriteVar(Arch, "AUDIO", "Musica", Opciones.sMusica)
    Call WriteVar(Arch, "AUDIO", "VolMusica", Opciones.MusicVolume)
    Call WriteVar(Arch, "AUDIO", "VolAmbient", Opciones.AmbientVol)
    Call WriteVar(Arch, "AUDIO", "Ambient", Opciones.Ambient)
    
    '[EXTRAS]
    Call WriteVar(Arch, "EXTRAS", "bCursores", Val(Opciones.bCursores))
    Call WriteVar(Arch, "EXTRAS", "MovEscritura", Val(Opciones.MovEscritura))
    Call WriteVar(Arch, "EXTRAS", "URLCON", Val(Opciones.URLCON))
    Call WriteVar(Arch, "EXTRAS", "NamePlayers", Val(Opciones.NamePlayers))
    Call WriteVar(Arch, "EXTRAS", "FirstRun", Val(Opciones.PrimeraVez))
    Call WriteVar(Arch, "EXTRAS", "GuildNews", Val(Opciones.GuildNews))
    Call WriteVar(Arch, "EXTRAS", "ClassicSpell", Val(Opciones.HechizosClasicos))
    Call WriteVar(Arch, "EXTRAS", "BloqCruceta", Val(Opciones.BloqCruceta))
    
End Sub

Public Sub CargarOpciones()
On Error GoTo errhandler
    Dim Leer As New clsIniReader
    
    If Not FileExist(App.Path & "\init\AODConfig.bnd", vbNormal) Then
        Call MsgBox("No encontro el archivo ""\init\AODCOnfig.bnd"". Reinstale el juego, si el problema persiste contacte con los administradores.")
        End
    End If
    
    Call Leer.Initialize(App.Path & "\init\AODConfig.bnd")
    
    'Carpeta Temporal
    Windows_Temp_Dir = General_Get_Temp_Dir
    Win2kXP = General_Windows_Is_2000XP
    Form_Caption = "AODrag v" & App.Major
    
    '[VIDEO]
    'Cambio de resolucion
    Opciones.NoRes = 1
    Opciones.BaseTecho = Val(Leer.GetValue("VIDEO", "BaseTecho"))
    If Opciones.BaseTecho < frmOpciones.SldTechos.min Then Opciones.BaseTecho = frmOpciones.SldTechos.min
    
    Opciones.VSynC = Val(Leer.GetValue("VIDEO", "VSynC"))
    Opciones.VProcessing = Val(Leer.GetValue("VIDEO", "VProcessing"))
    
    '[AUDIO]
    'Enable / Disable audio
    Opciones.Audio = Val(Leer.GetValue("AUDIO", "Sonido"))
    Opciones.FXVolume = Val(Leer.GetValue("AUDIO", "VolAudio"))
    Opciones.sMusica = Val(Leer.GetValue("AUDIO", "Musica"))
    Opciones.MusicVolume = Val(Leer.GetValue("AUDIO", "VolMusica"))
    Opciones.Ambient = Val(Leer.GetValue("AUDIO", "Ambient"))
    Opciones.AmbientVol = Val(Leer.GetValue("AUDIO", "VolAmbient"))
    
    '[EXTRAS]
    Opciones.bCursores = Val(Leer.GetValue("EXTRAS", "bCursores"))
    Opciones.MovEscritura = Val(Leer.GetValue("EXTRAS", "MovEscritura"))
    Opciones.URLCON = Val(Leer.GetValue("EXTRAS", "URLCON"))
    Opciones.NamePlayers = Val(Leer.GetValue("EXTRAS", "NamePlayers"))
    Opciones.PrimeraVez = Val(Leer.GetValue("EXTRAS", "FirstRun"))
    Opciones.GuildNews = Val(Leer.GetValue("EXTRAS", "GuildNews"))
    Opciones.HechizosClasicos = Val(Leer.GetValue("EXTRAS", "ClassicSpell"))
    Opciones.BloqCruceta = Val(Leer.GetValue("EXTRAS", "BloqCruceta"))
    
errhandler:
    Call LogError("SetItem::Error " & Err.number & " - " & Err.Description)
End Sub

Public Sub Relog()
   
    Call MostrarConnect
   
    EstadoLogin = E_MODO.LoginCuenta
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
           
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
    DoEvents

End Sub

Public Function GenerateKey() As String
Dim i As Byte, tempstring As String
    For i = 1 To 6
        If RandomNumber(1, 2) = 1 Then
            tempstring = tempstring & RandomNumber(1, 9)
        Else
            tempstring = tempstring & IIf(RandomNumber(1, 2) = 1, LCase$(Chr(97 + Rnd() * 862150000 Mod 26)), UCase$(Chr(97 + Rnd() * 862150000 Mod 26)))
        End If
    Next i
            
    GenerateKey = tempstring
End Function

Public Sub ChangeCursorMain(ByVal eCursor As eCursorState, ByVal frm As Form)
    If CurrentCursor <> eCursor Then
        Select Case eCursor
            Case cur_Normal
            frm.MousePointer = vbDefault

            Case cur_Action
            frm.MousePointer = 2

            Case cur_Wait
            frm.MousePointer = vbHourglass
        End Select
   
        CurrentCursor = eCursor
    End If
End Sub

Private Sub LoadCursor(ByVal sCursor As Byte, lHandle As Long)
    Dim GetCursor As Long
    If FileExist(App.Path & DirCursores & sCursor & ".ani", vbArchive) = True Then
        GetCursor = LoadCursorFromFile(App.Path & DirCursores & sCursor & ".ani")
        SetClassLong lHandle, -12, GetCursor
    End If
End Sub

Public Sub ListarServidores()
On Error Resume Next
    frmConnect.SVList.Clear
    Dim lista() As String
    Dim i As Byte
    lista = Split(frmMain.Inet1.OpenURL("https://aodrag.es/server-list.txt"), vbLf)

    Debug.Print lista(0)
    
    ' 26/01/2016 Irongete: Se pueden cargar hasta un maximo de 100 servidores
    For i = 0 To UBound(lista())
        Servidor(i).Ip = ReadField(1, lista(i), Asc("|"))
        Servidor(i).Puerto = ReadField(2, lista(i), Asc("|"))
        Servidor(i).Nombre = ReadField(3, lista(i), Asc("|"))
        Debug.Print Servidor(i).Nombre
        Debug.Print Servidor(i).Ip
        Debug.Print Servidor(i).Puerto
        frmConnect.SVList.AddItem Servidor(i).Nombre
    Next i
    
    CurServerIp = Servidor(1).Ip
    CurServerPort = Servidor(1).Puerto
    frmConnect.SVList.Text = Servidor(1).Nombre
End Sub

Public Sub ListarConsejos()
On Error Resume Next
    Dim i As Byte
    ListaConsejos = Split(frmMain.Inet1.OpenURL("http://www.aodrag.es/consejos.txt"), "|")
    
    For i = 1 To UBound(ListaConsejos())
        Consejos(i) = ReadField(1, ReadField(i, ListaConsejos(i), Asc("|")), Asc(":"))
    Next i
End Sub

Public Sub CerrarCuenta()
'Uso esto para cerrar el panel de cuenta y que no quede informacion suelta.
    frmCuenta.ListPJ.Clear
    Cuenta.name = ""
    Cuenta.Pass = ""
    
    Unload frmCuenta
End Sub

Public Sub ResetAllInfo()
    Dim i As Long
    
    'Unload all forms except frmMain, frmConnect and frmCrearPersonaje
    Dim frm As Form
    For Each frm In Forms
        If frm.name <> frmMain.name And frm.name <> frmRenderConnect.name And _
            frm.name <> frmCrearPersonaje.name Then
            
            Unload frm
        End If
    Next
    
    On Local Error GoTo 0
    
    frmConnect.MousePointer = vbNormal

'Hide main form
    frmMain.Visible = False
    frmMain.picSpell.Picture = LoadPicture("")
    
    'Stop audio
    Sound.Sound_Stop_All
    Sound.Ambient_Stop
    
    'Show connection form
    
    'Reset global vars
    UserDescansar = False
    UserParalizado = False
    pausa = False
    UserCiego = False
    UserMeditar = False
    UserNavegando = False
    UserMontando = False
    bFogata = False
    Comerciando = False
    
    UserMaxHP = 0
    UserMinHP = 0
    UserMaxMAN = 0
    UserMinMAN = 0
    UserFuerza = 0
    UserAgilidad = 0
    
    'Delete all kind of dialogs
    Call CleanDialogs
    
    'Reset some char variables...
    For i = 1 To LastChar
        charlist(i).invisible = False
        charlist(i).PartyId = 0
    Next i
    
    
End Sub

Public Sub LogError(ByVal sDesc As String)
'***************************************************
'Author: ^[GS]^
'Last Modification: 09/10/2012 - ^[GS]^
'***************************************************

On Error GoTo errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\Logs\Errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & sDesc
    Close #nfile
    
    Exit Sub

errhandler:

End Sub

Public Sub Make_Transparent_Richtext(ByVal hwnd As Long)
    If Win2kXP Then _
        Call SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
End Sub

Public Sub Make_Transparent_Form(ByVal hwnd As Long, Optional ByVal bytOpacity As Byte = 128)
    If Win2kXP Then
        Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(hwnd, 0, bytOpacity, LWA_ALPHA)
    End If
End Sub

Public Sub UnMake_Transparent_Form(ByVal hwnd As Long)
    If Win2kXP Then _
        Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And (Not WS_EX_TRANSPARENT))
End Sub

'*****************************************************************
'************************Generar particulas en Cuerpos************
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long
On Error Resume Next

If ParticulaInd <= 0 Then Exit Function

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Char_Particle_Create = Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).Angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
    StreamData(ParticulaInd).Radio)

End Function
'*******************************************************************
'******************Generar particulas en el mapa********************
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long
   
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)
 
General_Particle_Create = Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).Angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).Radio)
End Function
'*******************************************************************

Private Function SystemDrive() As String
' Obtiene la unidad del sistema
    Dim windows_dir As String
    Dim Length As Long
    windows_dir = Space$(255)
    Length = GetWindowsDirectory(windows_dir, Len(windows_dir))
    SystemDrive = Left$(windows_dir, 3) ' C:\
End Function

Function GetSerialHD() As Long
'Obtiene el numero de serie del disco de sistema
    Dim SerialNum As Long
    Dim res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    res = GetVolumeInformation(SystemDrive(), Temp1, _
    Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialHD = SerialNum
End Function

Public Function SEncriptar(ByVal Cadena As String) As String
'Encripta una cadena de texto
    Dim i As Long, RandomNum As Integer
    
    RandomNum = 99 * Rnd
    If RandomNum < 10 Then RandomNum = 10
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) + RandomNum)
    Next i
    SEncriptar = Cadena & Chr$(Asc(Left$(RandomNum, 1)) + 10) & Chr$(Asc(Right$(RandomNum, 1)) + 10)
    DoEvents

End Function

Public Function SDesencriptar(ByVal Cadena As String) As String
'Desencripta una cadena de texto
    Dim i As Long, NumDesencriptar As String
    
    NumDesencriptar = Chr$(Asc(Left$((Right(Cadena, 2)), 1)) - 10) & Chr$(Asc(Right$((Right(Cadena, 2)), 1)) - 10)
    Cadena = (Left$(Cadena, Len(Cadena) - 2))
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) - NumDesencriptar)
    Next i
    SDesencriptar = Cadena
    DoEvents

End Function

Public Function input_key_get(ByVal key_code As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    input_key_get = GetKeyState(key_code) And &H8000
End Function

Public Function ColorToDX8(ByVal long_color As Long) As Long
    ' DX8 engine
    Dim temp_color As String
    Dim red As Integer, blue As Integer, green As Integer
    
    temp_color = Hex$(long_color)
    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String(6 - Len(temp_color), "0") + temp_color
    End If
    
    red = CLng("&H" + mid$(temp_color, 1, 2))
    green = CLng("&H" + mid$(temp_color, 3, 2))
    blue = CLng("&H" + mid$(temp_color, 5, 2))
    
    ColorToDX8 = D3DColorXRGB(red, green, blue)

End Function

Public Function General_Windows_Is_2000XP() As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: Unknown
'
'**************************************************************

On Error GoTo ErrorHandler

Dim RetVal As Long

OSInfo.dwOSVersionInfoSize = Len(OSInfo)
RetVal = GetOSVersion(OSInfo)

If OSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT And OSInfo.dwMajorVersion >= 5 Then
    General_Windows_Is_2000XP = True
Else
    General_Windows_Is_2000XP = False
End If

Exit Function

ErrorHandler:
    General_Windows_Is_2000XP = False

End Function

Public Function PonerPuntos(ByVal Numero As Double) As String
'</Edurne>
Dim i As Integer
Dim Cifra As String
 
Cifra = str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next
 
PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
 
End Function

Public Function General_Distance_Get(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Integer
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: Unknown
'
'**************************************************************

General_Distance_Get = Abs(x1 - x2) + Abs(y1 - y2)

End Function
