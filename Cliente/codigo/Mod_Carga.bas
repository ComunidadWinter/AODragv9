Attribute VB_Name = "Mod_Carga"
'*******************************************MODULO DE CARGA*********************************************
'AUTOR: MANUEL (LORWIK)
'DESCRIPCION: RECOPILACION DE TODOS LOS CODIGOS NECESARIOS PARA LA CARGA DEL CLIENTE
'*******************************************************************************************************

Option Explicit

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tSetupMods
    bNoRes      As Boolean
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera

'Numero total de Grh indexados. Lo pongo aqui por que uno o mas Sub's de este modulo lo necesitan consultar.
Public GrhCount As Long
Private file As String
Public NumCuerpos As Integer
Public NumHeads As Integer
Public NumCascos As Integer
Public NumFxs As Integer

Sub CargarAtaques()
    Dim N As Integer
    Dim i As Long
    Dim j As Byte
    Dim MisAtaques() As tIndiceAtaque
    
    N = FreeFile()
    file = Get_Extract(Scripts, "Ataques.ind")
    Open file For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumAtaques
    
    'Resize array
    ReDim MisAtaques(0 To NumAtaques) As tIndiceAtaque
    ReDim AtaqueData(0 To NumAtaques) As AtaqueAnimData
    
    For i = 1 To NumAtaques
        Get #N, , MisAtaques(i)
        
        If MisAtaques(i).Body(1) Then
        
            For j = 1 To 4
                InitGrh AtaqueData(i).AtaqueWalk(j), MisAtaques(i).Body(j), 0
            Next j
            
            AtaqueData(i).HeadOffset.x = MisAtaques(i).HeadOffsetX
            AtaqueData(i).HeadOffset.y = MisAtaques(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
    Delete_File file
End Sub

Public Sub CargarCascos()

    Dim i As Integer
    Dim j As Byte
        
        file = Get_Extract(Scripts, "Cascos.ind")
        Open file For Binary Access Read As #1
         
            Get #1, , NumCascos   'cantidad de cascos
             
            ReDim Cascos(1 To NumHeads) As tHead
             
            Dim Texture As Integer
            Dim temp    As Integer
            Dim startX  As Integer
            Dim startY  As Integer
            Dim skip As Byte
            
            For i = 1 To NumCascos
                Get #1, , Texture 'number of .bmp
                Get #1, , startX
                Get #1, , startY
             
                Cascos(i).Texture = Texture
                Cascos(i).startX = startX
                Cascos(i).startY = startY
                
            Next i
         
        Close #1
    Delete_File file
End Sub

Public Sub CargarCabezas()

    Dim i As Integer
    
        file = Get_Extract(Scripts, "Cabezas.ind")
        Open file For Binary Access Read As #1
    
            Get #1, , NumHeads   'cantidad de cabezas
             
            ReDim heads(1 To NumHeads) As tHead
             
            Dim Texture As Integer
            Dim temp    As Integer
            Dim startX  As Integer
            Dim startY  As Integer
            Dim skip As Byte
            
            For i = 1 To NumHeads
                Get #1, , Texture 'number of .bmp
                Get #1, , startX
                Get #1, , startY
             
                heads(i).Texture = Texture
                heads(i).startX = startX
                heads(i).startY = startY
                
            Next i
         
    Close #1
    Delete_File file
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim j As Byte
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    file = Get_Extract(Scripts, "Personajes.ind")
    Open file For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            For j = 1 To 4
                InitGrh BodyData(i).Walk(j), MisCuerpos(i).Body(j), 0
            Next j

            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
    Delete_File file
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    file = Get_Extract(Scripts, "Fxs.ind")
    
    N = FreeFile
    Open file For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
    
    Delete_File file
End Sub

Sub CargarAnimArmas()
    On Error Resume Next
    Dim i As Integer
    Dim Archivo As String
    Dim Leer As New clsIniReader
    
    'Armas
    file = Get_Extract(Scripts, "armas.dat")
    Archivo = file
    Leer.Initialize Archivo
    NumWeaponAnims = Val(Leer.GetValue("INIT", "NumArmas"))
            
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    For i = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(i).WeaponWalk(1), Val(Leer.GetValue("ARMA" & i, "Dir1")), 0
        InitGrh WeaponAnimData(i).WeaponWalk(2), Val(Leer.GetValue("ARMA" & i, "Dir2")), 0
        InitGrh WeaponAnimData(i).WeaponWalk(3), Val(Leer.GetValue("ARMA" & i, "Dir3")), 0
        InitGrh WeaponAnimData(i).WeaponWalk(4), Val(Leer.GetValue("ARMA" & i, "Dir4")), 0
    Next i
    
    Delete_File file

End Sub

Sub CargarColores()
On Error Resume Next
    Dim ArchivoC As String
    Dim Leer As New clsIniReader
    
    file = Get_Extract(Scripts, "colores.dat")
    ArchivoC = file
    Leer.Initialize ArchivoC
    
    If Not FileExist(ArchivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
      For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
            ColoresPJ(i) = D3DColorXRGB(CByte(GetVar(ArchivoC, CStr(i), "R")), CByte(GetVar(ArchivoC, CStr(i), "G")), CByte(GetVar(ArchivoC, CStr(i), "B")))
      Next i
   
      ' Crimi
     ColoresPJ(50) = D3DColorXRGB(CByte(GetVar(ArchivoC, "CR", "R")), CByte(GetVar(ArchivoC, "CR", "G")), CByte(GetVar(ArchivoC, "CR", "B")))
 
     ' Ciuda
      ColoresPJ(49) = D3DColorXRGB(CByte(GetVar(ArchivoC, "CI", "R")), CByte(GetVar(ArchivoC, "CI", "G")), CByte(GetVar(ArchivoC, "CI", "B")))
End Sub

Sub CargarAnimEscudos()
    Dim i As Integer
    Dim Archivo As String
    Dim Leer As New clsIniReader
    
    'Escudos
    file = Get_Extract(Scripts, "escudos.dat")
    Archivo = file
    
    Leer.Initialize Archivo
    
    NumEscudosAnims = Val(Leer.GetValue("INIT", "NumEscudos"))
            
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    For i = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(i).ShieldWalk(1), Val(Leer.GetValue("ESC" & i, "Dir1")), 0
        InitGrh ShieldAnimData(i).ShieldWalk(2), Val(Leer.GetValue("ESC" & i, "Dir2")), 0
        InitGrh ShieldAnimData(i).ShieldWalk(3), Val(Leer.GetValue("ESC" & i, "Dir3")), 0
        InitGrh ShieldAnimData(i).ShieldWalk(4), Val(Leer.GetValue("ESC" & i, "Dir4")), 0
    Next i
    Delete_File file
    
End Sub

Public Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim Handle As Integer
    Dim fileVersion As Long
    Dim file As String
    
    file = Get_Extract(Scripts, "Graficos.ind")
    
    'Open files
    Handle = FreeFile()
    Open file For Binary Access Read As Handle
    Seek #1, 1
    
    'Get file version
    Get Handle, , fileVersion
    
    'Get number of grhs
    Get Handle, , GrhCount
    
    'Resize arrays
    ReDim GrhData(1 To GrhCount) As GrhData
    
    While Not EOF(Handle)
        Get Handle, , Grh
        If Grh <> 0 Then
            With GrhData(Grh)
                'Get number of frames
                Get Handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                .Active = True
                
                ReDim .Frames(1 To GrhData(Grh).NumFrames)
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get Handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > GrhCount Then
                            GoTo ErrorHandler
                        End If
                    Next Frame
                    
                    Get Handle, , .Speed
                    
                    If .Speed <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get Handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , GrhData(Grh).SX
                    If .SX < 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .SY
                    If .SY < 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                End If
            End With
        End If
    Wend
    
    Close Handle
   Delete_File file

    LoadGrhData = True
Exit Function
 
ErrorHandler:
    LoadGrhData = False
    Debug.Print "Error en LoadGrhData... Grh: " & Grh
End Function

Public Sub LoadMiniMap()
On Error GoTo ErrorHandler
    Dim file As String
    Dim count As Long
    Dim Handle As Integer
    
    'Open files
    Handle = FreeFile()
    
    file = Get_Extract(Scripts, "minimap.dat")
    Open file For Binary As Handle
        Seek Handle, 1
        For count = 1 To GrhCount
            If GrhData(count).Active Then
                Get Handle, , GrhData(count).MiniMap_color
            End If
        Next count
    Close Handle
    Delete_File file
    
ErrorHandler:
    Debug.Print "Error en LoadMiniMap."
End Sub

Public Sub CargarPasos()

    ReDim Pasos(1 To NUM_PASOS) As tPaso
    
    Pasos(CONST_BOSQUE).CantPasos = 2
    ReDim Pasos(CONST_BOSQUE).Wav(1 To Pasos(CONST_BOSQUE).CantPasos) As Integer
    Pasos(CONST_BOSQUE).Wav(1) = 201
    Pasos(CONST_BOSQUE).Wav(2) = 202
    
    Pasos(CONST_NIEVE).CantPasos = 2
    ReDim Pasos(CONST_NIEVE).Wav(1 To Pasos(CONST_NIEVE).CantPasos) As Integer
    Pasos(CONST_NIEVE).Wav(1) = 199
    Pasos(CONST_NIEVE).Wav(2) = 200
    
    Pasos(CONST_CABALLO).CantPasos = 2
    ReDim Pasos(CONST_CABALLO).Wav(1 To Pasos(CONST_CABALLO).CantPasos) As Integer
    Pasos(CONST_CABALLO).Wav(1) = 23
    Pasos(CONST_CABALLO).Wav(2) = 24
    
    Pasos(CONST_DUNGEON).CantPasos = 2
    ReDim Pasos(CONST_DUNGEON).Wav(1 To Pasos(CONST_DUNGEON).CantPasos) As Integer
    Pasos(CONST_DUNGEON).Wav(1) = 23
    Pasos(CONST_DUNGEON).Wav(2) = 24
    
    Pasos(CONST_DESIERTO).CantPasos = 2
    ReDim Pasos(CONST_DESIERTO).Wav(1 To Pasos(CONST_DESIERTO).CantPasos) As Integer
    Pasos(CONST_DESIERTO).Wav(1) = 197
    Pasos(CONST_DESIERTO).Wav(2) = 198
    
    Pasos(CONST_PISO).CantPasos = 2
    ReDim Pasos(CONST_PISO).Wav(1 To Pasos(CONST_PISO).CantPasos) As Integer
    Pasos(CONST_PISO).Wav(1) = 23
    Pasos(CONST_PISO).Wav(2) = 24
    
    Pasos(CONST_PESADO).CantPasos = 3
    ReDim Pasos(CONST_PESADO).Wav(1 To Pasos(CONST_PESADO).CantPasos) As Integer
    Pasos(CONST_PESADO).Wav(1) = 220
    Pasos(CONST_PESADO).Wav(2) = 221
    Pasos(CONST_PESADO).Wav(3) = 222

End Sub

Public Sub CargarParticulas()
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    
On Error GoTo ErrorHandler
    
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim StreamFile As String
    Dim Leer As New clsIniReader
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "particulas.ini", Windows_Temp_Dir, False) Then
        Err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo ErrorHandler
    End If
    
    StreamFile = Windows_Temp_Dir & "Particulas.ini"
    Leer.Initialize StreamFile
    
    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
        For loopc = 1 To TotalStreams
            StreamData(loopc).name = Leer.GetValue(Val(loopc), "Name")
            StreamData(loopc).NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
            StreamData(loopc).X1 = Leer.GetValue(Val(loopc), "X1")
            StreamData(loopc).Y1 = Leer.GetValue(Val(loopc), "Y1")
            StreamData(loopc).X2 = Leer.GetValue(Val(loopc), "X2")
            StreamData(loopc).Y2 = Leer.GetValue(Val(loopc), "Y2")
            StreamData(loopc).angle = Leer.GetValue(Val(loopc), "Angle")
            StreamData(loopc).vecx1 = Leer.GetValue(Val(loopc), "VecX1")
            StreamData(loopc).vecx2 = Leer.GetValue(Val(loopc), "VecX2")
            StreamData(loopc).vecy1 = Leer.GetValue(Val(loopc), "VecY1")
            StreamData(loopc).vecy2 = Leer.GetValue(Val(loopc), "VecY2")
            StreamData(loopc).life1 = Leer.GetValue(Val(loopc), "Life1")
            StreamData(loopc).life2 = Leer.GetValue(Val(loopc), "Life2")
            StreamData(loopc).friction = Leer.GetValue(Val(loopc), "Friction")
            StreamData(loopc).spin = Leer.GetValue(Val(loopc), "Spin")
            StreamData(loopc).spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
            StreamData(loopc).spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
            StreamData(loopc).AlphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
            StreamData(loopc).gravity = Leer.GetValue(Val(loopc), "Gravity")
            StreamData(loopc).grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
            StreamData(loopc).bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
            StreamData(loopc).XMove = Leer.GetValue(Val(loopc), "XMove")
            StreamData(loopc).YMove = Leer.GetValue(Val(loopc), "YMove")
            StreamData(loopc).move_x1 = Leer.GetValue(Val(loopc), "move_x1")
            StreamData(loopc).move_x2 = Leer.GetValue(Val(loopc), "move_x2")
            StreamData(loopc).move_y1 = Leer.GetValue(Val(loopc), "move_y1")
            StreamData(loopc).move_y2 = Leer.GetValue(Val(loopc), "move_y2")
            StreamData(loopc).Radio = Val(Leer.GetValue(Val(loopc), "Radio"))
            StreamData(loopc).life_counter = Leer.GetValue(Val(loopc), "life_counter")
            StreamData(loopc).Speed = Val(Leer.GetValue(Val(loopc), "Speed"))
            StreamData(loopc).NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
           
            ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
            GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
           
            For i = 1 To StreamData(loopc).NumGrhs
                StreamData(loopc).grh_list(i) = ReadField(str(i), GrhListing, 44)
            Next i
            StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
            For ColorSet = 1 To 4
                TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
                StreamData(loopc).colortint(ColorSet - 1).r = ReadField(1, TempSet, 44)
                StreamData(loopc).colortint(ColorSet - 1).g = ReadField(2, TempSet, 44)
                StreamData(loopc).colortint(ColorSet - 1).b = ReadField(3, TempSet, 44)
            Next ColorSet
        
    Next loopc
    
    Delete_File Windows_Temp_Dir & "particulas.ini"
    Set Leer = Nothing
    
Exit Sub
    
ErrorHandler:
    If FileExist(Windows_Temp_Dir & "particulas.ini", vbNormal) Then Delete_File Windows_Temp_Dir & "particulas.ini"
    
End Sub

Sub CargarMensajes()
On Error Resume Next
    Dim ArchivoC As String
    Dim Leer As New clsIniReader
    Dim i As Long
    Dim CantMensajes As Byte
    
    file = Get_Extract(Scripts, "mensajes.dat")
    ArchivoC = file
    Leer.Initialize ArchivoC
    
    If Not FileExist(ArchivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se han podido cargar los mensajes. Falta el archivo mensajes.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    CantMensajes = CByte(GetVar(ArchivoC, "Main", "Cant"))
    
    For i = 0 To CantMensajes
        MultiMensaje(i).mensaje = CStr(GetVar(ArchivoC, "Mensajes", "msg" & i))
    Next i

End Sub
