VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "jrInstalarFuente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Esta clase sirve para instalar y desinstalar fuentes True Type"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hPointer As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal hIDL As Long, ByVal cPath As String) As Boolean
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_FONTCHANGE = &H1D
Private Const CSIDL_FONTS = &H14


'variables locales para almacenar los valores de las propiedades
Private mvarFicheroTTF As String 'copia local
Private mvarPathTTF As String 'copia local
Private mvarDirectorioFuentesSistema As String 'copia local


Public Property Get DirectorioFuentesSistema() As String
Attribute DirectorioFuentesSistema.VB_Description = "Devuelve el directorio de fuentes del sistema (sólo lectura)."
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DirectorioFuentesSistema
    DirectorioFuentesSistema = mvarDirectorioFuentesSistema
End Property



Public Sub Desinstalar(Optional BorrarFichero As Boolean = False)
Attribute Desinstalar.VB_Description = "Desinstala la fuente del sistema y borra el fichero del directorio de fuentes del sistema si así se le indicó."
Dim res As Long

'compruebo que se hayan introducido los datos
If mvarFicheroTTF = "" Then
    Err.Raise vbObjectError + 3, , "No se ha introducido el nombre del fichero que contiene la fuente."
    Exit Sub
End If
res = RemoveFontResource(mvarFicheroTTF)
If res <> 0 Then
    'aviso a las demás aplicaciones del cambio
    res = SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0&, 0&)
    'borramos el fichero si BorrarFichero = true
    On Error Resume Next
    If BorrarFichero Then Kill mvarDirectorioFuentesSistema & "\" & mvarFicheroTTF
Else
    Err.Raise vbObjectError + 8, , "No se ha podido desinstalar la fuente " & mvarFicheroTTF
End If
End Sub

Public Sub Instalar()
Attribute Instalar.VB_Description = "Guarda una copia de la fuente en el directorio de fuentes del sistema y la instala."
Dim res As Long, FicheroFOT As String, DefaultDir As String
On Error Resume Next
'compruebo que se hayan introducido los datos
If mvarFicheroTTF = "" Then
    Err.Raise vbObjectError + 3, , "No se ha introducido el nombre del fichero que contiene la fuente."
    Exit Sub
End If
If mvarPathTTF = "" Then
    Err.Raise vbObjectError + 4, , "No se ha introducido el path del fichero que contiene la fuente."
    Exit Sub
End If
'compruebo que exista al fichero
If Dir(mvarPathTTF & mvarFicheroTTF) = "" Then
    Err.Raise vbObjectError + 5, , "No se ha podido encontrar el fichero " & mvarPathTTF & mvarFicheroTTF & "."
    Exit Sub
End If
'copio el fichero al directorio de fuentes del sistema

If FileExist(mvarDirectorioFuentesSistema & "\" & mvarFicheroTTF, vbNormal) Then Exit Sub

FileCopy mvarPathTTF & mvarFicheroTTF, mvarDirectorioFuentesSistema & "\" & mvarFicheroTTF
'genero el nombre del fichero FOT a partir del TTF
FicheroFOT = Left(mvarFicheroTTF, Len(mvarFicheroTTF) - 3) & "FOT"
'me aseguro de que el fichero fot se cree en el directorio de fuentes del sistema
DefaultDir = CurDir()
ChDir mvarDirectorioFuentesSistema
'creamos el font-resource
res = CreateScalableFontResource(0&, FicheroFOT, mvarFicheroTTF, mvarDirectorioFuentesSistema)
'si falla puede ser porque el fichero FOT ya exista, intentamos añadir la fuente de todas formas
'Añadimos el recurso a la tabla de fuentes del sistema
res = AddFontResource(mvarFicheroTTF)
If res > 0 Then
    'aviso a las demás aplicaciones del cambio
    res = SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0&, 0&)
Else
    Err.Raise vbObjectError + 6, , "Error ejecutando AddFontResource."
End If
End Sub

Public Property Let PathTTF(ByVal vData As String)
Attribute PathTTF.VB_Description = "Devuelve o establece la unidad y directorio donde se encuentra el fichero .ttf que queremos instalar (ejemplo: ""c:\\mi aplicación"")."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.PathTTF = 5
    mvarPathTTF = Trim(vData)
    'si el path no acaba en barra la pongo
    If Right(mvarPathTTF, 1) <> "\" Then mvarPathTTF = mvarPathTTF + "\"

End Property


Public Property Get PathTTF() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.PathTTF
    PathTTF = mvarPathTTF
End Property



Public Property Let FicheroTTF(ByVal vData As String)
Attribute FicheroTTF.VB_Description = "Devuelve os establece el nombre del fichero que contiene la fuente True Type que queremos instalar o desinstalar (ejemplo:""MiFuente.ttf"")."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FicheroTTF = 5
    mvarFicheroTTF = Trim(vData)
End Property


Public Property Get FicheroTTF() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FicheroTTF
    FicheroTTF = mvarFicheroTTF
End Property








Private Sub Class_Initialize()
'obtengo el directorio de fuentes del sistema
Dim ItemIDList As Long
Dim FolderPath As String * 260
Dim retVal As Long

ItemIDList = 0
retVal = SHGetSpecialFolderLocation(0&, CSIDL_FONTS, ItemIDList)
If retVal = 0 Then
  SHGetPathFromIDList ItemIDList, FolderPath
  CoTaskMemFree ItemIDList
  retVal = InStr(FolderPath, Chr$(0))
  If retVal > 0 Then
      mvarDirectorioFuentesSistema = Left(FolderPath, retVal - 1)
  Else
      mvarDirectorioFuentesSistema = ""
  End If
Else
  mvarDirectorioFuentesSistema = ""
End If
If mvarDirectorioFuentesSistema = "" Then
    'no he podido encontrar el directorio de fuentes del sistema
    Err.Raise vbObjectError + 1, , "No se ha podido determinar el directorio de fuentes del sistema."
End If
End Sub


