Attribute VB_Name = "Zip"
Public Type ZIPUSERFUNCTIONS
DLLPrnt As Long
DLLPassword As Long
DLLComment As Long
DLLService As Long
End Type


Public Type ZPOPT
fSuffix As Long
fEncrypt As Long
fSystem As Long
fVolume As Long
fExtra As Long
fNoDirEntries As Long
fExcludeDate As Long
fIncludeDate As Long
fVerbose As Long
fQuiet As Long
fCRLF_LF As Long
fLF_CRLF As Long
fJunkDir As Long
fRecurse As Long
fGrow As Long
fForce As Long
fMove As Long
fDeleteEntries As Long
fUpdate As Long
fFreshen As Long
fJunkSFX As Long
fLatestTime As Long
fComment As Long
fOffsets As Long
fPrivilege As Long
fEncryption As Long
fRepair As Long
flevel As Byte
date As String
szRootDir As String
End Type

Public Type ZIPnames
    s(0 To 99) As String
End Type

Public Type CBChar
    ch(4096) As Byte
End Type
Public Declare Function ZpInit Lib "zip32.dll" (ByRef Zipfun As ZIPUSERFUNCTIONS) As Long
Public Declare Function ZpSetOptions Lib "zip32.dll" (ByRef Opts As ZPOPT) As Long
Public Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long
'Public Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As String) As Long
Function FuncionParaProcesarPassword(ByRef B1 As Byte, L As Long, ByRef B2 As Byte, ByRef B3 As Byte) As Long
    FuncionParaProcesarPassword = 0
End Function
Function FuncionParaProcesarServicios(ByRef Fname As CBChar, ByVal X As Long) As Long
    FuncionParaProcesarServicios = 0
End Function
Function FuncionParaProcesarMensajes(ByRef Fname As CBChar, ByVal X As Long) As Long
    FuncionParaProcesarMensajes = 0
End Function
Function FuncionParaProcesarComentarios(Comentario As CBChar) As CBChar
    Comentario.ch(0) = vbNullString
    FuncionParaProcesarComentarios = Comentario
End Function
Public Function DevolverDireccionMemoria(Direccion As Long) As Long
    DevolverDireccionMemoria = Direccion
End Function
Public Sub Comprimir()
Dim Resultado As Long
Dim intContadorFicheros As Integer

Dim FuncionesZip As ZIPUSERFUNCTIONS
Dim OpcionesZip As ZPOPT
Dim NombresFicherosZip As ZIPnames
'Dim NombresFicherosZip As String
FuncionesZip.DLLComment = DevolverDireccionMemoria(AddressOf FuncionParaProcesarComentarios)
FuncionesZip.DLLPassword = DevolverDireccionMemoria(AddressOf FuncionParaProcesarPassword)
FuncionesZip.DLLPrnt = DevolverDireccionMemoria(AddressOf FuncionParaProcesarMensajes)
FuncionesZip.DLLService = DevolverDireccionMemoria(AddressOf FuncionParaProcesarServicios)

'For intContadorFicheros = 0 To 1
    NombresFicherosZip.s(0) = "Fotos/AodraGfoto.jpg" 'File1.List(intContadorFicheros)
'Next
 'NombresFicherosZip = "Fotos/AodraGfoto.jpg" 'File1.List(intContadorFicheros)
 
Resultado = ZpInit(FuncionesZip)
Resultado = ZpSetOptions(OpcionesZip)
Resultado = ZpArchive(1, "Fotos/foto.Zip", NombresFicherosZip)
DoEvents

If Dir(App.Path & "\Fotos\AoDraGfoto.bmp") <> "" Then
Kill (App.Path & "\Fotos\AoDraGfoto.bmp")
End If
If Dir(App.Path & "\Fotos\AoDraGfoto.jpg") <> "" Then
Kill (App.Path & "\Fotos\AoDraGfoto.jpg")
End If
End Sub

