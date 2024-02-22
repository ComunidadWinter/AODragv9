Attribute VB_Name = "Fotos"
'Codigo:--------------------------------------------------------------------------------
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Sub FotoFichero()

' Input:
' theFile file Name with path, where you want the .bmp to be saved

'
'Output:
' True if successful
'

Dim lString As String
Dim aa As Integer
Dim thefile As String
tt:
aa = aa + 1
On Error GoTo Trap
thefile = ""
thefile = Val(aa) & ".bmp"
'Check if the File Exist
If Dir(App.Path & "\Fotos\AoDraG" & thefile) <> "" Then GoTo tt

'To get the Entire Screen
Call keybd_event(vbKeySnapshot, 1, 0, 0)

'To get the Active Window
'Call keybd_event(vbKeySnapshot, 0, 0, 0)
DoEvents
SavePicture Clipboard.GetData(vbCFBitmap), App.Path & "\Fotos\AoDraG" & thefile
AddtoRichTextBox frmMain.RecTxt, "Foto Guardada en " & App.Path & "\Fotos\AoDraG" & thefile, 87, 87, 87, 0, 0
fSaveGuiToFile = True
Exit Sub

Trap:
'Error handling
MsgBox "Error Occured in fSaveGuiToFile. Error #: " & Err.Number & ", " & Err.Description

End Sub
'---------------------------------------------------------
Public Sub FotoFichero2()



Dim lString2 As String
'Dim aa As Integer
Dim thefile2 As String
'tt:
'aa = aa + 1
On Error GoTo Trap2
thefile2 = "foto.bmp"
'thefile = Val(aa) & ".bmp"
'Check if the File Exist
If Dir(App.Path & "\Fotos\AoDraG" & thefile2) <> "" Then
Kill (App.Path & "\Fotos\AoDraG" & thefile2)
End If

'To get the Entire Screen
Call keybd_event(vbKeySnapshot, 1, 0, 0)

'To get the Active Window
'Call keybd_event(vbKeySnapshot, 0, 0, 0)
DoEvents
SavePicture Clipboard.GetData(vbCFBitmap), App.Path & "\Fotos\AoDraG" & thefile2
'AddtoRichTextBox frmMain.RecTxt, "Foto Guardada en " & App.Path & "\Fotos\AoDraG" & thefile, 116,116,116, 0, 0
fSaveGuiToFile = True


'PASAMOS A JPG
Dim Ruta As String
Dim RUTA2 As String
Ruta = App.Path & "\Fotos\AoDraG" & thefile2
RUTA2 = App.Path & "\Fotos\AoDraGfoto.jpg"
Call BMPtoJPG(Ruta, RUTA2)


Exit Sub
Trap2:
'Error handling
'MsgBox "Error Occured in fSaveGuiToFile. Error #: " & Err.Number & ", " & Err.Description

End Sub





