Attribute VB_Name = "ModDibujado"
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Public Sub ArrayToPicturePNG(ByRef byteArray() As Byte, ByRef imgDest As IPicture) ' GSZAO
    Call SetBitmapBits(imgDest.Handle, UBound(byteArray), byteArray(0))
End Sub

Public Function ArrayToPicture(inArray() As Byte, offset As Long, Size As Long) As IPicture
    
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(offset), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture)
            End If
        End If
    End If
End Function

Sub DrawGrhtoHdc(ByVal desthDC As Long, ByVal grh_index As Long, ByVal Picture As Picture, ByVal X As Long, ByVal Y As Long)
    On Error Resume Next
    
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    Dim screen_x As Integer
    Dim screen_y As Integer
    Dim DestRect As RECT
    
    If grh_index <= 0 Then Exit Sub
    
    DestRect.Left = X
    DestRect.Top = Y
    DestRect.Right = DestRect.Left + GrhData(grh_index).pixelWidth
    DestRect.Bottom = DestRect.Top + GrhData(grh_index).pixelHeight
    
    screen_x = DestRect.Left
    screen_y = DestRect.Top

    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If
    
    Dim bmpData As StdPicture
    
    Dim InfoHead As INFOHEADER
    Dim buffer() As Byte
   
    InfoHead = File_Find(App.Path & "\Recursos\Graphics.DRAG", CStr(GrhData(grh_index).FileNum) & ".bmp")
   
    If InfoHead.lngFileSize <> 0 Then
        Extract_File_Memory Graphics, App.Path & "\Recursos\", CStr(GrhData(grh_index).FileNum) & ".bmp", buffer()
        
        Set bmpData = ArrayToPicture(buffer(), 0, UBound(buffer) + 1)
        
        src_x = GrhData(grh_index).SX
        src_y = GrhData(grh_index).SY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
        
        hdcsrc = CreateCompatibleDC(desthDC)
        PrevObj = SelectObject(hdcsrc, bmpData)
        
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
        DeleteDC hdcsrc
        
        Set bmpData = Nothing
        Erase buffer
    End If

End Sub

Public Sub DrawHeadHdc(ByVal desthDC As Long, ByVal Cabeza As Byte, ByVal Picture As Picture, ByVal X As Long, ByVal Y As Long, ByVal Heading As Byte, ByVal EsCabeza As Boolean)
    Dim textureX1 As Integer
    Dim textureX2 As Integer
    Dim textureY1 As Integer
    Dim textureY2 As Integer
    Dim offsetX As Integer
    Dim offsetY As Integer
    Dim Texture As Long
    Dim bmpData As StdPicture
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    
    Dim InfoHead As INFOHEADER
    Dim buffer() As Byte
        If EsCabeza = True Then
            If heads(Cabeza).Texture <= 0 Then Exit Sub
            Texture = heads(Cabeza).Texture
        Else
            If Cascos(Cabeza).Texture <= 0 Then Exit Sub
            Texture = Cascos(Cabeza).Texture
        End If
        
        textureX2 = 27
        textureY2 = 32
 
        If EsCabeza = True Then
            textureX1 = heads(Cabeza).startX
            textureY1 = ((Heading - 1) * textureY2) + heads(Cabeza).startY
        Else
            textureX1 = Cascos(Cabeza).startX
            textureY1 = ((Heading - 1) * textureY2) + Cascos(Cabeza).startY
        End If
 
        offsetX = (textureX2) - 30
        offsetY = (textureY2) - 35
        
        InfoHead = File_Find(App.Path & "\Recursos\Graphics.DRAG", CStr(heads(Cabeza).Texture) & ".bmp")
   
    If InfoHead.lngFileSize <> 0 Then
        Extract_File_Memory Graphics, App.Path & "\Recursos\", CStr(heads(Cabeza).Texture) & ".bmp", buffer()
        
        Set bmpData = ArrayToPicture(buffer(), 0, UBound(buffer) + 1)
        
        src_x = textureX1
        src_y = textureY1
        src_width = (textureX2 + src_x)
        src_height = (textureY2 + src_y)
        
        hdcsrc = CreateCompatibleDC(desthDC)
        PrevObj = SelectObject(hdcsrc, bmpData)
        
        BitBlt desthDC, X, Y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
        DeleteDC hdcsrc
        
        Set bmpData = Nothing
        Erase buffer
    End If

End Sub
