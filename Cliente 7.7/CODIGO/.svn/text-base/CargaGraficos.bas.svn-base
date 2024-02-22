Attribute VB_Name = "CargaGraficos"
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public str() As koj
'Public Const SRCCOPY = &HCC0020
Public Const DIB_RGB_COLORS = 0

Public Type koj
    where As Double
    valid As Byte
End Type

Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER
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

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

Public Type Bmpx
    gudtBMPFileHeader As BITMAPFILEHEADER   'Holds the file header
    gudtBMPInfo As BITMAPINFO               'Holds the bitmap info
    gudtBMPData() As Byte                   'Holds the pixel data
End Type

Public Bmx As Bmpx

Public Sub CargarGraficos()
    Dim num As Long, gr As Integer, n As Long

    'gr = FreeFile()
    'Open App.Path & "\Init\Gp.ind" For Binary Access Read As gr
    'Get gr, , num
   ReDim str(0 To 14550)
   ' Get gr, , str
   ' Close gr
   Call gpindex
   Call gpindex2
'pluto:2.14-----------
'For n = 0 To num
'Dim a As String
'a = str(n).where
'If str(n).where > 0 And n > 1 Then str(n).where = str(n).where - 3
'If n > 1 Then str(n).where = str(n).where + 315
'If n > 974 Then str(n).where = str(n).where + 143
'If n > 1364 Then str(n).where = str(n).where + 82
'If n > 6110 Then str(n).where = str(n).where + 8
'If n > 1246 Then str(n).where = str(n).where + 2
'End If 'where >
'a = str(n).where
'Next n
'--------------------
End Sub


Public Function FileSize(lngWidth As Long, lngHeight As Long) As Long
    If lngWidth Mod 4 > 0 Then
        FileSize = ((lngWidth \ 4) + 1) * 4 * lngHeight - 1
    Else
        FileSize = lngWidth * lngHeight - 1
    End If
End Function

Public Sub ExtractData(X As Integer, strFileName As String, lngOffset As Double)

Dim intBMPFile As Integer
Dim i As Integer

    Erase Bmx.gudtBMPInfo.bmiColors

    intBMPFile = FreeFile()
    Open strFileName For Binary Access Read Lock Write As intBMPFile

    Get intBMPFile, lngOffset, Bmx.gudtBMPFileHeader

    Get intBMPFile, , Bmx.gudtBMPInfo.bmiHeader
    If Bmx.gudtBMPInfo.bmiHeader.biClrUsed <> 0 Then
        For i = 0 To Bmx.gudtBMPInfo.bmiHeader.biClrUsed - 1
            Get intBMPFile, , Bmx.gudtBMPInfo.bmiColors(i).rgbBlue
            Get intBMPFile, , Bmx.gudtBMPInfo.bmiColors(i).rgbGreen
            Get intBMPFile, , Bmx.gudtBMPInfo.bmiColors(i).rgbRed
            Get intBMPFile, , Bmx.gudtBMPInfo.bmiColors(i).rgbReserved
        Next i
    ElseIf Bmx.gudtBMPInfo.bmiHeader.biBitCount = 8 Then
        Get intBMPFile, , Bmx.gudtBMPInfo.bmiColors
    End If

    If Bmx.gudtBMPInfo.bmiHeader.biBitCount = 8 Then
        ReDim Bmx.gudtBMPData(FileSize(Bmx.gudtBMPInfo.bmiHeader.biWidth, Bmx.gudtBMPInfo.bmiHeader.biHeight))
    Else
        If Bmx.gudtBMPInfo.bmiHeader.biSizeImage > 0 Then
            ReDim Bmx.gudtBMPData(Bmx.gudtBMPInfo.bmiHeader.biSizeImage - 1)
        Else
            If FileSize(Bmx.gudtBMPInfo.bmiHeader.biWidth, Bmx.gudtBMPInfo.bmiHeader.biHeight) < 1 Then
                Close intBMPFile
                Exit Sub
            Else
                ReDim Bmx.gudtBMPData(24 * FileSize(Bmx.gudtBMPInfo.bmiHeader.biWidth, Bmx.gudtBMPInfo.bmiHeader.biHeight) / 8)
            End If
        End If
    End If

    Get intBMPFile, , Bmx.gudtBMPData

    If Bmx.gudtBMPInfo.bmiHeader.biBitCount = 8 Then
        Bmx.gudtBMPFileHeader.bfOffBits = 1078
        Bmx.gudtBMPInfo.bmiHeader.biSizeImage = FileSize(Bmx.gudtBMPInfo.bmiHeader.biWidth, Bmx.gudtBMPInfo.bmiHeader.biHeight)
        Bmx.gudtBMPInfo.bmiHeader.biClrUsed = 0
        Bmx.gudtBMPInfo.bmiHeader.biClrImportant = 0
        Bmx.gudtBMPInfo.bmiHeader.biXPelsPerMeter = 0
        Bmx.gudtBMPInfo.bmiHeader.biYPelsPerMeter = 0
    End If
    Close intBMPFile
End Sub
