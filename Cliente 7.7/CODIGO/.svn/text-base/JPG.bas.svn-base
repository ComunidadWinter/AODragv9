Attribute VB_Name = "JPG"
'**************************************
' Name: Convert BMP to JPG
' Description:I know what you're thinkin
'     g, just another code to turn a BMP file
'     into a JPG file. Well my code does do th
'     is, but its A LOT easier. All the other


'     submissions are complicated and use seve
    '     ral user controls/modules/class modules.
    '     This is a module with a function BMPtoJP
    '     G and THATS IT! All you need is one lone
    '     of code and your done! No other modules,
    '     or any other files are required. It uses
    '     VIC32.DLL, you can download it at http:/
    '     /education.uregina.ca/courosa/ecmp355/Hy
    '     perST/Vic32.dll. Vote if you like.
' By: Munkee
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:BMPtoJPG "c:\xxxx.bmp","c:\xxxx
'     .jpg" thats all you need!
'
'Side Effects:None that I know of.
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.41722/lngWId.1/qx
'     /vb/scripts/ShowCode.htm
'for details.
'**************************************


'declarations


Type imgdes
    ibuff As Long
    stx As Long
    sty As Long
    endx As Long
    endy As Long
    buffwidth As Long
    palette As Long
    colors As Long
    imgtype As Long
    bmh As Long
    hBitmap As Long
    End Type


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


Declare Function bmpinfo Lib "VIC32.DLL" (ByVal Fname As String, bdat As BITMAPINFOHEADER) As Long


Declare Function allocimage Lib "VIC32.DLL" (image As imgdes, ByVal wid As Long, ByVal leng As Long, ByVal BPPixel As Long) As Long


Declare Function loadbmp Lib "VIC32.DLL" (ByVal Fname As String, desimg As imgdes) As Long


Declare Sub freeimage Lib "VIC32.DLL" (image As imgdes)


Declare Function convert1bitto8bit Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes) As Long


Declare Sub copyimgdes Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes)


Declare Function savejpg Lib "VIC32.DLL" (ByVal Fname As String, srcimg As imgdes, ByVal Quality As Long) As Long
    'end declarations
    'the sub


Public Sub BMPtoJPG(Thebmp As String, Thejpg As String)
    Dim tmpimage As imgdes ' Image descriptors
    Dim tmp2image As imgdes
    Dim rcode As Long
    Dim Quality As Long
    Dim vbitcount As Long
    Dim bdat As BITMAPINFOHEADER ' Reserve space for BMP struct
    Dim bmp_fname As String
    Dim jpg_fname As String
    bmp_fname = Thebmp
    jpg_fname = Thejpg
    Quality = 15
    ' Get info on the file we're to load
    rcode = bmpinfo(bmp_fname, bdat)


    If (rcode <> NO_ERROR) Then
        MsgBox "Cannot find file", 0, "Error encountered!"
        Exit Sub
    End If
    vbitcount = bdat.biBitCount


    If (vbitcount >= 16) Then ' 16-, 24-, or 32-bit image is loaded into 24-bit buffer
        vbitcount = 24
    End If
    ' Allocate space for an image
    rcode = allocimage(tmpimage, bdat.biWidth, bdat.biHeight, vbitcount)


    If (rcode <> NO_ERROR) Then
        MsgBox "Not enough memory", 0, "Error encountered!"
        Exit Sub
    End If
    ' Load image
    rcode = loadbmp(bmp_fname, tmpimage)


    If (rcode <> NO_ERROR) Then
        freeimage tmpimage ' Free image on error
        MsgBox "Cannot load file", 0, "Error encountered!"
        Exit Sub
    End If


    If (vbitcount = 1) Then ' If we loaded a 1-bit image, convert to 8-bit grayscale
        ' because jpeg only supports 8-bit grays
        '     cale or 24-bit color images
        rcode = allocimage(tmp2image, bdat.biWidth, bdat.biHeight, 8)


        If (rcode = NO_ERROR) Then
            rcode = convert1bitto8bit(tmpimage, tmp2image)
            freeimage tmpimage ' Replace 1-bit image with grayscale image
            copyimgdes tmp2image, tmpimage
        End If
    End If
    ' Save image
    rcode = savejpg(jpg_fname, tmpimage, Quality)
    freeimage tmpimage
End Sub

        

