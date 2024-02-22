Attribute VB_Name = "Jpg2"
Option Explicit

' ----==== GDIPlus Const ====----
Private Const GdiPlusVersion As Long = 1
Private Const mimeJPG As String = "image/jpeg"
Private Const EncoderParameterValueTypeLong As Long = 4
Private Const EncoderQuality As String = _
    "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

' ----==== Sonstige Types ====----
Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' ----==== GDIPlus Types ====----
Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter(15) As EncoderParameter
End Type

Private Type ImageCodecInfo
    Clsid As GUID
    FormatID As GUID
    CodecNamePtr As Long
    DllNamePtr As Long
    FormatDescriptionPtr As Long
    FilenameExtensionPtr As Long
    MimeTypePtr As Long
    Flags As Long
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPatternPtr As Long
    SigMaskPtr As Long
End Type

' ----==== GDIPlus Enums ====----
Private Enum status 'GDI+ Status
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

' ----==== GDI+ API Declarationen ====----
Private Declare Function GdiplusStartup Lib "gdiplus" _
    (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, _
    Optional ByRef lpOutput As Any) As status

Private Declare Function GdiplusShutdown Lib "gdiplus" _
    (ByVal token As Long) As status

Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" _
    (ByVal FileName As Long, ByRef Bitmap As Long) As status

Private Declare Function GdipSaveImageToFile Lib "gdiplus" _
    (ByVal image As Long, ByVal FileName As Long, _
    ByRef clsidEncoder As GUID, _
    ByRef encoderParams As Any) As status

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
    (ByVal Bitmap As Long, ByRef hbmReturn As Long, _
    ByVal background As Long) As status

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" _
    (ByVal hbm As Long, ByVal hpal As Long, _
    ByRef Bitmap As Long) As status

Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" _
    (ByRef numEncoders As Long, ByRef Size As Long) As status

Private Declare Function GdipGetImageEncoders Lib "gdiplus" _
    (ByVal numEncoders As Long, ByVal Size As Long, _
    ByRef Encoders As Any) As status

Private Declare Function GdipDisposeImage Lib "gdiplus" _
    (ByVal image As Long) As status

' ----==== OLE API Declarations ====----
Private Declare Function CLSIDFromString Lib "ole32" _
    (ByVal str As Long, id As GUID) As Long

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" _
    (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, _
    lplpvObj As Object)

' ----==== Kernel API Declarations ====----
Private Declare Function lstrlenW Lib "kernel32" _
    (lpString As Any) As Long

Private Declare Function lstrcpyW Lib "kernel32" _
    (lpString1 As Any, lpString2 As Any) As Long

' ----==== Variablen ====----
Private GdipToken As Long
Private GdipInitialized As Boolean

Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As status
    ' Initialisieren der GDI+ Instanz
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Function


Private Function ShutdownGDIPlus() As status
   
    ShutdownGDIPlus = GdiplusShutdown(GdipToken)
End Function

Private Function Execute(ByVal lReturn As status) As status
    Dim lCurErr As status
    If lReturn = status.OK Then
        lCurErr = status.OK
    Else
        lCurErr = lReturn
        MsgBox "Error"
    End If
    Execute = lCurErr
End Function



Public Function SavePictureAsJPG(ByVal Pic As StdPicture, _
    ByVal FileName As String, Optional ByVal Quality As Long = 85) _
    As Boolean
    
    Dim retStatus As status
    Dim retVal As Boolean
    Dim lBitmap As Long
    'quitar esto
    Exit Function
    
    ' Erzeugt eine GDI+ Bitmap vom StdPicture Handle -> lBitmap
    retStatus = Execute(GdipCreateBitmapFromHBITMAP(Pic.Handle, 0, _
        lBitmap))
    
    If retStatus = OK Then
        
        Dim PicEncoder As GUID
        Dim tParams As EncoderParameters
        
        '// Ermitteln der CLSID vom mimeType Encoder
        retVal = GetEncoderClsid(mimeJPG, PicEncoder)
        If retVal = True Then
            
            If Quality > 100 Then Quality = 100
            If Quality < 0 Then Quality = 0
            
            ' Initialisieren der Encoderparameter
            tParams.Count = 1
            With tParams.Parameter(0) ' Quality
                ' Setzen der Quality GUID
                CLSIDFromString StrPtr(EncoderQuality), .GUID
                .NumberOfValues = 1
                .type = EncoderParameterValueTypeLong
                .Value = VarPtr(Quality)
            End With
            
            ' Speichert lBitmap als JPG
            retStatus = Execute(GdipSaveImageToFile(lBitmap, _
                StrPtr(FileName), PicEncoder, tParams))
            
            If retStatus = OK Then
                SavePictureAsJPG = True
            Else
                SavePictureAsJPG = False
            End If
        Else
            SavePictureAsJPG = False
            MsgBox "Konnte keinen passenden Encoder ermitteln.", _
            vbOKOnly, "Encoder Error"
        End If
        
        ' Lösche lBitmap
        Call Execute(GdipDisposeImage(lBitmap))
        
    End If
End Function

Private Function HandleToPicture(ByVal hGDIHandle As Long, _
    ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hpal As Long = 0) As StdPicture
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As IID
    Dim oPicture As IPicture
    
    ' Initialisiert die PICTDESC Structur
    With tPictDesc
        .cbSizeOfStruct = Len(tPictDesc)
        .picType = ObjectType
        .hgdiObj = hGDIHandle
        .hPalOrXYExt = hpal
    End With
    
    ' Initialisiert das IPicture Interface ID
    With IID_IPicture
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    ' Erzeugen des Objekts
    OleCreatePictureIndirect tPictDesc, IID_IPicture, True, oPicture
    
    ' Rückgabe des Pictureobjekts
    Set HandleToPicture = oPicture
    
End Function

Private Function GetEncoderClsid(mimeType As String, pClsid As GUID) _
    As Boolean
    
    Dim num As Long
    Dim Size As Long
    Dim pImageCodecInfo() As ImageCodecInfo
    Dim j As Long
    Dim Buffer As String
    
    Call GdipGetImageEncodersSize(num, Size)
    If (Size = 0) Then
        GetEncoderClsid = False  '// fehlgeschlagen
        Exit Function
    End If
    
    ReDim pImageCodecInfo(0 To Size \ Len(pImageCodecInfo(0)) - 1)
    Call GdipGetImageEncoders(num, Size, pImageCodecInfo(0))
    
    For j = 0 To num - 1
        Buffer = Space$(lstrlenW(ByVal pImageCodecInfo(j).MimeTypePtr))
        
        Call lstrcpyW(ByVal StrPtr(Buffer), ByVal _
            pImageCodecInfo(j).MimeTypePtr)
            
        If (StrComp(Buffer, mimeType, vbTextCompare) = 0) Then
            pClsid = pImageCodecInfo(j).Clsid
            Erase pImageCodecInfo
            GetEncoderClsid = True  '// erfolgreich
            Exit Function
        End If
    Next j
    
    Erase pImageCodecInfo
    GetEncoderClsid = False  '// fehlgeschlagen
End Function




Public Sub IniciarGDI(Flags As Boolean)
'If Flags Then
'Call Execute(StartUpGDIPlus(GdiPlusVersion))
'Else
'Call Execute(ShutdownGDIPlus)
'End If
End Sub




