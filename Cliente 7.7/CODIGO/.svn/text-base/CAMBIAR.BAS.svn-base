    '''''''''''Cambiar la resolucion
    
    Dim typDevM As typDevMODE
    Dim lRes As Long
    
    lRes = EnumDisplaySettings(0, 0, typDevM)
    
    'uso esta procedure para saber la actual rez
    'porque EnumDisplay me devuelve un valor
    'incorrecto (ie. a mi siempre me dio 640x480
    'aun estando en 800x600)
        Dim intWidth As Integer
        Dim intHeight As Integer

        oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
       oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
    'zzzzzzzzzzzzzzzzzzzzzzzzz

    If oldResWidth <> 800 And oldResHeight <> 600 Then
        'Call ChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = 16
        End With
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
        'If lRes = DISP_CHANGE_FAILED Then MsgBox "yay"
        'If lRes = DISP_CHANGE_SUCCESSFUL Then Call ChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
      Else
        bNoResChange = True
    End If