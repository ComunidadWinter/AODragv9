Private Sub Form_Unload(Cancel As Integer)
    If bNoResChange = False Then
        Dim typDevM As typDevMODE
        Dim lRes As Long
    
        lRes = EnumDisplaySettings(0, 0, typDevM)
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
           .dmPelsHeight = oldResHeight
        End With
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If
End
End Sub