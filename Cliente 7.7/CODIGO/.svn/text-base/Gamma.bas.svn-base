Attribute VB_Name = "Gamma"
Dim GammaControler As DirectDrawGammaControl    'The object that gets/sets gamma ramps
Dim GammaRamp As DDGAMMARAMP                    'The gamma ramp we'll use to alter the screen state
Dim OriginalRamp As DDGAMMARAMP                 'The gamma ramp we'll use to store the original screen state
Public intRedVal As Integer                        'Store the currend red value w.r.t. original
Public intGreenVal As Integer                      'Store the currend green value w.r.t. original
Public intBlueVal As Integer                       'Store the currend blue value w.r.t. original

Private Function ConvToSignedValue(lngValue As Long) As Integer
'This was written by the same person who did the "updateGamma" code
    If lngValue <= 32767 Then
        ConvToSignedValue = CInt(lngValue)
        Exit Function
    End If
    ConvToSignedValue = lngValue - 65535
End Function
Private Function ConvToUnSignedValue(intValue As Integer) As Long
'This was written by the same person who did the "updateGamma" code
    If intValue >= 0 Then
        ConvToUnSignedValue = intValue
        Exit Function
    End If
    ConvToUnSignedValue = intValue + 65535
End Function

Sub SetGamma(intRed As Integer, intGreen As Integer, intBlue As Integer)
Dim i As Integer
    'Alter the gamma ramp to the percent given by comparing to original state
    'A value of zero ("0") for intRed, intGreen, or intBlue will result in the
    'gamma level being set back to the original levels. Anything ABOVE zero will
    'fade towards FULL colour, anything below zero will fade towards NO colour
    For i = 0 To 255
        If intRed < 0 Then GammaRamp.red(i) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.red(i)) * (100 - Abs(intRed)) / 100)
        If intRed = 0 Then GammaRamp.red(i) = OriginalRamp.red(i)
        If intRed > 0 Then GammaRamp.red(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.red(i))) * (100 - intRed) / 100))
        If intGreen < 0 Then GammaRamp.green(i) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.green(i)) * (100 - Abs(intGreen)) / 100)
        If intGreen = 0 Then GammaRamp.green(i) = OriginalRamp.green(i)
        If intGreen > 0 Then GammaRamp.green(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.green(i))) * (100 - intGreen) / 100))
        If intBlue < 0 Then GammaRamp.blue(i) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.blue(i)) * (100 - Abs(intBlue)) / 100)
        If intBlue = 0 Then GammaRamp.blue(i) = OriginalRamp.blue(i)
        If intBlue > 0 Then GammaRamp.blue(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.blue(i))) * (100 - intBlue) / 100))
    Next
    GammaControler.SetGammaRamp DDSGR_DEFAULT, GammaRamp
End Sub

Sub InitGamma()
    Set GammaControler = PrimarySurface.GetDirectDrawGammaControl
    GammaControler.GetGammaRamp DDSGR_DEFAULT, OriginalRamp
    intRedVal = 0
    intGreenVal = 0
    intBlueVal = 0
End Sub
