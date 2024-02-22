Attribute VB_Name = "Mod_General"
Option Explicit
'PLUTO:6.0a--------------------------------
Public Naci As String
Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO2
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_VENENO
    FONTTYPE_PLUTO
    FONTTYPE_COMERCIO
    FONTTYPE_GLOBAL
End Enum
Public Type tFont
    red As Byte
    green As Byte
    blue As Byte
    bold As Boolean
    italic As Boolean
End Type
Public FontTypes(19) As tFont
'--------------------------------------------


'Delzak
'Public TeclasModificadas As Boolean
Public bO As Integer
Public bK As Long


Public banners As String
Public bInvMod As Boolean  'El inventario se modificó?

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameLimiter As Long

Public lFrameModLimiter As Long
Public lFrameTimer As Long
Public sHKeys() As String

Public EsLocal As Boolean

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Function DirGraficos() As String
DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
If navida = 0 Then
DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
Else
DirMidi = App.Path & "\Midi\Navidad\"
End If
End Function
Public Function SD(ByVal n As Integer) As Integer
'Suma digitos
Dim auxint As Integer
Dim digit As Byte
Dim suma As Integer
auxint = n

Do
    digit = (auxint Mod 10)
    suma = suma + digit
    auxint = auxint \ 10

Loop While (auxint <> 0)

SD = suma

End Function

Public Function SDM(ByVal n As Integer) As Integer
'Suma digitos cada digito menos dos
Dim auxint As Integer
Dim digit As Integer
Dim suma As Integer
auxint = n

Do
    digit = (auxint Mod 10)
    
    digit = digit - 1
    
    suma = suma + digit
    
    auxint = auxint \ 10

Loop While (auxint <> 0)

SDM = suma

End Function

Public Function Complex(ByVal n As Integer) As Integer

If n Mod 2 <> 0 Then
    Complex = n * SD(n)
Else
    Complex = n * SDM(n)
End If

End Function

Public Function ValidarLoginMSG(ByVal n As Integer) As Integer
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(n)
AuxInteger2 = SDM(n)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Sub PlayWaveAPI(file As String)

'on error Resume Next
'Dim rc As Integer

'rc = sndPlaySound(file, SND_ASYNC)

End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Sub CargarAnimArmas()

On Error Resume Next

Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "armas.dat"
DoEvents

NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))

ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

For loopc = 1 To NumWeaponAnims
    InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
    InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
Next loopc

End Sub
Sub Bloqui()

a:
'Exit Sub
GoTo a
End Sub
Sub CargarAnimEscudos()

On Error Resume Next

Dim loopc As Integer
Dim arch As String
arch = App.Path & "\init\" & "escudos.dat"
DoEvents

NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))

ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData

For loopc = 1 To NumEscudosAnims
    InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
    InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
Next loopc

End Sub

Sub Addtostatus(RichTextBox As RichTextBox, Text As String, red As Byte, green As Byte, blue As Byte, bold As Byte, italic As Byte)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************

frmCargando.status.SelStart = Len(RichTextBox.Text)
frmCargando.status.SelLength = 0
frmCargando.status.SelColor = RGB(red, green, blue)

If bold Then
    frmCargando.status.SelBold = True
Else
    frmCargando.status.SelBold = False
End If

If italic Then
    frmCargando.status.SelItalic = True
Else
    frmCargando.status.SelItalic = False
End If

frmCargando.status.SelText = Chr(13) & Chr(10) & Text

End Sub

Sub AddtoRichTextBox(RichTextBox As RichTextBox, Text As String, Optional red As Integer = -1, Optional green As Integer, Optional blue As Integer, Optional bold As Boolean, Optional italic As Boolean, Optional bCrLf As Boolean)
        With RichTextBox
            If (Len(.Text)) > 20000 Then .Text = ""
            .SelStart = Len(RichTextBox.Text)
            .SelLength = 0
        
            .SelBold = IIf(bold, True, False)
            .SelItalic = IIf(italic, True, False)
            
            If Not red = -1 Then .SelColor = RGB(red, green, blue)
    
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
            
            RichTextBox.Refresh
        End With
    End Sub
'[END]'
Sub LimpiarRich(RichTextBox As RichTextBox, Text As String, Optional red As Integer = -1, Optional green As Integer, Optional blue As Integer, Optional bold As Boolean, Optional italic As Boolean, Optional bCrLf As Boolean)
        With RichTextBox
            .Text = ""
            .SelStart = Len(RichTextBox.Text)
            .SelLength = 0
        
            .SelBold = IIf(bold, True, False)
            .SelItalic = IIf(italic, True, False)
            
            If Not red = -1 Then .SelColor = RGB(red, green, blue)
    
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
            
           ' RichTextBox.Refresh
        End With
    End Sub

Sub AddtoTextBox(TextBox As TextBox, Text As String)
'******************************************
'Adds text to a text box at the bottom.
'Automatically scrolls to new text.
'******************************************

TextBox.SelStart = Len(TextBox.Text)
TextBox.SelLength = 0


TextBox.SelText = Chr(13) & Chr(10) & Text

End Sub
Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************

Dim loopc As Integer

For loopc = 1 To LastChar
    If CharList(loopc).Active = 1 Then
        MapData(CharList(loopc).pos.X, CharList(loopc).pos.Y).CharIndex = loopc
    End If
Next loopc

End Sub
Public Sub InitFonts()

    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        '.red = 255
       ' .green = 255
        '.blue = 255
        
        
        .red = 255
        .green = 255
        .blue = 255
        .bold = True
        .italic = True
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        '.red = 255
        .red = 255
        .bold = True
      
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        '.red = 32
        '.green = 51
        '.blue = 223
        '.bold = 1
        '.italic = 1
        .red = 255
        .green = 191
        .blue = 0
        .bold = 1
        .italic = 0
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        'pluto:7.0
        '.red = 65
        '.green = 190
        '.blue = 156
        .red = 128
        .green = 191
        .blue = 128
        
    End With
    'no se usa
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .red = 65
        .green = 190
        .blue = 156
        .bold = 1
    End With
    'no se usa
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .red = 191
        .green = 191
        .blue = 191
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        '.red = 255
        '.green = 180
        '.blue = 250
         .red = 191
        .green = 191
        .blue = 255
        .bold = 1
    End With
    
        With FontTypes(FontTypeNames.FONTTYPE_GLOBAL)
         .red = 128
        .green = 128
        .blue = 255
        .bold = 1
    End With
    
    'FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 255
    FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 94
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .red = 0
        .green = 191
        .blue = 0
        .bold = 1
    End With
    'no se usa
    FontTypes(FontTypeNames.FONTTYPE_SERVER).green = 185
    'no se usa
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .red = 0
        .green = 128
        .blue = 0
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        '.red = 130
        '.green = 130
        '.blue = 255
        .red = 64
        .green = 64
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
       ' .green = 200
        '.blue = 255
        .red = 19
        .green = 5
        .blue = 188
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .red = 255
        .green = 0
        .blue = 0
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .green = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_VENENO)
       
        .green = 94
      
 
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PLUTO)
        '.red = 255
        '.green = 150
        .red = 255
        .green = 191
        .blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_COMERCIO)
        '.red = 221
        '.green = 216
        '.blue = 9
         .red = 132
        .green = 132
        .blue = 0
        .bold = 1
    End With
End Sub

Sub SaveGameini()
'Grabamos los datos del usuario en el Game.ini

    Config_Inicio.name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort

Call EscribirGameIni(Config_Inicio)

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function



Function CheckUserData(checkemail As Boolean) As Boolean
'Validamos los datos del user
Dim loopc As Integer
Dim CharAscii As Integer

If checkemail Then
 If UserEmail = "" Then
    MsgBox ("Direccion de email invalida")
    Exit Function
 End If
End If

If UserPassword = "" Then
    MsgBox ("Ingrese un password.")
    Exit Function
End If

For loopc = 1 To Len(UserPassword)
    CharAscii = Asc(Mid$(UserPassword, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Password invalido.")
        Exit Function
    End If
Next loopc

If UserName = "" Then
    MsgBox ("Nombre invalido.")
    Exit Function
End If

If Len(UserName) > 30 Then
    MsgBox ("El nombre debe tener menos de 30 letras.")
    Exit Function
End If

For loopc = 1 To Len(UserName)

    CharAscii = Asc(Mid$(UserName, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Nombre invalido.")
        Exit Function
    End If
    
Next loopc


CheckUserData = True

End Function
Sub UnloadAllForms()
On Error Resume Next
    Dim mifrm As Form
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************

'if backspace allow
If KeyAscii = 8 Then
    LegalCharacter = True
    Exit Function
End If

'Only allow space,numbers,letters and special characters
If KeyAscii < 32 Or KeyAscii = 44 Then
    LegalCharacter = False
    Exit Function
End If

If KeyAscii > 126 Then
    LegalCharacter = False
    Exit Function
End If

'Check for bad special characters in between
If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
    LegalCharacter = False
    Exit Function
End If

'else everything is cool
LegalCharacter = True

End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************

'Set Connected
Connected = True

Call SaveGameini

'Unload the connect form
Unload frmConnect

frmMain.Label8.Caption = UserName
'Load main form
frmMain.Visible = True
frmMain.Sh.Enabled = True
End Sub
Sub CargarTip()
Exit Sub
Dim n As Integer
n = RandomNumber(1, UBound(Tips))
If n > UBound(Tips) Then n = UBound(Tips)
frmtip.tip.Caption = Tips(n)

End Sub

Sub MoveNorth()
If FPSFast = True Then
    If UserParalizado Then Exit Sub
End If
'pluto:2.18
 If CharList(UserCharIndex).Heading <> NORTH Then Call SendData("ª" & NORTH)
    
If UserParalizado Or UserMeditar Or UserDescansar Then Exit Sub

'pluto:2.3
If UserPeso > UserPesoMax Then
Call AddtoRichTextBox(frmMain.RecTxt, "Llevas demasiada carga, no puedes moverte.", 150, 150, 150, True, False, False)
Exit Sub
End If

If Cartel Then Cartel = False
If LegalPos(UserPos.X, UserPos.Y - 1) Then
    Call SendData("*")
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
        Call MoveCharbyHead(UserCharIndex, NORTH)
        Call MoveScreen(NORTH)
        DoFogataFx
    End If
'Else
   
   'End If
'pluto:2.8.0
'If CurMap = 192 And MapData(UserPos.x, UserPos.Y - 1).CharIndex > 0 Then
'If CharList(MapData(UserPos.x, UserPos.Y - 1).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList(MapData(UserPos.x, UserPos.Y - 1).CharIndex).Body.Walk(1).GrhIndex < 4525 Then
'SendData ("BOLL" & NORTH & "," & MapData(UserPos.x, UserPos.Y - 1).CharIndex)
'End If

'End If


End If
End Sub

Sub MoveEast()
If FPSFast = True Then
    If UserParalizado Then Exit Sub
End If
'pluto:2.18
 If CharList(UserCharIndex).Heading <> EAST Then Call SendData("ª" & EAST)

'pluto:2.17
If UserParalizado Or UserMeditar Or UserDescansar Then Exit Sub

'pluto:2.3
If UserPeso > UserPesoMax Then
Call AddtoRichTextBox(frmMain.RecTxt, "Llevas demasiada carga, no puedes moverte.", 150, 150, 150, True, False, False)
Exit Sub
End If

If Cartel Then Cartel = False
If LegalPos(UserPos.X + 1, UserPos.Y) Then
    Call SendData("+")
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
        Call MoveCharbyHead(UserCharIndex, EAST)
        Call MoveScreen(EAST)
        Call DoFogataFx
    End If
Else
  '  If CharList(UserCharIndex).Heading <> EAST Then
   '         Call SendData("ª" & EAST)
   ' End If
'pluto:2.8.0
'If CurMap = 192 And MapData(UserPos.x + 1, UserPos.Y).CharIndex > 0 Then
'If CharList(MapData(UserPos.x + 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList(MapData(UserPos.x + 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex < 4525 Then
'SendData ("BOLL" & EAST & "," & MapData(UserPos.x + 1, UserPos.Y).CharIndex)
'End If
'End If

End If
End Sub

Sub MoveSouth()
If FPSFast = True Then
    If UserParalizado Then Exit Sub
End If
'pluto:2.18
 If CharList(UserCharIndex).Heading <> SOUTH Then Call SendData("ª" & SOUTH)

'pluto:2.17
If UserParalizado Or UserMeditar Or UserDescansar Then Exit Sub

'pluto:2.3
If UserPeso > UserPesoMax Then
Call AddtoRichTextBox(frmMain.RecTxt, "Llevas demasiada carga, no puedes moverte.", 150, 150, 150, True, False, False)
Exit Sub
End If

If Cartel Then Cartel = False

If LegalPos(UserPos.X, UserPos.Y + 1) Then
    Call SendData("=")
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
        MoveCharbyHead UserCharIndex, SOUTH
        MoveScreen SOUTH
        DoFogataFx
    End If
'Else
    'If CharList(UserCharIndex).Heading <> SOUTH Then
     '       Call SendData("ª" & SOUTH)
    'End If

'pluto:2.8.0
'If CurMap = 192 And MapData(UserPos.x, UserPos.Y + 1).CharIndex > 0 Then
'If CharList(MapData(UserPos.x, UserPos.Y + 1).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList(MapData(UserPos.x, UserPos.Y + 1).CharIndex).Body.Walk(1).GrhIndex < 4525 Then
'SendData ("BOLL" & SOUTH & "," & MapData(UserPos.x, UserPos.Y + 1).CharIndex)
'End If
'End If


End If
End Sub

Sub MoveWest()
If FPSFast = True Then
    If UserParalizado Then Exit Sub
End If
'pluto:2.18
 If CharList(UserCharIndex).Heading <> WEST Then Call SendData("ª" & WEST)

'pluto:2.17
If UserParalizado Or UserMeditar Or UserDescansar Then Exit Sub


'pluto:2.3
If UserPeso > UserPesoMax Then
Call AddtoRichTextBox(frmMain.RecTxt, "Llevas demasiada carga, no puedes moverte.", 150, 150, 150, True, False, False)
Exit Sub
End If

If Cartel Then Cartel = False
If LegalPos(UserPos.X - 1, UserPos.Y) Then
    Call SendData("M")
    If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
            MoveCharbyHead UserCharIndex, WEST
            MoveScreen WEST
            DoFogataFx
    End If
'Else
 '   If CharList(UserCharIndex).Heading <> WEST Then
  '          Call SendData("ª" & WEST)
   ' End If
'pluto:2.8.0
'If CurMap = 192 And MapData(UserPos.x - 1, UserPos.Y).CharIndex > 0 Then
'If CharList(MapData(UserPos.x - 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList(MapData(UserPos.x - 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex < 4525 Then
'SendData ("BOLL" & WEST & "," & MapData(UserPos.x - 1, UserPos.Y).CharIndex)
'End If
'End If



End If
End Sub

Sub RandomMove()
If FPSFast = True Then
    If UserParalizado Then Exit Sub
End If
Dim j As Integer

j = RandomNumber(1, 4)

Select Case j
    Case 1
        Call MoveEast
    Case 2
        Call MoveNorth
    Case 3
        Call MoveWest
    Case 4
        Call MoveSouth
End Select

End Sub

Sub CheckKeys()
On Error Resume Next

'*****************************************************************
'Checks keys and respond
'*****************************************************************
Static KeyTimer As Integer

'Makes sure keys aren't being pressed to fast
If KeyTimer > 0 Then
    KeyTimer = KeyTimer - 1
    Exit Sub
End If



'Don't allow any these keys during movement..
If UserMoving = 0 Then
    If Not UserEstupido Then
            'Move Up
            'If GetAsyncKeyState(vbKeyUp) < 0 Then
            If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call MoveNorth
                Exit Sub
            End If
        
            'Move Right
            'If GetAsyncKeyState(vbKeyRight) < 0 And GetAsyncKeyState(vbKeyShift) >= 0 Then
                If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then

                Call MoveEast
                Exit Sub
            End If
        
            'Move down
            'If GetAsyncKeyState(vbKeyDown) < 0 Then
            If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then

                Call MoveSouth
                Exit Sub
            End If
        
            'Move left
            'If GetAsyncKeyState(vbKeyLeft) < 0 And GetAsyncKeyState(vbKeyShift) >= 0 Then
                  If GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then

                  Call MoveWest
                  Exit Sub
            End If
    Else
        Dim kp As Boolean
        kp = (GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
        GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
        GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
        GetAsyncKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
        If kp Then Call RandomMove
    End If
End If

End Sub

Sub MoveScreen(Heading As Byte)
If FPSFast = True Then
    If UserParalizado Then Exit Sub
End If
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select

'Fill temp pos
tX = UserPos.X + X
tY = UserPos.Y + Y

If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1

    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
        
'        If FramesPerSec > 20 Then Sleep 3000
        
'        If FramesPerSec > 18 Or FramesPerSec < 17 Then
 '       Select Case FramesPerSecCounter
  '          Case 18 To 17
   '             lFrameModLimiter = 55
    '        Case 16
     '           lFrameModLimiter = 50
      '      Case 15
       '         lFrameModLimiter = 45
        '    Case 14
         '       lFrameModLimiter = 40
          '  Case 15
           '     lFrameModLimiter = 35
            'Case 14
'                lFrameModLimiter = 30
 '           Case 13 To 0
  '              lFrameModLimiter = 25
   '     End Select
    '    End If
    '[END]'

    Call DoFogataFx
End If

End Sub

Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************

Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function

Public Function DirMapas() As String
DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
End Function
Sub SwitchMap(Map As Integer)

Dim loopc As Integer
Dim Y As Integer
Dim X As Integer
Dim tempint As Integer
      Dim datofantasma As Integer

Open DirMapas & "Mapa" & Map & ".map" For Binary As #1
Seek #1, 1
        
'map Header
Get #1, , MapInfo.MapVersion
Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
        
'Load arrays
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize


        '.dat file
        Get #1, , MapData(X, Y).Trigger
        
        For loopc = 4 To 1 Step -1
            Get #1, , datofantasma
            Get #1, , MapData(X, Y).Graphic(loopc).GrhIndex
            
            'Set up GRH
            If MapData(X, Y).Graphic(loopc).GrhIndex > 0 Then
                InitGrh MapData(X, Y).Graphic(loopc), MapData(X, Y).Graphic(loopc).GrhIndex
            End If
            
        Next loopc
        Get #1, , tempint
    
        Get #1, , MapData(X, Y).Blocked

               
        Get #1, , tempint
        
    'Erase NPCs
    If MapData(X, Y).CharIndex > 0 Then
        'pluto:6.5
       ' If x = 0 Or y = 0 Then
       'AddtoRichTextBox frmMain.RecTxt, "Fallo switchMap: " & CharList(MapData(x, y).CharIndex).nombre & " " & x & y, 255, 255, 255, 1, 1
        'End If
        
        'pluto:6.5--------------------
        If CharList(MapData(X, Y).CharIndex).pos.X = 0 Or CharList(MapData(X, Y).CharIndex).pos.Y = 0 Then
        'AddtoRichTextBox frmMain.RecTxt, "Fallo2 switchMap: " & CharList(MapData(x, y).CharIndex).Nombre & " x:" & x & " y: " & y, 0, 0, 0, 1, 1
        MapData(X, Y).CharIndex = 0
        Else
        Call EraseChar(MapData(X, Y).CharIndex)
        End If
        '----------------------------
    End If
        
        'Erase OBJs
        MapData(X, Y).objgrh.GrhIndex = 0

    Next X
Next Y

Close #1

MapInfo.name = ""
MapInfo.Music = ""

CurMap = Map
If HayMiniMap = True Then
Mod_TileEngine.GenerarMiniMapa 'minimap
End If
End Sub


Public Function ReadField(pos As Integer, Text As String, SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************

Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = pos Then
    ReadField = Mid(Text, LastPos + 1)
End If


End Function

Function FileExist(file As String, FileType As VbFileAttribute) As Boolean
If Dir(file, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function

Sub WriteClientVer()

Dim hFile As Integer
    
hFile = FreeFile()
Open App.Path & "\init\Ver.bin" For Binary Access Write As #hFile
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)

Put #hFile, , CInt(App.Major)
Put #hFile, , CInt(App.Minor)
Put #hFile, , CInt(App.Revision)

Close #hFile

End Sub






Public Function CurServerPasRecPort() As Integer
    'pluto:6.4
    If ServActual = 1 Then
    CurServerPasRecPort = "7666"
    Else
    'CurServerPasRecPort = "10281"
    CurServerPasRecPort = "7666"
    End If
End Function


Public Function CurServerIp() As String
'nati:modifico esto tambien
'18.231.49.236:9000

If EsLocal Then
    CurServerIp = "127.0.0.1"
Else
    If ServActual = 1 Then
      CurServerIp = "127.0.0.1"
    Else
      CurServerIp = "127.0.0.1"
    End If
End If
Debug.Print CurServerIp
End Function

Public Function CurServerPort() As Integer
    'pluto:6.4
    If ServActual = 1 Then
    CurServerPort = "7666" '7664 para pruebas
    frmMain.Socket1.Disconnect
    Else
    CurServerPort = "7666"
    'CurServerPort = "7666"
    frmMain.Socket1.Disconnect
    End If
Debug.Print CurServerPort


End Function



Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
    Dim t() As String
    Dim i As Long
    
    Dim UpToDate As Boolean
    Dim Patch As String
    
    'Parseo los comandos
    t = Split(command, " ")
    For i = LBound(t) To UBound(t)
        Select Case UCase$(t(i))
            Case "/NORES" 'no cambiar la resolucion


            Case "/LOCAL"
                EsLocal = True
        End Select
    Next i
    

End Sub




Sub Main()
On Error Resume Next

Dim flechudo As Integer
SetKey ("CLIENTE AODRAG v5.0")

Call WriteClientVer

If App.PrevInstance Then
    Call MsgBox("¡Aodrag ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    End
End If



Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 6) As Long
Dim ix As Byte
Dim IX2 As Byte
'pluto:7.0
'-------------
'pluto:6.8
FormP.Show
FormP.Visible = False
'-------
'pluto:6.0a tipos de mapas
'seguros para todos
Segura(1) = 1
Segura(81) = 1
Segura(183) = 1
Segura(184) = 1
Segura(34) = 1
Segura(20) = 1
'para crimis
Segura(170) = 2
Segura(62) = 2
Segura(63) = 2
Segura(64) = 2
'para ciudas
Segura(58) = 3
Segura(59) = 3
Segura(60) = 3
Segura(61) = 3
Segura(83) = 3
Segura(84) = 3
Segura(85) = 3
Segura(66) = 3
 'inseguras
 Segura(150) = 4
Segura(151) = 4
Segura(157) = 4
Segura(111) = 4
Segura(112) = 4
'para newbies
Segura(261) = 5
Segura(262) = 5
Segura(263) = 5
Segura(264) = 5
Segura(200) = 5
Segura(201) = 5
Segura(233) = 5
Segura(234) = 5
Segura(235) = 5
Segura(71) = 5
Segura(72) = 5
Segura(73) = 5
Segura(205) = 5
Segura(206) = 5
Segura(207) = 5
Segura(208) = 5
Segura(218) = 5
Segura(215) = 5
Segura(96) = 5
Segura(97) = 5
Segura(98) = 5
Segura(19) = 5
Segura(24) = 5
Segura(27) = 5
Segura(266) = 5
Segura(267) = 5
Segura(243) = 5
Segura(78) = 5
Segura(79) = 5
Segura(80) = 5
'pluto:7.0
'Luzaviso dificultad mapa
luzaviso(1) = 0 '(mapa seguro)
luzaviso(2) = 5
luzaviso(3) = 9
luzaviso(4) = 10
luzaviso(5) = 12
luzaviso(6) = 10
luzaviso(7) = 9
luzaviso(8) = 9
luzaviso(9) = 9
luzaviso(10) = 12
luzaviso(11) = 10
luzaviso(12) = 10
luzaviso(13) = 12
luzaviso(14) = 9
luzaviso(15) = 14
luzaviso(16) = 14
luzaviso(17) = 15
luzaviso(18) = 14
luzaviso(19) = 12
luzaviso(20) = 12
luzaviso(21) = 15
luzaviso(22) = 12
luzaviso(23) = 9
luzaviso(24) = 12
luzaviso(25) = 12
luzaviso(26) = 24
luzaviso(27) = 14
luzaviso(28) = 5
luzaviso(29) = 14
luzaviso(30) = 16
luzaviso(31) = 9
luzaviso(32) = 9
luzaviso(33) = 22
luzaviso(34) = 0 '(mapa seguro)
luzaviso(35) = 9
luzaviso(36) = 16
luzaviso(37) = 14
luzaviso(38) = 18
luzaviso(39) = 18
luzaviso(40) = 22
luzaviso(41) = 0
luzaviso(42) = 0
luzaviso(43) = 20
luzaviso(44) = 20
luzaviso(45) = 18
luzaviso(46) = 6
luzaviso(47) = 0
luzaviso(48) = 50
luzaviso(49) = 0
luzaviso(50) = 18
luzaviso(51) = 20
luzaviso(52) = 9
luzaviso(53) = 18
luzaviso(54) = 20
luzaviso(55) = 16
luzaviso(56) = 14
luzaviso(57) = 16
luzaviso(58) = 6
luzaviso(59) = 2
luzaviso(60) = 9
luzaviso(61) = 0
luzaviso(62) = 0
luzaviso(63) = 0
luzaviso(64) = 0
luzaviso(65) = 10
luzaviso(66) = 9
luzaviso(67) = 12
luzaviso(68) = 11
luzaviso(69) = 20
luzaviso(70) = 12
luzaviso(71) = 16
luzaviso(72) = 16
luzaviso(73) = 18
luzaviso(74) = 18
luzaviso(75) = 16
luzaviso(76) = 20
luzaviso(77) = 0
luzaviso(78) = 10
luzaviso(79) = 18
luzaviso(80) = 16
luzaviso(81) = 18
luzaviso(82) = 0
luzaviso(83) = 0
luzaviso(84) = 0
luzaviso(85) = 16
luzaviso(86) = 0
luzaviso(87) = 0
luzaviso(88) = 0
luzaviso(89) = 0
luzaviso(90) = 20
luzaviso(91) = 0
luzaviso(92) = 0
luzaviso(93) = 0
luzaviso(94) = 0
luzaviso(95) = 0
luzaviso(96) = 10
luzaviso(97) = 12
luzaviso(98) = 12
luzaviso(99) = 0
luzaviso(100) = 0
luzaviso(101) = 0
luzaviso(102) = 0
luzaviso(103) = 0
luzaviso(104) = 0
luzaviso(105) = 0
luzaviso(106) = 0
luzaviso(107) = 0
luzaviso(108) = 20
luzaviso(109) = 0
luzaviso(110) = 50
luzaviso(111) = 0
luzaviso(112) = 20
luzaviso(113) = 50
luzaviso(114) = 50
luzaviso(115) = 22
luzaviso(116) = 24
luzaviso(117) = 0
luzaviso(118) = 0
luzaviso(119) = 0
luzaviso(120) = 18
luzaviso(121) = 0
luzaviso(122) = 0
luzaviso(123) = 0
luzaviso(124) = 18
luzaviso(125) = 16
luzaviso(126) = 0
luzaviso(127) = 0
luzaviso(128) = 0
luzaviso(129) = 0
luzaviso(130) = 0
luzaviso(131) = 0
luzaviso(132) = 16
luzaviso(133) = 0
luzaviso(134) = 16
luzaviso(135) = 0
luzaviso(136) = 0
luzaviso(137) = 0
luzaviso(138) = 0
luzaviso(139) = 40
luzaviso(140) = 38
luzaviso(141) = 38
luzaviso(142) = 40
luzaviso(143) = 16
luzaviso(144) = 50
luzaviso(145) = 43
luzaviso(146) = 43
luzaviso(147) = 0
luzaviso(148) = 12
luzaviso(149) = 0
luzaviso(150) = 0
luzaviso(151) = 0
luzaviso(152) = 0
luzaviso(153) = 0
luzaviso(154) = 16
luzaviso(155) = 0
luzaviso(156) = 30
luzaviso(157) = 0
luzaviso(158) = 40
luzaviso(159) = 50
luzaviso(160) = 50
luzaviso(161) = 35
luzaviso(162) = 35
luzaviso(163) = 0
luzaviso(164) = 0
luzaviso(165) = 0
luzaviso(166) = 50 'Castillo norte
luzaviso(167) = 50 'Castillo sur
luzaviso(168) = 50 'Castillo este
luzaviso(169) = 50 'Castillo oeste
luzaviso(170) = 0
luzaviso(171) = 50
luzaviso(172) = 42
luzaviso(173) = 42
luzaviso(174) = 42
luzaviso(175) = 42
luzaviso(176) = 42
luzaviso(177) = 30
luzaviso(178) = 35
luzaviso(179) = 35
luzaviso(180) = 0
luzaviso(181) = 6
luzaviso(182) = 0
luzaviso(183) = 0
luzaviso(184) = 0
luzaviso(185) = 50 'fortaleza
luzaviso(186) = 0
luzaviso(187) = 0
luzaviso(188) = 0
luzaviso(189) = 20
luzaviso(190) = 0
luzaviso(191) = 0
luzaviso(192) = 0
luzaviso(193) = 24
luzaviso(194) = 0
luzaviso(195) = 0
luzaviso(196) = 10
luzaviso(197) = 16
luzaviso(198) = 16
luzaviso(199) = 16
luzaviso(200) = 18
luzaviso(201) = 16
luzaviso(202) = 18
luzaviso(203) = 16
luzaviso(204) = 16
luzaviso(205) = 14
luzaviso(206) = 14
luzaviso(207) = 14
luzaviso(208) = 12
luzaviso(209) = 0
luzaviso(210) = 0
luzaviso(211) = 0
luzaviso(212) = 0
luzaviso(213) = 0
luzaviso(214) = 0
luzaviso(215) = 9
luzaviso(216) = 42
luzaviso(217) = 36
luzaviso(218) = 18
luzaviso(219) = 0
luzaviso(220) = 0
luzaviso(221) = 0
luzaviso(222) = 0
luzaviso(223) = 0
luzaviso(224) = 14
luzaviso(225) = 16
luzaviso(226) = 14
luzaviso(227) = 14
luzaviso(228) = 14
luzaviso(229) = 14
luzaviso(230) = 0
luzaviso(231) = 0
luzaviso(232) = 0
luzaviso(233) = 16
luzaviso(234) = 12
luzaviso(235) = 18
luzaviso(236) = 0
luzaviso(237) = 0
luzaviso(238) = 0
luzaviso(239) = 0
luzaviso(240) = 0
luzaviso(241) = 0
luzaviso(242) = 0
luzaviso(243) = 18
luzaviso(244) = 22
luzaviso(245) = 24
luzaviso(246) = 38
luzaviso(247) = 30
luzaviso(248) = 38
luzaviso(249) = 0
luzaviso(250) = 0
luzaviso(251) = 0 'Capitán defensor de ciudad
luzaviso(252) = 0 'Capitán defensor de ciudad
luzaviso(253) = 0 'Capitán defensor de ciudad
luzaviso(254) = 0 'Capitán defensor de ciudad
luzaviso(255) = 0 'Capitán defensor de ciudad
luzaviso(256) = 0 'Capitán defensor de ciudad
luzaviso(257) = 0 'Capitán defensor de ciudad
luzaviso(258) = 0 'Capitán defensor de ciudad
luzaviso(259) = 0 'Capitán defensor de ciudad
luzaviso(260) = 0 'Capitán defensor de ciudad
luzaviso(261) = 16
luzaviso(262) = 14
luzaviso(263) = 18
luzaviso(264) = 14
luzaviso(265) = 0
luzaviso(266) = 12
luzaviso(267) = 16
luzaviso(268) = 50 ' Ettin
luzaviso(269) = 50 'Ettin
luzaviso(270) = 50 'Ettin
luzaviso(271) = 50 'Ettin
luzaviso(272) = 14
luzaviso(273) = 0
luzaviso(274) = 0
luzaviso(275) = 0
luzaviso(276) = 0
luzaviso(277) = 0
luzaviso(278) = 0
luzaviso(279) = 0
luzaviso(280) = 0
luzaviso(281) = 0
luzaviso(282) = 0
luzaviso(283) = 0
luzaviso(284) = 0
luzaviso(285) = 0
luzaviso(286) = 0
luzaviso(287) = 0
luzaviso(288) = 0
luzaviso(289) = 0
luzaviso(290) = 0
luzaviso(291) = 0
luzaviso(292) = 0
luzaviso(293) = 0
luzaviso(294) = 0
luzaviso(295) = 0
luzaviso(296) = 0
luzaviso(297) = 0
luzaviso(298) = 0
luzaviso(299) = 0
luzaviso(300) = 0

'-------------------------------



Call CargarGraficos

'quitar esto
'ChDrive App.Path
'ChDir App.Path
 
 web = GetVar(App.Path & "\Init\Web.dat", "WEB", "INIT")

'Cargamos el archivo de configuracion inicial
If FileExist(App.Path & "\init\Inicio.con", vbNormal) Then
    Config_Inicio = LeerGameIni()
End If


If FileExist(App.Path & "\init\ao.dat", vbNormal) Then
    Open App.Path & "\init\ao.dat" For Binary As #53
        Get #53, , RenderMod
    Close #53
    If (RenderMod.bUseVideo = False) Then
        If MsgBox("¿Desea usar la memoria de video?, en caso de que experimiente bajos FPS puede cambiar esta opción desdel AOSetup que se encuentra en la carpeta del juego.", vbExclamation + vbYesNo) = vbYes Then
            RenderMod.bUseVideo = True
            Open App.Path & "\init\ao.dat" For Binary As #53
                Put #53, , RenderMod
            Close #53
        End If
    End If

    Musica = IIf(RenderMod.bNoMusic = 1, 1, 0)
    Fx = IIf(RenderMod.bNoSound = 1, 1, 0)

    'RenderMod.iImageSize = 0
    Select Case RenderMod.iImageSize
        Case 4
            RenderMod.iImageSize = 0
        Case 3
            RenderMod.iImageSize = 1
        Case 2
            RenderMod.iImageSize = 2
        Case 1
            RenderMod.iImageSize = 3
        Case 0
            RenderMod.iImageSize = 4
    End Select
End If



tipf = Config_Inicio.tip

frmCargando.Show
frmCargando.Refresh

UserParalizado = False

'IPdelServidor = "92.43.20.27"
'IPdelServidor = "argendrag.no-ip.org"
'PuertoDelServidor = "7666"

AddtoRichTextBox frmCargando.status, "Iniciando constantes...", 0, 0, 0, 0, 0, 1

Call LeerLineaComandos


ReDim Ciudades(1 To NUMCIUDADES) As String
Ciudades(1) = "Ullathorpe"
Ciudades(2) = "Nix"
Ciudades(3) = "Banderbill"

ReDim CityDesc(1 To NUMCIUDADES) As String
CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"
ListaRazas(6) = "Orco"
ListaRazas(7) = "Vampiro"

ReDim ListaClases(1 To NUMCLASES) As String
ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Bandido"
ListaClases(9) = "Paladin"
ListaClases(10) = "Cazador"
ListaClases(11) = "Pescador"
ListaClases(12) = "Herrero"
ListaClases(13) = "Leñador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Pirata"
ListaClases(17) = "Ermitaño"
ListaClases(18) = "Arquero"
'pluto:2.3
ListaClases(19) = "Domador"
ReDim SkillsNames(1 To NUMSKILLS) As String
SkillsNames(1) = "Suerte"
SkillsNames(2) = "Aprendizaje de Magias"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con Armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Combate con Proyectiles"
SkillsNames(20) = "Golpeo con Armas Dobles"
SkillsNames(21) = "Navegacion"
SkillsNames(22) = "Daños en Magia"
SkillsNames(23) = "Defensa en Magias"
SkillsNames(24) = "Inmunidad a Magias"
SkillsNames(25) = "Daño en Armas"
SkillsNames(26) = "Defensa en Armas"
SkillsNames(27) = "Aprendizaje de Armas"
SkillsNames(28) = "Daño de Proyectiles"
SkillsNames(29) = "Defensa de Proyectiles"
SkillsNames(30) = "Aprendizaje de Proyectiles"
SkillsNames(31) = "Tactica Combate Proyectiles"
ReDim UserSkills(1 To NUMSKILLS) As Integer
ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"


frmOldPersonaje.NameTxt.Text = ""
frmOldPersonaje.PasswordTxt.Text = ""

'AddtoRichTextBox frmCargando.status, "Hecho", , , , 1
AddtoRichTextBox frmCargando.status, "AodraG v7.7", , , , 1
IniciarObjetosDirectX

AddtoRichTextBox frmCargando.status, "Cargando Sonidos....", 0, 0, 0, 0, 0, 1
AddtoRichTextBox frmCargando.status, "Hecho", , , , 1

Dim loopc As Integer

LastTime = GetTickCount

ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)

'Call InitTileEngine(frmMain.hWnd, 152, 7, 32, 32, 13, 17, 9)
Call InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top + 1, frmMain.MainViewShp.Left + 1, 32, 32, 13, 17, 9)
Call AddtoRichTextBox(frmCargando.status, "Creando animaciones...")

Call CargarAnimsExtra
'Call CargarTips
UserMap = 1
Call CargarArrayLluvia
Call CargarAnimArmas
Call CargarAnimEscudos

'pluto:6.0A
Call InitFonts

AddtoRichTextBox frmCargando.status, "¡Bienvenido al Mundo AodraG!", , , , 1

'Call DrawGrhtoHdc(frmCargando.Picture1.hWnd, frmCargando.Picture1.hdc, 18014, SR, DR)
'Call DrawGrhtoHdc(frmCargando.Picture2.hWnd, frmCargando.Picture2.hdc, 3040, SR, DR)
'Call DDrawTransGrhIndextoSurface(frmCargando.Picture1.Picture.hpal, 18014, 100, 100, 0, 0)

Sleep 3000

Unload frmCargando

'quitar esto
'LoopMidi = True

'pluto:6.3--------------
If navida = 1 Then
Config_Inicio.DirMusica = "Midi\Navidad"
Else
Config_Inicio.DirMusica = "Midi\"
End If
'----------------
Call audio.Initialize(DirectX, frmMain.hWnd, App.Path & "\" & Config_Inicio.DirSonidos & "\", App.Path & "\" & Config_Inicio.DirMusica & "\")
    'Enable / Disable audio
If Musi = 1 Then audio.MusicActivated = True Else audio.MusicActivated = False
If Son = 1 Then audio.SoundActivated = True Else audio.SoundActivated = False



'Call PlayMp3("d:\ole.mp3")
'If Musica = 0 Then
'Call CargarMIDI(DirMidi & MIdi_Inicio & ".mid")
'Play_Midi
'End If
 Call audio.PlayMIDI(MIdi_Inicio & ".mid")
'Call audio.PlayMIDI("music.mp3")
          



frmPres.Picture = LoadPicture(App.Path & "\Graficos\LOGODRAG.BMP")

'frmPres.WindowState = vbMaximized

'pluto:2.18
'Dim variable As String
'Dim ie As Object
'variable = "http://www.juegosdrag.es/aodragx.htm"
'ShellExecute frmCargando.hWnd, "open", "http://www.aodrag.com", vbNullString, vbNullString, conSwNormal
'Set ie = CreateObject("InternetExplorer.Application")
'ie.Visible = True
'ie.Navigate variable
'------

'frmPres.Show
'pluto:6.4
'frmPres.Navegador.Navigate "http://www.servilink.com.ar/aodrag.jpg"
'Dim Html As String
'Html = "about:" & _
'"<html>" & _
'"<body leftMargin=0 topMargin=0 marginheight=0 marginwidth=0 scroll=no>" & _
'"<img src=http://www.juegosdrag.es/into2.jpg width= 800 height= 600 ></img></body></html>"
'frmPres.Navegador.Navigate Html



'Do While Not finpres
'    DoEvents
'Loop

'Unload frmPres

Naci = Val(GetVar(App.Path & "\Init\Web.dat", "WEB", "NACI"))
If Naci = 0 Then
frmNaci.Show
Do While Naci = 0
DoEvents
Loop
Call WriteVar(App.Path & "\Init\Web.dat", "WEB", "NACI", Naci)
End If


frmConnect.Visible = True

frmMain.Socket1.HostName = CurServerIp
frmMain.Socket1.RemotePort = CurServerPort

'Loop principal!
'[CODE]:MatuX'
    MainViewRect.Left = MainViewLeft + 32 * RenderMod.iImageSize
    MainViewRect.Top = MainViewTop + 32 * RenderMod.iImageSize
    MainViewRect.Right = (MainViewRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainViewRect.Bottom = (MainViewRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

    MainDestRect.Left = ((TilePixelWidth * TileBufferSize) - TilePixelWidth) + 32 * RenderMod.iImageSize
    MainDestRect.Top = ((TilePixelHeight * TileBufferSize) - TilePixelHeight) + 32 * RenderMod.iImageSize
    MainDestRect.Right = (MainDestRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainDestRect.Bottom = (MainDestRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

    Dim OffsetCounterX As Integer
    Dim OffsetCounterY As Integer
'[END]'



PrimeraVez = True
prgRun = True
pausa = False
bInvMod = True
lFrameLimiter = DirectX.TickCount
'[CODE 001]:MatuX'
lFrameModLimiter = 60

'[END]'
Do While prgRun

    If RequestPosTimer > 0 Then
        RequestPosTimer = RequestPosTimer - 1
        If RequestPosTimer = 0 Then
            'Pedimos que nos envie la posicion
            Call SendData("RPY")
        End If
    End If

    Call RefreshAllChars
    
    Dim vete As Byte

    '[CODE 001]:MatuX
    '
    '   EngineRun
    If EngineRun Then
        '[DO]:Dibuja el siguiente frame'
        '[CODE 000]:MatuX'
        'If frmMain.WindowState <> 1 And CurMap > 0 And EngineRun Then
        If frmMain.WindowState <> 1 Then
        '[END]'
            'Call ShowNextFrame(frmMain.Top, frmMain.Left)
            '****** Move screen Left, Right, Up and Down if needed ******
            If AddtoUserPos.X <> 0 Then
            If FPSFast = True Then
                OffsetCounterX = (OffsetCounterX - (2 * Sgn(AddtoUserPos.X)))
                Else
                OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.X)))
                End If
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
            ElseIf AddtoUserPos.Y <> 0 Then
            If FPSFast = True Then
            
                OffsetCounterY = OffsetCounterY - (2 * Sgn(AddtoUserPos.Y))
                Else
                OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.Y))
                End If
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If
    
            '****** Update screen ******
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            'Call DoNightFX
            'Call DoLightFogata(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
            '[CODE 000]:MatuX
                'Call MostrarFlags
                'If IScombate Then Call Dialogos.DrawText(260, 260, "MODO COMBATE", vbRed, 0)
                'pluto:6.9
                'If FramesPerSec > 20 Then Sleep 3000
                frmMain.Label3.Caption = CurMap
                frmMain.Label10.Caption = "X:" & UserPos.X
                frmMain.Label11.Caption = "Y:" & UserPos.Y

                If FPSFLAG Then
                    Call Dialogos.DrawText(735, 260, "FPS: " & FramesPerSec, vbWhite, 0)
                    Call Dialogos.DrawText(735, 270, "Map: " & CurMap, vbWhite, 0)
                    Call Dialogos.DrawText(735, 280, "X: " & UserPos.X & " Y: " & UserPos.Y, vbWhite, 0)
                End If
'pluto:6.5
'-----------------------------------------------
'If FramesPerSec = 1 Then
'If logged = True Then SendData ("B2")
'frmMain.Cheat.Enabled = False
Dim Nus As Integer
'Nus = Val(GetVar(App.Path & "\Init\Update.ini", "FICHERO", "z"))
'Nus = Nus + 1
'Call WriteVar(App.Path & "\Init\Update.ini", "FICHERO", "z", Val(Nus))
'MsgBox "Acelerador Detectado!!. Ha quedado registrado este intento de usar un Cheat, si vemos que vuelves intentarlo serán baneados todos tus personajes. ¡¡ ESTAS AVISADO !!"
'End
'End If
'----------------------------------------------
                
   
    
    
                'pluto:6.0A
                If ECiudad Then
                Call Dialogos.DrawText(260, 260, EstadoCiudad, vbYellow, 0)
                End If
                                               
                If PYFLAG Then
                    Call Dialogos.DrawText(260, 665 - (10 * (Party.numMiembros + 1)), "Miembros: " & Party.numMiembros, vbWhite, 0)
                    Dim cnmi As Integer
                    Dim cco As Long
                    For cnmi = 1 To Party.numMiembros
                    'pluto:6.0A
                    If CharList(Party.Miembros(cnmi).Index).Muerto = False Then cco = vbWhite Else cco = vbRed
                    
                    If Party.Miembros(cnmi).X = 0 Or Party.Miembros(cnmi).Y = 0 Then
                    Call Dialogos.DrawText(260, 665 - (10 * cnmi), Party.Miembros(cnmi).Nombre & " (" & Party.Miembros(cnmi).privi & "%)", cco, 0)
                    Else
                    Call Dialogos.DrawText(260, 665 - (10 * cnmi), Party.Miembros(cnmi).Nombre & " (" & Party.Miembros(cnmi).privi & "%) X: " & Party.Miembros(cnmi).X & " Y: " & Party.Miembros(cnmi).Y, cco, 0)
                    End If
                    Next
                  
                End If
                'pluto:2.9.0
                If CurMap = 192 Then
                Call Dialogos.DrawText(445, 260, "Local..........: " & Goleslocal, vbWhite, 0)
                Call Dialogos.DrawText(445, 275, "Visitante...: " & Golesvisitante, vbWhite, 0)
                End If
                'pluto:2.12
                If CurMap = 194 Then
                Call Dialogos.DrawText(445, 260, "El Mejor Luchador es " & UserTorneo2 & " con " & RecordTorneo2 & " Victorias", vbWhite, 0)
                Call Dialogos.DrawText(495, 275, "Bote Acumulado: " & BoteTorneo2, vbWhite, 0)
                Call Dialogos.DrawText(445, 290, "El Bote es para quien consiga 10 victorias consecutivas.", vbWhite, 0)

                End If
                
                
                If Dialogos.CantidadDialogos <> 0 Then Call Dialogos.MostrarTexto
                If Cartel Then Call DibujarCartel
                If bInvMod Then DibujarInv
    
                Call DrawBackBufferSurface
                
                Call RenderSounds
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
    End If
    'FPSFast = True

        If Not pausa And frmMain.Visible And Not frmForo.Visible Then
            CheckKeys
            LastTime = GetTickCount

    End If
    
    'If Musica = 0 Then
        'If Not SegState Is Nothing Then
            'If Not Perf.IsPlaying(Seg, SegState) Then Play_Midi
        'End If
   ' End If
         'Musica = 0
    'End If
    '[END]'
    
    '[CODE 001]:MatuX
    ' Frame Limiter
        'FramesPerSec = FramesPerSec + 1
        If DirectX.TickCount - lFrameTimer > 1000 Then
            FramesPerSec = FramesPerSecCounter
            FramesPerSecCounter = 0
            lFrameTimer = DirectX.TickCount
'pluto:7.0-------------------
            If DopeEstimulo = 1 Then
            frmMain.Estimulo.Visible = Not frmMain.Estimulo.Visible
            Else
            frmMain.Estimulo.Visible = False
            End If
            
            If Fnorte = 1 Then
            frmMain.Norte.Visible = Not frmMain.Norte.Visible
            Else
            frmMain.Norte.Visible = False
            End If
            
            If Fsur = 1 Then
            frmMain.Sur.Visible = Not frmMain.Sur.Visible
            Else
            frmMain.Sur.Visible = False
            End If
            
            If Feste = 1 Then
            frmMain.Este.Visible = Not frmMain.Este.Visible
            Else
            frmMain.Este.Visible = False
            End If
            
            If Foeste = 1 Then
            frmMain.Oeste.Visible = Not frmMain.Oeste.Visible
            Else
            frmMain.Oeste.Visible = False
            End If
            
            If Ffortaleza = 1 Then
            frmMain.Fortaleza.Visible = Not frmMain.Fortaleza.Visible
            Else
            frmMain.Fortaleza.Visible = False
            End If
            'luz aviso dificultad mapa
            Dim laviso As String
            Dim cluz As Byte
            cluz = cluz + 1
            If cluz = 4 Then cluz = 0
            
            Select Case Luzaviso2
            Case 1
            laviso = "verde-" & cluz & ".jpg"
            frmMain.luzaviso.Picture = LoadPicture(App.Path & "\Graficos\" & laviso)
            Case 2
            laviso = "amarillo-" & cluz & ".jpg"
            frmMain.luzaviso.Picture = LoadPicture(App.Path & "\Graficos\" & laviso)
            Case 3
            laviso = "rojo-" & cluz & ".jpg"
            frmMain.luzaviso.Picture = LoadPicture(App.Path & "\Graficos\" & laviso)

            End Select
            
'------------------------------------
        End If
        
        'While DirectX.TickCount - lFrameLimiter < lFrameModLimiter: Wend
 
        Dim Vel As Integer
        If FPSFast = True Then
        Vel = 15
        Else
        Vel = 60
        End If
        While (GetTickCount - lFrameTimer) \ Vel < FramesPerSecCounter
            Sleep 5
        Wend
        
        'lFrameLimiter = DirectX.TickCount
  
  'pluto:2.4.2
  'Sistema de timers renovado:
esttick = GetTickCount
'pluto:6.0A
If NoPuedeMagia = True And vete = 0 Then
timers(4) = 0
vete = 1
End If

For loopc = 1 To UBound(timers)
timers(loopc) = timers(loopc) + (esttick - ulttick)

 'timer de trabajo
If timers(1) >= tUs Then
timers(1) = 0
 NoPuedeUsar = False
 End If
 'timer de attaque (77)
If timers(2) >= tAt Then
 timers(2) = 0
 UserCanAttack = 1
End If
'pluto:2.4.5
If timers(3) >= tTr Then
timers(3) = 0
NoPuedeTirar = False
End If
If timers(4) >= tMg Then
timers(4) = 0
NoPuedeMagia = False
vete = 0
End If
If timers(6) >= 5000 Then
timers(6) = 0
        'If IndiceLabel <> -1 Then
        'Call frmMain.ReestablecerLabel
        'IndiceLabel = -1
        'Call Eliminar_ToolTip
    'End If
End If

Next loopc
ulttick = GetTickCount



'pluto:2.15
If UCase$(UserClase) = "ARQUERO" Or UCase$(UserClase) = "CAZADOR" Then
flechudo = (Val(frmMain.LvlLbl) * 20)
'ElseIf UCase$(UserClase) = "CAZADOR" Then
'flechudo = (Val(frmMain.LvlLbl) * 10)
Else
flechudo = 0
End If

If flechudo > 800 Then flechudo = 800
If timers(5) >= tFle - flechudo Then  'tFle Then
timers(5) = 0
NoPuedeFlechas = False
End If
'pluto:2.9.0
Dim viz As Byte
If FrmGol.Visible = True And viz = 0 Then timers(3) = 0: viz = 1
If timers(3) = 3000 And viz = 1 Then FrmGol.Visible = False: viz = 0
'pluto:2.15
Dim vaz As Byte
If frmMain.Label4.Visible = True And vaz = 0 Then timers(3) = 0: vaz = 1
If timers(3) = 3000 And vaz = 1 Then frmMain.Label4.Visible = False: vaz = 0


    '[END]'
 'Delzak)
'If TeclasModificadas = False Then Call CustomKeys.LoadDefaults
          
                    
                                    DoEvents
Loop

EngineRun = False
frmCargando.Show
AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
LiberarObjetosDX


If bNoResChange = True Then
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


Call UnloadAllForms

Config_Inicio.tip = tipf
Call EscribirGameIni(Config_Inicio)

End

ManejadorErrores:
    LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End
    
End Sub



Sub WriteVar(file As String, Main As String, Var As String, Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, Value, file

End Sub

Function GetVar(file As String, Main As String, Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim L As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file

GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)

End Function


'[CODE 002]:MatuX
'
'  Función para chequear el email
'
Public Function CheckMailString(ByRef sString As String) As Boolean
        On Error GoTo errHnd:
        Dim lPos  As Long, lX    As Long
        Dim iAsc  As Integer
    
        '1er test: Busca un simbolo @
        lPos = InStr(sString, "@")
        If (lPos <> 0) Then
            '2do test: Busca un simbolo . después de @ + 1
            If Not (IIf((InStr(lPos, sString, ".", vbBinaryCompare) > (lPos + 1)), True, False)) Then _
                Exit Function
    
            '3er test: Valída el ultimo caracter
            If Not (CMSValidateChar_(Asc(Right(sString, 1)))) Then _
                Exit Function
    
            '4to test: Recorre todos los caracteres y los valída
            For lX = 0 To Len(sString) - 1 'el ultimo no porque ya lo probamos
                If Not (lX = (lPos - 1)) Then
                    iAsc = Asc(Mid(sString, (lX + 1), 1))
                    If Not (iAsc = 46 And lX > (lPos - 1)) Then _
                        If Not CMSValidateChar_(iAsc) Then _
                            Exit Function
                End If
            Next lX
    
            'Finale
            CheckMailString = True
        End If
    
errHnd:
        'Error Handle
    End Function
    
Private Function CMSValidateChar_(ByRef iAsc As Integer) As Boolean
'pluto:6.9 añade 65 y 95
CMSValidateChar_ = IIf( _
                    (iAsc >= 45 And iAsc <= 57) Or _
                    (iAsc >= 65 And iAsc <= 90) Or _
                    (iAsc >= 97 And iAsc <= 122) Or _
                    (iAsc = 95), True, False)
End Function



Function HayAgua(X As Integer, Y As Integer) As Boolean

If MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
   MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
   MapData(X, Y).Graphic(2).GrhIndex = 0 Then
            HayAgua = True
Else
            HayAgua = False
End If

End Function
    Public Sub ShowSendTxt()
        If Not frmCantidad.Visible Then
            frmMain.SendTxt.Visible = True
            frmMain.SendTxt.SetFocus
        End If
    End Sub

Public Function porcentaje(ByVal total As Long, ByVal Porc As Long) As Long
On Error GoTo Fallo
porcentaje = (total * Porc) / 100

Exit Function
Fallo:
Call LogError("porcentaje " & Err.Number & " D: " & Err.Description)
End Function
'Delzak sos offline
'Cuando abre desde GM
'Public Sub RellenarDatos(datos As String)

'frmContestarSos.Label6.Caption = ReadField(1, datos, Asc(";"))
'frmContestarSos.Label2.Caption = UserName
'frmContestarSos.Text1 = ReadField(2, datos, Asc(";"))
'frmContestarSos.Label3 = ReadField(1, ReadField(3, datos, Asc(";")), Asc(" ")) & " " & ReadField(3, ReadField(3, datos, Asc(";")), Asc(" "))
'frmContestarSos.Label1 = ""
'frmContestarSos.List1.Visible = False

'End Sub

'Public Sub BuscaMensaje(numero As String)

'numero = Right$(numero, Len(numero) - 8)

'Call SendData("/DAME" & numero)
'End Sub
