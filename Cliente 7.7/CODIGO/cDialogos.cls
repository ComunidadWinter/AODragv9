VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDialogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_DIALOGOS = 10000
Private Const MAXLONG = 15

Private Type tDialogo
    Formato()  As String
    Formateado As Boolean
    texto     As String
    Delay     As Integer
    X         As Integer
    Y         As Integer
    vida      As Integer
    UserIndex As Integer
    Color     As Long
End Type

'Vector que contiene los DIALOGOS
Private Dialogos(1 To MAX_DIALOGOS) As tDialogo
'Tiempo de vida de cada mensaje
Private Delay As Integer
'Apunta a el ultimo mensaje
Public UltimoDialogo As Integer
'Contiene la cantidad de mensajes activos
Public CantidadDialogos As Integer

Dim TextosSurf As New Dictionary


Private Function PrimerIndiceVacio() As Integer
On Error Resume Next
Dim i As Integer
Dim flag As Boolean

i = 1

Do While i <= MAX_DIALOGOS And Not flag
    flag = (Dialogos(i).vida = 0)
    If Not flag Then i = i + 1
Loop

If flag Then
    PrimerIndiceVacio = i
Else
    PrimerIndiceVacio = 1
End If

End Function


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?Formatea la cadena S¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer, Dialgi As Integer)
On Error Resume Next
If anti + i <= Len(s) + 1 Then
    If ((i >= MAXLONG) And Mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
        Dialogos(Dialgi).Formato(k) = Mid(s, anti, i + 1)
        k = k + 1
        anti = anti + i + 1
        i = 0
    Else
        i = i + 1
    End If
    Call DarFormato(s, i, k, anti, Dialgi)
End If

End Function

Private Function HayEspacio(s As String) As Boolean
On Error Resume Next
If InStr(1, s, " ") <> 0 Then
    HayEspacio = True
Else
    HayEspacio = False
End If
End Function

Private Sub DarleFormatoTXT(Dialgi As Integer)
On Error Resume Next
 'Le damos formato al txt
  If (Len(Dialogos(Dialgi).texto) > MAXLONG) And HayEspacio(Dialogos(Dialgi).texto) Then
                Dim j As Integer
                ReDim Dialogos(Dialgi).Formato(0 To (Len(Dialogos(Dialgi).texto) \ (MAXLONG \ 2)))
                Dim i As Integer, k As Integer, anti As Integer
                anti = 1
                k = 0
                i = 0
                Call DarFormato(Dialogos(Dialgi).texto, i, k, anti, Dialgi)
                i = 0
                '[MerLiNz:ANTICUELGES]
                'Do While Dialogos(Dialgi).Formato(i) <> ""
                Do While Dialogos(Dialgi).Formato(i) <> "" And i < UBound(Dialogos(Dialgi).Formato)
                '[\END]
                    i = i + 1
                Loop
                ReDim Preserve Dialogos(Dialgi).Formato(0 To i - 1)
                Dialogos(Dialgi).Formateado = True
  End If
End Sub
Public Sub CrearDialogo(txt As String, ByVal User As Integer, xColor As Long)
On Error Resume Next
Dim MiUserIndex As Integer
Dim IndiceLibre As Integer
Dim Color As Long

Select Case xColor
    Case 1
        Color = vbWhite
    Case 2
        Color = vbRed
    Case 3
        Color = &HFFC0FF
    Case 4
        Color = &H196FF
    Case 5
        Color = vbYellow
    Case 6
        Color = &HFF9600
    Case 7
        Color = vbCyan
    Case 8
        Color = vbGreen
End Select

If BuscarUserIndex(User, MiUserIndex) Then
    Dialogos(MiUserIndex).vida = 0
    Dialogos(MiUserIndex).Color = Color
    Dialogos(MiUserIndex).Formateado = False
    Dialogos(MiUserIndex).texto = txt
    Dialogos(MiUserIndex).Delay = Delay
    Dialogos(MiUserIndex).UserIndex = User
    Dialogos(MiUserIndex).vida = Delay
    DarleFormatoTXT MiUserIndex
Else
    IndiceLibre = PrimerIndiceVacio
    Dialogos(IndiceLibre).Color = Color
    Dialogos(IndiceLibre).Formateado = False
    Dialogos(IndiceLibre).texto = txt
    Dialogos(IndiceLibre).Delay = Delay
    Dialogos(IndiceLibre).UserIndex = User
    Dialogos(IndiceLibre).vida = Delay
    DarleFormatoTXT IndiceLibre
    If UltimoDialogo > IndiceLibre Then
        UltimoDialogo = IndiceLibre
    End If
End If
    
CantidadDialogos = CantidadDialogos + 1
  'pluto:28-3-04
  If CantidadDialogos = 9999 Then CantidadDialogos = 1
End Sub
Private Function BuscarUserIndex(User As Integer, MiUser As Integer) As Boolean
On Error Resume Next
If CantidadDialogos > 0 Then
    Dim i As Integer
    Dim flag As Boolean
    i = 1
    
    Do While i <= MAX_DIALOGOS And Not flag
           flag = (Dialogos(i).UserIndex = User)
           If Not flag Then i = i + 1
    Loop
    
    If flag Then
            MiUser = i
            BuscarUserIndex = True
    Else
            BuscarUserIndex = False
    End If
End If

End Function

Public Sub Update_Dialog_Pos(X As Integer, Y As Integer, Index As Integer)
On Error Resume Next
Dim MiUserIndex As Integer
If BuscarUserIndex(Index, MiUserIndex) Then
  If Dialogos(MiUserIndex).vida > 0 Then
        If Dialogos(MiUserIndex).Formateado Then
            Dialogos(MiUserIndex).X = X - (MAXLONG \ 2) * 3
            Dialogos(MiUserIndex).Y = Y - (UBound(Dialogos(MiUserIndex).Formato) * 3)
        Else
            Dialogos(MiUserIndex).X = X - (Len(Dialogos(MiUserIndex).texto) \ 2) * 3
            Dialogos(MiUserIndex).Y = Y
        End If
        If Dialogos(MiUserIndex).vida > 0 Then
           Dialogos(MiUserIndex).vida = Dialogos(MiUserIndex).vida - 1
        End If
        If Dialogos(MiUserIndex).vida < 1 Then
            If MiUserIndex = UltimoDialogo Then
                
                Do While Dialogos(UltimoDialogo).vida = 0
                    UltimoDialogo = UltimoDialogo - 1
                    If UltimoDialogo < 1 Then
                        UltimoDialogo = 0
                        CantidadDialogos = 0
                        Exit Sub
                    End If
                Loop
                
            End If
             
            CantidadDialogos = CantidadDialogos - 1
        End If
  End If
End If
End Sub
Public Sub DrawText(lngXPos As Integer, lngYPos As Integer, strText As String, lngColor As Long, fxc As Integer)
On Error Resume Next
    'Set Ifnt = font
If strText <> "" Then
   Dim font As New StdFont
    'Dim Ifnt As IFont
    
    
    Dim Surf As DirectDrawSurface7
        Dim ddRect As RECT
            With ddRect
            .Left = 0
            .Top = 0
            .Right = 300
            .Bottom = 300
        End With
    If Not TextosSurf.Exists(strText & lngColor) Then
    
    
    
        Dim ddck        As DDCOLORKEY
        Dim SurfaceDesc As DDSURFACEDESC2
        

        With SurfaceDesc
            .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
            If RenderMod.bUseVideo Then
                .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN 'Or DDSCAPS_3DDEVICE Or DDSCAPS_ALPHA
            Else
                .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY 'Or DDSCAPS_3DDEVICE
            End If
            .lHeight = ddRect.Bottom
            .lWidth = ddRect.Right
        End With

    
        Set Surf = DirectDraw.CreateSurface(SurfaceDesc)
        ddck.low = 0
        ddck.high = 0
        Call Surf.BltColorFill(ddRect, 0)
        Surf.SetColorKey DDCKEY_SRCBLT, ddck
        
        Surf.SetFontTransparency True
        Surf.SetForeColor RGB(20, 20, 20)
        Surf.SetFont frmMain.font 'lfnt
        Surf.DrawText 0, 0, strText, False
        
        Surf.SetForeColor lngColor
        Surf.DrawText 2, 1, strText, False
        
        
        Call TextosSurf.Add(strText & lngColor, Surf)
    Else
        Set Surf = TextosSurf(strText & lngColor)
    End If
    
    
    BackBufferSurface.BltFast lngXPos - 2, lngYPos - 1, Surf, ddRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    If TextosSurf.Count > 600 Then
        Call TextosSurf.RemoveAll
    End If
       
End If
   
End Sub

Public Sub MostrarTexto()
On Error Resume Next
Dim i As Integer
Dim bodito As Integer
For i = 1 To CantidadDialogos
    If Dialogos(i).vida > 0 Then
        If Dialogos(i).Formateado Then
            Dim j As Integer, desp As Integer
       'pluto:6.0A
       If CharList(Dialogos(i).UserIndex).Body.Walk(1).GrhIndex = 0 Then
       bodito = 70
       Else
       bodito = 0
       End If

        
            desp = (-frmMain.font.Size) * UBound(Dialogos(i).Formato) - 5
            For j = 0 To UBound(Dialogos(i).Formato)
             DrawText Dialogos(i).X, Dialogos(i).Y + desp - bodito, Dialogos(i).Formato(j), Dialogos(i).Color, 0
             desp = desp + (frmMain.font.Size) + 5
            Next
         Else
        'If CharList(Dialogos(i).UserIndex).Body.HeadOffset.y = -103 Then bodito = 25 Else bodito = 0
           DrawText Dialogos(i).X, Dialogos(i).Y + bodito, Dialogos(i).texto, Dialogos(i).Color, 0
         End If
    Else
      Dialogos(i).texto = ""
      Dialogos(i).UserIndex = 0
      Dialogos(i).Formateado = False
      Dialogos(i).Delay = 0
    End If
  
Next
End Sub

Public Sub QuitarDialogo(ByVal CharIndex As Integer)
On Error Resume Next
Dim i As Integer

If BuscarUserIndex(CharIndex, i) Then
      Dialogos(i).vida = 0
      Dialogos(i).texto = ""
      Dialogos(i).UserIndex = 0
      Dialogos(i).Formateado = False
      Dialogos(i).Delay = 0
      CantidadDialogos = CantidadDialogos - 1
      
      If i = UltimoDialogo Then
        Do While Dialogos(i).vida = 0
                i = i - 1
                If i < 1 Then
                    UltimoDialogo = 0
                    CantidadDialogos = 0
                    Exit Sub
                End If
        Loop
      End If
End If



End Sub

Public Sub BorrarDialogos()

On Error Resume Next

Dim i As Integer

For i = 1 To MAX_DIALOGOS
      Dialogos(i).vida = 0
      Dialogos(i).texto = ""
      Dialogos(i).UserIndex = 0
      Dialogos(i).Formateado = False
      Dialogos(i).Delay = 0
Next i

End Sub

Private Sub Class_Initialize()

Delay = 300
UltimoDialogo = 0
CantidadDialogos = 0

End Sub
