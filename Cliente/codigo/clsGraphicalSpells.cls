VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalSpells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private SpellOffset As Long             'Number of lines we scrolled down from topmost
Public SpellSelectedItem As Long              'Currently selected item
Private TempItemSelect As Long

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer

Private slotDragged As Integer
 
Public Event dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)

Private ShowValue As Boolean
Private ImgContainer As Boolean
Private ShowInfo As Boolean

Private WithEvents SpellWindow As PictureBox    'Placeholder where to render the inventory
Attribute SpellWindow.VB_VarHelpID = -1

Public Sub DrawSpells()
    Dim i As Byte
    Dim X As Integer
    Dim Y As Integer
    
    For i = SpellOffset * (SpellWindow.ScaleWidth \ 43) + 1 To UBound(UserHechizos)
        If UserHechizos(i).GrhIndex Then
            X = ((i - 1) Mod (SpellWindow.Width / 43)) * 43
            Y = ((i - 1) \ (SpellWindow.Width / 43)) * 43

            
            'Dibujamos un cuadradito en el item seleccionado
            If SpellSelectedItem = i Then
                'Dibujamos el Grh del item iluminado
                Call Engine.Draw_GrhIndex(UserHechizos(i).GrhIndex, X, Y, LightIluminado())
            Else
                'Dibujamos el Grh del item desiluminado
                Call Engine.Draw_GrhIndex(UserHechizos(i).GrhIndex, X, Y, LightOscurito())
            End If
            
        End If
    Next i
End Sub

Public Sub Initialize(ByRef Picture As PictureBox, _
                            Optional ByVal TileWidth As Integer = 43, _
                            Optional ByVal TileHeight As Integer = 43, Optional ByVal startX As Integer = 0, _
                            Optional ByVal startY As Integer = 0)

    Set SpellWindow = Picture
    
    SpellWindow.AutoRedraw = True
    
    SpellWindow.ScaleMode = 3

    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
End Sub

Private Sub SpellWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer

    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > SpellWindow.Width Or Y > SpellWindow.Height Then Exit Sub
       
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
       
    TempItemSelect = ClickItem(CInt(X), CInt(Y))
    
    If Button = vbRightButton Then
        TempItem = temp_x + (temp_y + SpellOffset) * (SpellWindow.ScaleWidth \ ownTilePixelWidth) + 1 - temp_y
        
        If (temp_y > 0) Then
            TempItem = TempItem + (temp_y) ' Corrección de filas
        End If
        
        If TempItem <= UBound(UserHechizos) Then
            slotDragged = TempItem
        End If
        
    End If
End Sub

Private Sub SpellWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Exit Sub
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer

    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > SpellWindow.Width Or Y > SpellWindow.Height Then Exit Sub
       
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + SpellOffset) * (SpellWindow.ScaleWidth \ ownTilePixelWidth) + 1 - temp_y
    
    If (temp_y > 0) Then
        TempItem = TempItem + (temp_y) ' Corrección de filas
    End If
    
    If TempItem <= UBound(UserHechizos) Then
        frmMain.lblItem.Caption = UserHechizos(TempItem).name ' Que nos diga el nombre del objeto cuando pasamos el mouse sobre él
        'SpellWindow.ToolTipText = UserHechizos(TempItem).Name ' Que nos diga el nombre del hechizo cuando pasamos el mouse sobre él
    End If
    
    Engine.DrawSpells
End Sub


Private Sub SpellWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim prevSelItem As Long

    'Exit if it got outside the control's area
    If X < 0 Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        X = 0
    ElseIf Y < 0 Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        Y = 0
    ElseIf X > SpellWindow.Width Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        X = SpellWindow.Width
    ElseIf Y > SpellWindow.Height Then
        If frmMain.UsandoDrag Then frmMain.UsabaDrag = True
        Y = SpellWindow.Height
    End If
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > SpellWindow.Width Or Y > SpellWindow.Height Then _
        Exit Sub
        
    prevSelItem = SpellSelectedItem
    SpellSelectedItem = ClickItem(CInt(X), CInt(Y))
    
    If Button = 2 Then ' GSZAO - info del hechizo con segundo click
        Call WriteSpellInfo(SpellSelectedItem)
        ShowInfo = True
    End If
    
    If Button = vbRightButton Then
        If SpellSelectedItem <> slotDragged Then
            Call frmMain.picSpell_dragDone(slotDragged, TempItemSelect)
            frmMain.UsandoDrag = False
        Else
            frmMain.UsandoDrag = False
        End If
    End If

    Engine.DrawSpells
    
End Sub

Private Sub SpellWindow_Click()
Dim tX As Byte
Dim tY As Byte

    If TempItemSelect > MAX_SPELL_SLOTS Then Exit Sub
    
    If UserHechizos(TempItemSelect).name <> "(None)" And MainTimer.Check(TimersIndex.Work, False) And UsaMacro = False Then
        If ShowInfo = False Then
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
            Else
                Call WriteCastSpell(SpellSelectedItem)
                
                If UserHechizos(TempItemSelect).name = "Acechar" Then
                    Call ConvertCPtoTP(366, 256, tX, tY)
                    Call WriteWorkLeftClick(tX, tY, eSkill.magia)
                Else
                    Call WriteWork(eSkill.magia)
                End If
                
                UsaMacro = True
            End If
        Else
            ShowInfo = False
        End If
        frmMain.lblItem.Caption = UserHechizos(SpellSelectedItem).name
    End If
    
    Engine.DrawSpells
            
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long

    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItemSelect = temp_x + (temp_y + SpellOffset) * (SpellWindow.ScaleWidth \ ownTilePixelWidth) + 1 - temp_y

    If (temp_y > 0) Then
        TempItemSelect = TempItemSelect + (temp_y) ' Corrección de filas
    End If

    If TempItemSelect <= UBound(UserHechizos) Then
    
        If UserHechizos(TempItemSelect).name <> "(None)" Then
            ClickItem = TempItemSelect
        Else
            ClickItem = 0
            frmMain.lblItem.Caption = vbNullString
        End If
        
    End If
    
    Engine.DrawSpells
        
End Function

Private Function LowMana() As Boolean
   ' Function para hacer el if cuando tenga poca mana se ponga rojo jaja.
End Function

Public Property Get SelectedItem() As Long
'***************************************************
'Author: ^[GS]^
'Last Modify Date: 09/07/2012 - ^[GS]^
'***************************************************
    SelectedItem = SpellSelectedItem
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Long
'***************************************************
'Author: ^[GS]^
'Last Modify Date: 09/07/2012 - ^[GS]^
'***************************************************
    GrhIndex = UserHechizos(slot).GrhIndex
End Property


