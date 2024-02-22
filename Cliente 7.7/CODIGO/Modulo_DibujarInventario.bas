Attribute VB_Name = "DibujarInventario"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

'[CODE]:MatuX
'
'  Casi todo recodeado menos los calculos
'
'[END]'

Public Const XCantItems = 5

Public OffsetDelInv As Integer
Public ItemElegido As Integer
Public mx As Integer
Public my As Integer

Public AuxSurface   As DirectDrawSurface7
Private BoxSurface   As DirectDrawSurface7
Private SelSurface   As DirectDrawSurface7
Private bStaticInit  As Boolean   'Se inicializaron las Statics?
Private r1           As RECT, r2 As RECT, auxr As RECT
Private rBox         As RECT  'Pos del cuadradito rojo
Private rBoxFrame(2) As RECT
Private iFrameMod    As Integer


Function ClicEnItemElegido(X As Integer, Y As Integer) As Boolean


mx = X \ 32 + 1
my = Y \ 32 + 1
If my < 1 Then Exit Function
bInvMod = True
If ItemElegido = 0 Or FLAGORO Then
    ClicEnItemElegido = False
Else
    ClicEnItemElegido = (UserInventory(ItemElegido).OBJIndex > 0) And (ItemElegido = (mx + (my - 1) * 5) + OffsetDelInv)
End If
End Function

Sub ItemClick(X As Integer, Y As Integer)
Dim lPreItem As Long

bInvMod = False
mx = X \ 32 + 1
my = Y \ 32 + 1

lPreItem = (mx + (my - 1) * 5) + OffsetDelInv

If lPreItem <= MAX_INVENTORY_SLOTS Then _
If UserInventory(lPreItem).GrhIndex > 0 Then _
    ItemElegido = lPreItem: bInvMod = True
End Sub

Public Sub DibujarInvBox()
    On Error Resume Next
    If bStaticInit And ItemElegido <> 0 Then
        'Call BoxSurface.BltColorFill(auxr, vbBlack)
        Call BoxSurface.BltFast(0, 0, SelSurface, auxr, DDBLTFAST_SRCCOLORKEY)
        
        With Grh(1)
            .FrameCounter = 2
             Call BoxSurface.BltFast(0, 0, SurfaceDB.GetBMP(GrhData(GrhData(.GrhIndex).Frames(.FrameCounter)).FileNum), rBoxFrame(.FrameCounter - 1), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        End With
        Call BoxSurface.BltToDC(frmMain.picInv.hdc, auxr, rBox)
        Call frmMain.picInv.Refresh
    End If
End Sub

Sub DibujarInv()
Dim IX As Integer

If Not bStaticInit Then _
    Call InitMem

r1.Top = 0: r1.Left = 0: r1.Right = 32: r1.Bottom = 32
r2.Top = 0: r2.Left = 0: r2.Right = 32: r2.Bottom = 32

frmMain.picInv.Cls
frmMain.Arma.Caption = "0"
frmMain.Tronco.Caption = "0"
frmMain.Cabeza.Caption = "0"
frmMain.Botas.Caption = "0"
frmMain.Escudo.Caption = "0"

For IX = OffsetDelInv + 1 To UBound(UserInventory)
    If UserInventory(IX).GrhIndex > 0 Then
       ' AuxSurface.BltColorFill auxr, vbBlack
        'pluto:6.5
        If GrhData(UserInventory(IX).GrhIndex).FileNum = 0 Then GoTo nap
'AuxSurface.BltFast 0, 0, SurfaceDB.GetBMP(GrhData(UserInventory(IX).GrhIndex).FileNum), auxr, DDBLTFAST_NOCOLORKEY

'pluto:7.0 quito la de arriba y pongo estas dos para inventario sin fondo negro.
AuxSurface.BltFast 0, 0, SurfaceDB.GetBMP(8008), auxr, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT

AuxSurface.BltFast 0, 0, SurfaceDB.GetBMP(GrhData(UserInventory(IX).GrhIndex).FileNum), auxr, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT

nap:
        AuxSurface.SetForeColor vbWhite

        
      AuxSurface.SetFont FrmPremios.font
        AuxSurface.DrawText 0, 0, UserInventory(IX).Amount, False


        If UserInventory(IX).Equipped Then
           
           AuxSurface.SetForeColor vbYellow
           AuxSurface.DrawText 0, 0, UserInventory(IX).Amount, False
          AuxSurface.SetFont frmMain.font
            AuxSurface.DrawText 20, 20, "+", False
            AuxSurface.SetForeColor vbWhite
'pluto:6.0A---silueta humana----------
If UserInventory(IX).OBJType = 2 Then
frmMain.Arma.Caption = UserInventory(IX).MaxHIT
End If
If UserInventory(IX).OBJType = 3 And UserInventory(IX).SubTipo = 0 Then
frmMain.Tronco.Caption = UserInventory(IX).DefMax
End If
If UserInventory(IX).OBJType = 3 And UserInventory(IX).SubTipo = 1 Then
frmMain.Cabeza.Caption = UserInventory(IX).DefMax
End If
If UserInventory(IX).OBJType = 3 And UserInventory(IX).SubTipo = 2 Then
frmMain.Escudo.Caption = UserInventory(IX).DefMax
End If
If UserInventory(IX).OBJType = 3 And UserInventory(IX).SubTipo = 3 Then
frmMain.Botas.Caption = UserInventory(IX).DefMax
End If
  '---------------
  
        End If

AuxSurface.SetFont frmMain.font
        
        If ItemElegido = IX Then
            With r2: .Left = (mx - 1) * 32: .Right = r2.Left + 32: .Top = (my - 1) * 32: .Bottom = r2.Top + 32: End With
            Call AuxSurface.BltFast(0, 0, SurfaceDB.GetBMP(GrhData(GrhData(Grh(1).GrhIndex).Frames(2)).FileNum), rBoxFrame(2), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        End If
        AuxSurface.BltToDC frmMain.picInv.hdc, auxr, r2
    End If

    r2.Left = r2.Left + 32
    r2.Right = r2.Right + 32
    r1.Left = r1.Left + 32
    r1.Right = r1.Right + 32
    If r2.Left >= 160 Then
        r2.Left = 0
        r1.Left = 0
        r1.Right = 32
        r2.Right = 32
        r2.Top = r2.Top + 32
        r1.Top = r1.Top + 32
        r2.Bottom = r2.Bottom + 32
        r1.Bottom = r1.Bottom + 32
    End If
Next IX

bInvMod = False

If ItemElegido = 0 Then _
    Call ItemClick(2, 2)

End Sub
Private Sub InitMem()
    Dim ddck        As DDCOLORKEY
    Dim SurfaceDesc As DDSURFACEDESC2
    
    'Back Buffer Surface
    r1.Right = 32: r1.Bottom = 32
    r2.Right = 32: r2.Bottom = 32
    
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lHeight = r1.Bottom
        .lWidth = r1.Right
    End With

    ' Create surface
    Set AuxSurface = DirectDraw.CreateSurface(SurfaceDesc)
    Set BoxSurface = DirectDraw.CreateSurface(SurfaceDesc)
    Set SelSurface = DirectDraw.CreateSurface(SurfaceDesc)

    'Set color key
    AuxSurface.SetColorKey DDCKEY_SRCBLT, ddck
    BoxSurface.SetColorKey DDCKEY_SRCBLT, ddck
    SelSurface.SetColorKey DDCKEY_SRCBLT, ddck

    auxr.Right = 32: auxr.Bottom = 32

    AuxSurface.SetFontTransparency True
    AuxSurface.SetFont frmMain.font
    SelSurface.SetFontTransparency True
    SelSurface.SetFont frmMain.font

    'RedBox Frame Position List
    With rBoxFrame(0): .Left = 0:  .Top = 0: .Right = 32: .Bottom = 32: End With
    With rBoxFrame(1): .Left = 32: .Top = 0: .Right = 64: .Bottom = 32: End With
    With rBoxFrame(2): .Left = 64: .Top = 0: .Right = 96: .Bottom = 32: End With
    iFrameMod = 1

    bStaticInit = True
End Sub
