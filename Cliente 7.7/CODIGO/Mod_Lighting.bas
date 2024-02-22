Attribute VB_Name = "Mod_Lighting"
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

Dim surface As DDSURFACEDESC2
Dim numRedBits As Integer
Dim numGreenBits As Integer
Dim numBlueBits As Integer

'Public Sub GetPixelFormat()
'    Dim format As DDPIXELFORMAT
'    Dim loGreenBit As Integer
'    Dim hiGreenBit As Integer
'
'    BackBufferSurface.GetPixelFormat format
'    loGreenBit = LowBitPos(format.lGBitMask)
'    hiGreenBit = HighBitPos(format.lGBitMask)
'    numGreenBits = (hiGreenBit - loGreenBit) + 1
'
'    If numGreenBits = 5 Then
'        numRedBits = 10
'    Else
'        numRedBits = 11
'    End If
'End Sub

