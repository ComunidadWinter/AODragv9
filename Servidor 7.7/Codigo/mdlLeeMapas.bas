Attribute VB_Name = "mdlLeeMapas"
'Argentum Online 0.11.20
'Copyright (C) 2002 Márquez Pablo Ignacio
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

'unsigned int DLLIMPORT MAPCargaMapa (const char *archmap, const char *archinf);
'unsigned int DLLIMPORT MAPCierraMapa(unsigned int dm);
'
'unsigned int DLLIMPORT MAPLeeMapa(unsigned int dm, BLOQUE *tile_map, BLOQUE_INF *tile_inf );
'

Public Type TileMap
    bloqueado As Byte
    grafs(1 To 4) As Integer
    trigger As Integer

    t1 As Integer 'espacio al pedo
End Type

Public Type TileInf
    dest_mapa As Integer
    dest_x As Integer
    dest_y As Integer
    
    npc As Integer
    
    obj_ind As Integer
    obj_cant As Integer
    
    t1 As Integer
    t2 As Integer
End Type

'Public Declare Function MAPCargaMapa Lib "LeeMapas.dll" (ByVal archmap As String, ByVal archinf As String) As Long
'Public Declare Function MAPCierraMapa Lib "LeeMapas.dll" (ByVal Dm As Long) As Long
'
'Public Declare Function MAPLeeMapa Lib "LeeMapas.dll" (ByVal Dm As Long, Tile_Map As TileMap, Tile_Inf As TileInf) As Long

