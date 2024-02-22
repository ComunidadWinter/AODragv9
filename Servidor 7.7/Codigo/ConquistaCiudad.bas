Attribute VB_Name = "ConquistaCiudad"
Sub ConquistarCiudad(ByVal Mapa As Integer, ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim obj As obj
Dim X As Byte
Dim Y As Byte
Dim au(1 To 5) As Integer
Dim aumax As Byte
Dim xc As Byte

'pluto:2.17 coquista ciudad un criminal
'-----------------------------------------

aumax = 0
au(1) = 0
au(2) = 0
au(3) = 0
au(4) = 0
au(5) = 0
'cambia decoracion a caos
'MapInfo(mapa).Dueño = 1
If UserList(UserIndex).Faccion.FuerzasCaos > 0 And MapInfo(Mapa).Dueño = 1 Then
MapInfo(Mapa).Dueño = 2

'ulla mapa
If Mapa = 251 Then
MapInfo(1).Dueño = 2
au(1) = 1
aumax = 1
End If
'desierto mapa
If Mapa = 252 Then
MapInfo(20).Dueño = 2
au(1) = 20
aumax = 1
End If

'[Tite] Nix neutral (comentar)

'Nix Mapa
'If Mapa = 253 Then
'MapInfo(34).Dueño = 2
'au(1) = 34
'aumax = 1
'End If

'[\Tite]

'lindos multiple mapa
If Mapa = 254 Then
MapInfo(62).Dueño = 2
MapInfo(64).Dueño = 2
MapInfo(63).Dueño = 2
au(1) = 63
au(2) = 62
au(3) = 64
aumax = 3
End If

'descanso mapa
If Mapa = 255 Then
MapInfo(81).Dueño = 2
au(1) = 81
aumax = 1
End If


'atlantis multiple mapa
If Mapa = 256 Then
MapInfo(83).Dueño = 2
MapInfo(84).Dueño = 2
MapInfo(85).Dueño = 2
aumax = 3
au(1) = 83
au(2) = 84
au(3) = 85
End If

'esperanza multiple mapa
If Mapa = 257 Then
'MapInfo(111).Dueño = 2
MapInfo(112).Dueño = 2
au(1) = 112
'au(2) = 111
aumax = 1
End If
'arghal multiple mapa
If Mapa = 258 Then
au(1) = 151
au(2) = 150
aumax = 2
MapInfo(150).Dueño = 2
MapInfo(151).Dueño = 2
End If

'quest mapa
If Mapa = 259 Then
MapInfo(157).Dueño = 2
au(1) = 157
aumax = 1
End If

'laurana multiple mapa
If Mapa = 260 Then
au(1) = 183
au(2) = 184
aumax = 2
MapInfo(184).Dueño = 2
MapInfo(183).Dueño = 2
End If


For xc = 1 To aumax
For Y = 1 To 100
For X = 1 To 100
If X > 0 And Y > 0 And X < 101 And Y < 101 Then
If MapData(au(xc), X, Y).OBJInfo.ObjIndex = 1024 Then
obj.Amount = 1
obj.ObjIndex = 1023
Call EraseObj(ToMap, 0, au(xc), MapData(au(xc), X, Y).OBJInfo.Amount, au(xc), X, Y)
Call MakeObj(ToMap, 0, au(xc), obj, au(xc), X, Y)
End If
End If
Next X
Next Y
Next xc
'---------------------------------
Call SendData(ToAll, 0, 0, "|| Conquistada Ciudad " & MapInfo(au(1)).Name & " por las Fuerzas del Caos." & "´" & FontTypeNames.FONTTYPE_info)
Call ReSpawnCambioGuardias

End If ' fuerza caos

If UserList(UserIndex).Faccion.ArmadaReal > 0 And MapInfo(Mapa).Dueño = 2 Then

MapInfo(Mapa).Dueño = 1
'ulla mapa
If Mapa = 251 Then
MapInfo(1).Dueño = 1
au(1) = 1
aumax = 1
End If
'desierto mapa
If Mapa = 252 Then
MapInfo(20).Dueño = 1
au(1) = 20
aumax = 1
End If

'[Tite]Nix neutral(comentar)

'nix mapa
'If Mapa = 253 Then
'MapInfo(34).Dueño = 1
'au(1) = 34
'aumax = 1
'End If

'[\Tite]

'lindos multiple mapa
If Mapa = 254 Then
MapInfo(62).Dueño = 1
MapInfo(64).Dueño = 1
MapInfo(63).Dueño = 1
au(1) = 63
au(2) = 62
au(3) = 64
aumax = 3
End If

'descanso mapa
If Mapa = 255 Then
MapInfo(81).Dueño = 1
au(1) = 81
aumax = 1
End If


'atlantis multiple mapa
If Mapa = 256 Then
MapInfo(85).Dueño = 1
MapInfo(83).Dueño = 1
MapInfo(84).Dueño = 1
aumax = 3
au(1) = 83
au(2) = 84
au(3) = 85
End If

'esperanza multiple mapa
If Mapa = 257 Then
'MapInfo(111).Dueño = 1
MapInfo(112).Dueño = 1
au(1) = 112
'au(2) = 111
aumax = 1
End If
'arghal multiple mapa
If Mapa = 258 Then
au(1) = 151
au(2) = 150
aumax = 2
MapInfo(150).Dueño = 1
MapInfo(151).Dueño = 1
End If

'quest mapa
If Mapa = 259 Then
MapInfo(157).Dueño = 1
au(1) = 157
aumax = 1
End If

'laurana multiple mapa
If Mapa = 260 Then
au(1) = 183
au(2) = 184
aumax = 2
MapInfo(184).Dueño = 1
MapInfo(183).Dueño = 1
End If


For xc = 1 To aumax
For Y = 1 To 100
For X = 1 To 100
If X > 0 And Y > 0 And X < 101 And Y < 101 Then
If MapData(au(xc), X, Y).OBJInfo.ObjIndex = 1023 Then
obj.Amount = 1
obj.ObjIndex = 1024
Call EraseObj(ToMap, 0, au(xc), MapData(au(xc), X, Y).OBJInfo.Amount, au(xc), X, Y)
Call MakeObj(ToMap, 0, au(xc), obj, au(xc), X, Y)
End If
End If
Next X
Next Y
Next xc
'---------------------------------
Call SendData(ToAll, 0, 0, "|| Conquistada Ciudad " & MapInfo(au(1)).Name & " por las Fuerzas del Bien." & "´" & FontTypeNames.FONTTYPE_info)
Call ReSpawnCambioGuardias

End If 'real
'envia dueño mapas
 Dim n As Integer
 Dim ci As String
 ci = ""
 'For n = 1 To NumMaps
 'If MapInfo(n).Dueño > 0 Then ci = ci + str(MapInfo(n).Dueño) & ","
 
'[Tite]Nix neutral
ci = str(MapInfo(1).Dueño) & "," & str(MapInfo(20).Dueño) & "," & str(MapInfo(63).Dueño) & "," & str(MapInfo(81).Dueño) & "," & str(MapInfo(84).Dueño) & "," & str(MapInfo(112).Dueño) & "," & str(MapInfo(151).Dueño) & "," & str(MapInfo(157).Dueño) & "," & str(MapInfo(184).Dueño)
'ci = str(MapInfo(1).Dueño) & "," & str(MapInfo(20).Dueño) & "," & str(MapInfo(34).Dueño) & "," & str(MapInfo(63).Dueño) & "," & str(MapInfo(81).Dueño) & "," & str(MapInfo(84).Dueño) & "," & str(MapInfo(112).Dueño) & "," & str(MapInfo(151).Dueño) & "," & str(MapInfo(157).Dueño) & "," & str(MapInfo(184).Dueño)
'[\Tite]

' Next n
 Call SendData(ToAll, 0, 0, "K4" & ci)
'------------------------------------------------------
Exit Sub

errhandler:
    Call LogError("Error en ConquistaCiudad")
End Sub
'[Tite] Añado esta subrutina para enviar un listado de las ciudades y sus dueños
Sub sendCiudades(UserIndex As Integer)
Dim n As Integer
For n = 1 To NumMaps
    If n = 1 Or n = 20 Or n = 63 Or n = 81 Or n = 84 Or n = 112 Or n = 151 Or n = 157 Or n = 184 Then
        If MapInfo(n).Dueño = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & MapInfo(n).Name & ": Armada Real" & "´" & FontTypeNames.FONTTYPE_info)
        ElseIf MapInfo(n).Dueño = 2 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & MapInfo(n).Name & ": Fuerzas del caos" & "´" & FontTypeNames.FONTTYPE_info)
        End If
    End If
Next
End Sub
'[\Tite]
