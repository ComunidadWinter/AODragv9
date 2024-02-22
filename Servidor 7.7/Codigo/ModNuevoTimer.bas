Attribute VB_Name = "ModNuevoTimer"

' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF
'pluto: 6.9
Dim Rapidos As Long
Rapidos = TActual - UserList(UserIndex).Counters.TimerLanzarSpell
If Rapidos < TOPELANZAR Then
Call SendData(ToGM, 0, 0, "||" & UserList(UserIndex).Name & " lanza en:" & Rapidos & "´" & FontTypeNames.FONTTYPE_talk)
Call LogCasino("Lanza: " & UserList(UserIndex).Name & " HD: " & UserList(UserIndex).Serie & " en " & Rapidos)
End If

If Rapidos >= 40 * IntervaloUserPuedeCastear Then
    If Actualizar Then UserList(UserIndex).Counters.TimerLanzarSpell = TActual
    IntervaloPermiteLanzarSpell = True
Else
    IntervaloPermiteLanzarSpell = False
End If

End Function


Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
    IntervaloPermiteAtacar = True
Else
    IntervaloPermiteAtacar = False
End If
End Function



' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= 40 * IntervaloUserPuedeTrabajar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
    IntervaloPermiteTrabajar = True
Else
    IntervaloPermiteTrabajar = False
End If
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerUsar = TActual
    IntervaloPermiteUsar = True
Else
    IntervaloPermiteUsar = False
End If

End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF
Dim Rapidos As Long
Rapidos = TActual - UserList(UserIndex).Counters.TimerUsarArco
If Rapidos < TOPEFLECHA Then
Call SendData(ToGM, 0, 0, "||" & UserList(UserIndex).Name & " tira flecha en:" & Rapidos & "´" & FontTypeNames.FONTTYPE_talk)
Call LogCasino("Flecha: " & UserList(UserIndex).Name & " HD: " & UserList(UserIndex).Serie & " en " & Rapidos)
End If
If TActual - UserList(UserIndex).Counters.TimerUsarArco >= IntervaloUserPuedeFlechas Then
    If Actualizar Then UserList(UserIndex).Counters.TimerUsarArco = TActual
    IntervaloPermiteUsarArcos = True
Else
    IntervaloPermiteUsarArcos = False
End If

End Function

Public Function IntervaloPermiteTomar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerTomar >= 40 * IntervaloUserPuedeTomar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerTomar = TActual
    IntervaloPermiteTomar = True
Else
    IntervaloPermiteTomar = False
End If

End Function


