Attribute VB_Name = "Matematicas"
Option Explicit

Sub AddtoVar(ByRef Var As Variant, ByVal Addon As Variant, ByVal max As Variant)
'Le suma un valor a una variable respetando el maximo valor
On Error GoTo fallo
If Var >= max Then
    Var = max
Else
    Var = Var + Addon
    If Var > max Then
        Var = max
    End If
End If
Exit Sub
fallo:
Call LogError("addtovar " & Err.number & " D: " & Err.Description)

End Sub

Public Function Porcentaje(ByVal total As Long, ByVal Porc As Long) As Long
On Error GoTo fallo
Porcentaje = (total * Porc) / 100

Exit Function
fallo:
Call LogError("porcentaje " & Err.number & " D: " & Err.Description)
End Function

Function Distancia(wp1 As WorldPos, wp2 As WorldPos)
On Error GoTo fallo
Dim DistanciaX As Integer
Dim DistanciaY As Integer
DistanciaX = Abs(wp1.X - wp2.X)
DistanciaY = Abs(wp1.Y - wp2.Y)
If DistanciaX > 8 Then
Distancia = 20
Exit Function
End If
If DistanciaY > 6 Then
Distancia = 20
Exit Function
End If
Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100)

Exit Function
fallo:
Call LogError("distancia " & Err.number & " D: " & Err.Description)

End Function

Function Distance(x1 As Variant, Y1 As Variant, x2 As Variant, Y2 As Variant) As Double
On Error GoTo fallo
'Encuentra la distancia entre dos puntos

Distance = Sqr(((Y1 - Y2) ^ 2 + (x1 - x2) ^ 2))
Exit Function
fallo:
Call LogError("distance " & Err.number & " D: " & Err.Description)

End Function

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
On Error GoTo fallo
Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

Exit Function
fallo:
Call LogError("randomnumber " & Err.number & " D: " & Err.Description)

End Function
Function Vabs(X As Double) As Integer
    If X < 0 Then
        Vabs = 0 - X
    Else
        Vabs = X
    End If
End Function

