Attribute VB_Name = "Mod_Clima"
'********************************Modulo Climas*********************************
'Author: Manuel (Lorwik)
'Last Modification: 17/11/2011
'Controla el clima y lo envia al cliente.
'Nota: Cuando reformemos el sistema de lluvias, todo va a ir aqui.
'******************************************************************************

Option Explicit

Public DayStatus As Byte

'******************************************************************************
'Sorteo Del Horario
'******************************************************************************
Public Function SortearHorario()
'Sorteamos el clima, si hay tormenta y es de Mañana o de Dia ponemos el efecto de tarde,
'pero si es de Tarde o de Noche no ponemos nigun efecto.

    If Hour(Now) >= 6 And Hour(Now) < 12 Then
        Call ColorClima(0)
        frmMain.Horario.Caption = "Hora: Mañana - [" & Hour(Now) & ":" & Minute(Now) & "]"
    ElseIf Hour(Now) >= 12 And Hour(Now) < 18 Then
        Call ColorClima(1)
        frmMain.Horario.Caption = "Hora: MedioDia - [" & Hour(Now) & ":" & Minute(Now) & "]"
    ElseIf Hour(Now) >= 18 And Hour(Now) < 20 Then
        Call ColorClima(2)
        frmMain.Horario.Caption = "Hora: Tarde - [" & Hour(Now) & ":" & Minute(Now) & "]"
    ElseIf Hour(Now) >= 20 And Hour(Now) < 6 Then
        Call ColorClima(3)
        frmMain.Horario.Caption = "Hora: Noche - [" & Hour(Now) & ":" & Minute(Now) & "]"
    End If
End Function

'Enviamos el Clima
Public Function ColorClima(Clima As Byte)
'****************************************
'Enviamos el clima
'0 Mañana, 1 Mediodia, 2 Tarde, 3 Noche
'****************************************
Dim UserIndex As Integer
Dim i As Long
    
    Clima = DayStatus
    
    For i = 1 To LastUser
        Call WriteNoche(i, Clima)
    Next i
End Function


