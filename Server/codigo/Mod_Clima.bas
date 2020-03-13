Attribute VB_Name = "Mod_Clima"
'********************************Modulo Climas*********************************
'Author: Manuel (Lorwik)
'Last Modification: 17/11/2011
'Controla el clima y lo envia al cliente.
'Nota: Cuando reformemos el sistema de lluvias, todo va a ir aqui.
'******************************************************************************
Option Explicit

Enum eColorEstado
    Amanecer = 0
    MedioDia
    Tarde
    Noche
    Lluvia
    Nieve
    Niebla
    FogLluvia 'Niebla mas lluvia
End Enum

Public DayStatus As Byte 'Establece el color actual del dia
Public Lloviendo As Boolean

'Todo en minutos:
Private Const CicloClima As Byte = 45 'Intervalo de clima
Private Const RandomCiclo As Byte = 10 'Intervalo para el tiempo random
Private Const DuracionClima As Byte = 5 'Tiempo que va a durar el clima una vez activo
Private Const FogProb As Byte = 5 'Porcentaje del 1 al 100 de la lluvia sea con niebla

Public Sub SortearHorario(Optional ByVal Clima As Byte)
'***************************************************************************************
'Autor: Lorwik
'Ultima modificación: 23/12/2018
'Descripción: Sorteamos el clima, si hay tormenta y es de Mañana o de Dia
'ponemos el efecto de tarde, pero si es de Tarde o de Noche no ponemos nigun efecto.
'***************************************************************************************

    If Lloviendo = True Then 'Si esta lloviendo ignoramos el resto y solo mandamos el estado lluvia
        Call ColorClima(Clima)
        frmMain.Horario.Caption = "Hora: Lloviendo - [" & Hour(Now) & ":" & Minute(Now) & "]"
    Else
        If (Hour(Now) >= 3 And Hour(Now) < 5) Or (Hour(Now) >= 15 And Hour(Now) < 17) Then 'Amanecer
            Call ColorClima(eColorEstado.Amanecer)
            frmMain.Horario.Caption = "Hora: Mañana - [" & Hour(Now) & ":" & Minute(Now) & "]"
            
        ElseIf (Hour(Now) >= 6 And Hour(Now) < 8) Or (Hour(Now) >= 18 And Hour(Now) < 20) Then 'MedioDia
            Call ColorClima(eColorEstado.MedioDia)
            frmMain.Horario.Caption = "Hora: MedioDia - [" & Hour(Now) & ":" & Minute(Now) & "]"
            
        ElseIf (Hour(Now) >= 9 And Hour(Now) < 11) Or (Hour(Now) >= 21 And Hour(Now) < 23) Then 'Tarde
            Call ColorClima(eColorEstado.Tarde)
            frmMain.Horario.Caption = "Hora: Tarde - [" & Hour(Now) & ":" & Minute(Now) & "]"
            
        ElseIf (Hour(Now) >= 0 And Hour(Now) < 2) Or (Hour(Now) >= 12 And Hour(Now) < 14) Then 'Noche
            Call ColorClima(eColorEstado.Noche)
            frmMain.Horario.Caption = "Hora: Noche - [" & Hour(Now) & ":" & Minute(Now) & "]"
        End If
    End If
End Sub

'Enviamos el Clima
Private Sub ColorClima(Clima As Byte)
'****************************************
'Autor: Lorwik
'Ultima modificación: ?????
'Enviamos el clima
'0 Amanecer, 1 Mediodia, 2 Tarde, 3 Noche
'****************************************
Dim UserIndex As Integer
Dim i As Long
    
    DayStatus = Clima
    
    For i = 1 To LastUser
        Call WriteClima(i, Clima)
    Next i
End Sub

Public Sub SortearClima(ByVal Contador As Byte, Optional ByVal Forzado As Boolean = False, Optional ByVal Clima As Byte = 4)
'**********************************************
'Autor: Lorwik
'Ultima modificación: 23/12/2018
'**********************************************
    Static Intervalo As Byte
    Static Count As Byte
    
    'Si aun no tenemos un intervalo lo pedimos
    If Intervalo = 0 Then Intervalo = ObtenerTiempoRandom
    
    '¿El contador que recibimos es igual al intervalo que tenemos o lo hemos forzado?
    If Intervalo = Contador Or Forzado = True Then
        Lloviendo = True
        If FogProb >= RandomNumber(1, 100) Then 'Probabilidad de que la lluvia sea con niebla
            Clima = 7
        End If
        Call SortearHorario(Clima)
    End If
    
    'Si ya esta lloviendo vamos a llevar la cuenta de cuanto tiempo lleva
    If Lloviendo = True Then Count = Count + 1
    
    Debug.Print "Lloviendo: " & Lloviendo & "contador: " & Contador & " Intervalo: " & Intervalo & " Count: " & Count
    
    '¿Ha pasado el tiempo que tenia que durar la lluvia?
    If DuracionClima = Count Then
        Lloviendo = False
        Count = 0
        Call SortearHorario
        Intervalo = ObtenerTiempoRandom 'Volvemos a sortear el tiempo random
    End If
End Sub

Private Function ObtenerTiempoRandom() As Byte
'****************************************************
'Autor: Lorwik
'Ultima modificación: 23/12/2018
'Descripción: Randomizamos el intervalo del clima para que no sea siempre el mismo exacto
'****************************************************

    ObtenerTiempoRandom = RandomNumber(CicloClima, CicloClima + RandomCiclo)

End Function
