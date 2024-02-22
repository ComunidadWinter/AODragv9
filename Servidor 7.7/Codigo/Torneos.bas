Attribute VB_Name = "Torneos"
Dim Participantes As String
Dim X As Integer

Sub CrearTorneo(rdata As String)
On Error GoTo fallo
For X = 1 To 8
TorneoPluto.Participantes(X) = ""
Next
TorneoPluto.FaseTorneo = 1
TorneoPluto.Creador = ReadField(1, rdata, 44)
TorneoPluto.Ttip = val(ReadField(2, rdata, 44))
TorneoPluto.Tcua = val(ReadField(3, rdata, 44))
TorneoPluto.Tpj = val(ReadField(4, rdata, 44))
TorneoPluto.Tmax = val(ReadField(5, rdata, 44))
TorneoPluto.Tmin = val(ReadField(6, rdata, 44))
TorneoPluto.Tins = val(ReadField(7, rdata, 44))
TorneoPluto.Participantes(1) = TorneoPluto.Creador
'Call SendData2(ToIndex, UserIndex, 0, 91, rdata)
Exit Sub
fallo:
Call LogError("creartorneo " & Err.number & " D: " & Err.Description)

End Sub

Sub EnviarTorneo(index As Integer)
On Error GoTo fallo
Participantes = ""
For X = 1 To 8
Participantes = Participantes + TorneoPluto.Participantes(X) & ","
Next
Call SendData2(ToTorneo, index, 0, 91, TorneoPluto.FaseTorneo & "," & TorneoPluto.Creador & "," & TorneoPluto.Ttip & "," & TorneoPluto.Tcua & "," & TorneoPluto.Tpj & "," & TorneoPluto.Tmax & "," & TorneoPluto.Tmin & "," & TorneoPluto.Tins & "," & Participantes)

Exit Sub
fallo:
Call LogError("enviartorneo " & Err.number & " D: " & Err.Description)

End Sub

Sub ParticipaTorneo(User As String)
On Error GoTo fallo
Dim gente As Integer
Dim index As Integer
If TorneoPluto.Ttip = 1 Then gente = 2 Else gente = 8
For X = 1 To gente
If TorneoPluto.Participantes(X) = User Then Exit Sub
If TorneoPluto.Participantes(X) = "" Then TorneoPluto.Participantes(X) = User
Next
 index = NameIndex(User)
Call EnviarTorneo(index)

Exit Sub
fallo:
Call LogError("participatorneo " & Err.number & " D: " & Err.Description)

End Sub

