Attribute VB_Name = "DiaEspecial"
Sub CargarDiaEspecial()
Dim npcfile As String
Dim Bicho As Integer

npcfile = DatPath & "NPCs-HOSTILES.dat"
a:
Bicho = RandomNumber(500, 711)
'Bicho = 538
    If val(GetVar(npcfile, "NPC" & Bicho, "diaespecial")) = 1 Then
    BichoDelDia = Bicho
    NombreBichoDelDia = GetVar(npcfile, "NPC" & Bicho, "Name")
    Else
    GoTo a
    End If

End Sub
