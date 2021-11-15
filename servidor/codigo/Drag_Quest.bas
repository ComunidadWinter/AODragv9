Attribute VB_Name = "Drag_Quest"
Option Explicit


Public Enum eEstadoQuest
    Nada = 1
    Completada = 2
    FaltaNivel = 3
    EnProgreso = 4
    NoAceptada = 5
    CompletadaFaltaEntregar = 6
End Enum



Public Function EstadoQuest(ByVal NPCId As Integer, ByVal PersonajeId As Integer) As eEstadoQuest
  Set RS = New ADODB.Recordset
  Set RS = SQL.Execute("SELECT estado, id_quest FROM rel_quest_personaje_estado WHERE estado > '2' AND id_personaje = '" & PersonajeId & "' ORDER BY estado DESC LIMIT 1")
  If Not (RS.EOF = True Or RS.BOF = True) Then
    EstadoQuest = RS!estado
  Else
    EstadoQuest = 1
    Exit Function
  End If
End Function


Public Function QuestCompletadaPorPersonaje(ByVal QuestId As Integer, ByVal PersonajeId As Integer)
  Set RS = New ADODB.Recordset
  Set RS = SQL.Execute("SELECT estado FROM rel_quest_personaje WHERE estado = '2' AND id_personaje = '" & PersonajeId & "' AND id_quest = '" & QuestId & "'")
  If Not (RS.EOF = True Or RS.BOF = True) Then
    QuestCompletadaPorPersonaje = False
  Else
    QuestCompletadaPorPersonaje = True
  End If
  RS.Close
End Function

Public Function ListaQuestsEnNPCParaJugador(ByVal NPCId As Integer, ByVal PersonajeId As Integer)

  ' listar todas las quests del npc


End Function



