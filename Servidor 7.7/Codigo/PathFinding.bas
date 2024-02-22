Attribute VB_Name = "PathFinding"
Option Explicit

Private Const ROWS = 100
Private Const COLUMS = 100
Private Const MAXINT = 1000
Private Const Walkable = 0

Private Type tIntermidiateWork
    Known As Boolean
    DistV As Integer
    PrevV As tVertice
End Type

Dim TmpArray(1 To ROWS, 1 To COLUMS) As tIntermidiateWork

Dim TilePosX As Integer, TilePosY As Integer

Dim MyVert As tVertice
Dim MyFin As tVertice

Dim Iter As Integer

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer)
On Error GoTo fallo
Limites = vcolu >= 1 And vcolu <= COLUMS And vfila >= 1 And vfila <= ROWS
Exit Function
fallo:
Call LogError("limites " & Err.number & " D: " & Err.Description)


End Function

Private Function IsWalkable(ByVal Map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
On Error GoTo fallo
IsWalkable = MapData(Map, row, Col).Blocked = 0 And MapData(Map, row, Col).NpcIndex = 0

If MapData(Map, row, Col).UserIndex <> 0 Then
     If MapData(Map, row, Col).UserIndex <> Npclist(NpcIndex).PFINFO.TargetUser Then IsWalkable = False
End If
Exit Function
fallo:
Call LogError("iswalkable " & Err.number & " D: " & Err.Description)

End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef t() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)
  On Error GoTo fallo
    Dim V As tVertice
    Dim j As Integer
    'Look to North
    j = vfila - 1
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                    'Nos aseguramos que no hay un camino más corto
                    If t(j, vcolu).DistV = MAXINT Then
                        'Actualizamos la tabla de calculos intermedios
                        t(j, vcolu).DistV = t(vfila, vcolu).DistV + 1
                        t(j, vcolu).PrevV.X = vcolu
                        t(j, vcolu).PrevV.Y = vfila
                        'Mete el vertice en la cola
                        V.X = vcolu
                        V.Y = j
                        Call Push(V)
                    End If
            End If
    End If
    j = vfila + 1
    'look to south
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If t(j, vcolu).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    t(j, vcolu).DistV = t(vfila, vcolu).DistV + 1
                    t(j, vcolu).PrevV.X = vcolu
                    t(j, vcolu).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu
                    V.Y = j
                    Call Push(V)
                End If
            End If
    End If
    'look to west
    If Limites(vfila, vcolu - 1) Then
            If IsWalkable(MapIndex, vfila, vcolu - 1, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If t(vfila, vcolu - 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    t(vfila, vcolu - 1).DistV = t(vfila, vcolu).DistV + 1
                    t(vfila, vcolu - 1).PrevV.X = vcolu
                    t(vfila, vcolu - 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu - 1
                    V.Y = vfila
                    Call Push(V)
                End If
            End If
    End If
    'look to east
    If Limites(vfila, vcolu + 1) Then
            If IsWalkable(MapIndex, vfila, vcolu + 1, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If t(vfila, vcolu + 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    t(vfila, vcolu + 1).DistV = t(vfila, vcolu).DistV + 1
                    t(vfila, vcolu + 1).PrevV.X = vcolu
                    t(vfila, vcolu + 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu + 1
                    V.Y = vfila
                    Call Push(V)
                End If
            End If
    End If
   
   Exit Sub
fallo:
Call LogError("processadjacents " & Err.number & " D: " & Err.Description)

End Sub


Public Sub SeekPath(ByVal NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
'############################################################
'This Sub seeks a path from the npclist(npcindex).pos
'to the location NPCList(NpcIndex).PFINFO.Target.
'The optional parameter MaxSteps is the maximum of steps
'allowed for the path.
'############################################################
On Error GoTo fallo
Dim cur_npc_pos As tVertice
Dim tar_npc_pos As tVertice
Dim V As tVertice
Dim NpcMap As Integer
Dim steps As Integer

NpcMap = Npclist(NpcIndex).Pos.Map

steps = 0

cur_npc_pos.X = Npclist(NpcIndex).Pos.Y
cur_npc_pos.Y = Npclist(NpcIndex).Pos.X

tar_npc_pos.X = Npclist(NpcIndex).PFINFO.Target.X '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.X
tar_npc_pos.Y = Npclist(NpcIndex).PFINFO.Target.Y '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.Y

Call InitializeTable(TmpArray, cur_npc_pos)
Call InitQueue

'We add the first vertex to the Queue
Call Push(cur_npc_pos)

Do While (Not IsEmpty)
    If steps > MaxSteps Then Exit Do
    V = Pop
    If V.X = tar_npc_pos.X And V.Y = tar_npc_pos.Y Then Exit Do
    Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.X, NpcIndex)
Loop

Call MakePath(NpcIndex)
Exit Sub
fallo:
Call LogError("seekpath " & Err.number & " D: " & Err.Description)

End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
'#######################################################
'Builds the path previously calculated
'#######################################################
On Error GoTo fallo
Dim Pasos As Integer
Dim miV As tVertice
Dim i As Integer

Pasos = TmpArray(Npclist(NpcIndex).PFINFO.Target.Y, Npclist(NpcIndex).PFINFO.Target.X).DistV
Npclist(NpcIndex).PFINFO.PathLenght = Pasos


If Pasos = MAXINT Then
    Npclist(NpcIndex).PFINFO.NoPath = True
    Npclist(NpcIndex).PFINFO.PathLenght = 0
    Exit Sub
End If

ReDim Npclist(NpcIndex).PFINFO.Path(0 To Pasos) As tVertice

miV.X = Npclist(NpcIndex).PFINFO.Target.X
miV.Y = Npclist(NpcIndex).PFINFO.Target.Y

For i = Pasos To 1 Step -1
    Npclist(NpcIndex).PFINFO.Path(i) = miV
    miV = TmpArray(miV.Y, miV.X).PrevV
Next i

Npclist(NpcIndex).PFINFO.CurPos = 1
Npclist(NpcIndex).PFINFO.NoPath = False
   Exit Sub
fallo:
Call LogError("makepath " & Err.number & " D: " & Err.Description)

End Sub

Private Sub InitializeTable(ByRef t() As tIntermidiateWork, ByRef s As tVertice, Optional ByVal MaxSteps As Integer = 30)
'#########################################################
'Initialize the array where we calculate the path
'#########################################################
On Error GoTo fallo
Dim j As Integer, k As Integer
Const anymap = 1
For j = s.Y - MaxSteps To s.Y + MaxSteps
    For k = s.X - MaxSteps To s.X + MaxSteps
        If InMapBounds(anymap, j, k) Then
            t(j, k).Known = False
            t(j, k).DistV = MAXINT
            t(j, k).PrevV.X = 0
            t(j, k).PrevV.Y = 0
        End If
    Next
Next

t(s.Y, s.X).Known = False
t(s.Y, s.X).DistV = 0
Exit Sub
fallo:
Call LogError("initializetable " & Err.number & " D: " & Err.Description)

End Sub

