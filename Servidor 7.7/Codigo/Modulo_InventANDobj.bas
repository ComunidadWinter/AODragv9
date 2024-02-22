Attribute VB_Name = "InvNpc"
Option Explicit
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, obj As obj) As WorldPos
On Error GoTo fallo

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    Call Tilelibre(Pos, NuevaPos)
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
          Call MakeObj(ToMap, 0, Pos.Map, _
          obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
          TirarItemAlPiso = NuevaPos
    End If

Exit Function
fallo:
Call LogError("TIRARITEMALPISO" & Err.number & " D: " & Err.Description)


End Function

Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc)
'TIRA TODOS LOS ITEMS DEL NPC
On Error GoTo fallo
Dim Notodo As Integer

If npc.Invent.NroItems > 0 Then
    
    Dim i As Byte
    Dim MiObj As obj
    
    For i = 1 To MAX_INVENTORY_SLOTS
    
        If npc.Invent.Object(i).ObjIndex > 0 Then
              MiObj.Amount = npc.Invent.Object(i).Amount
              MiObj.ObjIndex = npc.Invent.Object(i).ObjIndex
              
    'pluto:7.0
Notodo = RandomNumber(1, 100)
    If ObjData(MiObj.ObjIndex).Drop = 0 Then
    
         If MiObj.ObjIndex = 12 Then Notodo = Notodo - 30
        If MiObj.ObjIndex = 882 Then Notodo = 1
        If npc.GiveEXP < 1000 Then Notodo = 1
         If Notodo > 30 Then GoTo nada
         Call TirarItemAlPiso(npc.Pos, MiObj)
    Else
        If ObjData(MiObj.ObjIndex).Drop + 1 > Notodo Then Call TirarItemAlPiso(npc.Pos, MiObj)
    End If ' drop>0

End If 'items >0
nada:
Next i

End If

'pluto:2.19 ogro--> amuleto quitar paralisis 5% drop
If (npc.numero = 586 Or npc.numero = 691) And Notodo > 95 Then
MiObj.Amount = 1
MiObj.ObjIndex = 696
Call TirarItemAlPiso(npc.Pos, MiObj)
End If
'----------
'pluto:6.0A--> amuleto resu 5% sombra
If (npc.numero = 698) And Notodo > 95 Then
MiObj.Amount = 1
MiObj.ObjIndex = 697
Call TirarItemAlPiso(npc.Pos, MiObj)
End If
'----------
Dim alea As Integer
Dim alea2 As Integer
Dim ca As Byte

'pluto:6.0A raids tiran gemas
If npc.Raid > 0 Then
    
    If Notodo > 60 Then
MiObj.Amount = 1
MiObj.ObjIndex = RandomNumber(1202, 1206)
Call TirarItemAlPiso(npc.Pos, MiObj)
    End If

'tiramos objetos
'ca = Int(npc.Raid / 10)
'pluto:6.5
ca = (npc.numero - 699) + 4
'ca = RandomNumber(1, 5)
For alea2 = 1 To ca
alea = RandomNumber(1, Reo3)
MiObj.Amount = 1
MiObj.ObjIndex = ObjRegalo3(alea)
Call TirarItemAlPiso(npc.Pos, MiObj)
Next
'pluto:6.5 añado bolsas
    If Notodo + ca > 55 Then
     MiObj.Amount = 1
     MiObj.ObjIndex = 1244 'bolsa vida
    End If
End If ' raid

'objeto aleatorio de regalo

'pluto:2.4 cambio aleatorio por cofre y cambio sala invocacion por cofre
alea = RandomNumber(1, 350)
alea2 = RandomNumber(1, 500)
If alea > 348 And npc.GiveEXP > 4000 Then
MiObj.Amount = 1
MiObj.ObjIndex = 963
Call TirarItemAlPiso(npc.Pos, MiObj)
End If

'pluto:2.24 evento-----------------
If NumeroObjEvento > 0 Then
    If alea > 340 And npc.GiveEXP > 1000 Then
    MiObj.Amount = 1
    MiObj.ObjIndex = NumeroObjEvento
    Call TirarItemAlPiso(npc.Pos, MiObj)
    End If
End If
'----------------------------------


'pluto:2.22
If alea2 = 5 And npc.GiveEXP > 40000 Then
MiObj.Amount = 1
MiObj.ObjIndex = 40
Call TirarItemAlPiso(npc.Pos, MiObj)
ElseIf alea2 = 6 And npc.GiveEXP > 40000 Then
MiObj.Amount = 1
MiObj.ObjIndex = 41
Call TirarItemAlPiso(npc.Pos, MiObj)
ElseIf alea2 = 7 And npc.GiveEXP > 40000 Then
MiObj.Amount = 1
MiObj.ObjIndex = 961
Call TirarItemAlPiso(npc.Pos, MiObj)
End If
'-------------------------------

If npc.Pos.Map = mapi Then
MiObj.Amount = 1
MiObj.ObjIndex = 963
Call TirarItemAlPiso(npc.Pos, MiObj)
End If

'pluto:6.5 LOST CAIDO tira bolsa
If npc.numero = 705 And alea > 300 Then
MiObj.Amount = 1
MiObj.ObjIndex = 1251
Call TirarItemAlPiso(npc.Pos, MiObj)
End If

'pluto:6.0A devir--> diamante sangre
If npc.numero = 699 And alea = 349 Then
MiObj.Amount = 1
MiObj.ObjIndex = 1096
Call TirarItemAlPiso(npc.Pos, MiObj)
End If

'hilo de viudas negras
If npc.numero = 624 And alea = 349 Then
MiObj.Amount = 1
MiObj.ObjIndex = 1218
Call TirarItemAlPiso(npc.Pos, MiObj)
End If
'----------
'pluto:6.0A dragones--> huevo
If npc.NPCtype = DRAGON And alea = 349 Then
MiObj.Amount = 1
MiObj.ObjIndex = 1095
Call TirarItemAlPiso(npc.Pos, MiObj)
End If
'----------

'pluto:2.11 hobbit
If npc.numero = 623 And alea > 346 Then
MiObj.Amount = 1
MiObj.ObjIndex = 1015
Call TirarItemAlPiso(npc.Pos, MiObj)
End If
'pluto.2.14
If npc.numero = 631 And alea > 306 Then
MiObj.Amount = 1
MiObj.ObjIndex = 839
Call TirarItemAlPiso(npc.Pos, MiObj)
End If
If npc.numero = 659 And alea > 346 Then
MiObj.Amount = 1
MiObj.ObjIndex = 836
Call TirarItemAlPiso(npc.Pos, MiObj)
End If
'-----------------------------

If alea = 50 And npc.GiveEXP > 30000 Then
Dim ale2 As Integer
ale2 = RandomNumber(1, 6)
MiObj.Amount = 1
MiObj.ObjIndex = 985 + ale2
Call TirarItemAlPiso(npc.Pos, MiObj)
End If



'--------fin pluto:2.4-------

'pluto:regalos santa claus
If npc.NPCtype = 13 Then
MiObj.Amount = 1
MiObj.ObjIndex = 866
Call TirarItemAlPiso(npc.Pos, MiObj)
End If

Exit Sub
fallo:
Call LogError("NPCTIRARITEMS" & Err.number & " D: " & Err.Description)

End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo fallo
'Call LogTarea("Function QuedanItems npcindex:" & NpcIndex & " objindex:" & ObjIndex)

Dim i As Integer
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For i = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
            QuedanItems = True
            Exit Function
        End If
    Next
End If
QuedanItems = False

Exit Function
fallo:
Call LogError("QUEDANITEMS" & Err.number & " D: " & Err.Description)

End Function

Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
On Error GoTo fallo
'Devuelve la cantidad original del obj de un npc

Dim ln As String, npcfile As String
Dim i As Integer
If Npclist(NpcIndex).numero > 499 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "NPCs.dat"
End If
 
For i = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).numero, "Obj" & i)
    If ObjIndex = val(ReadField(1, ln, 45)) Then
        EncontrarCant = val(ReadField(2, ln, 45))
        Exit Function
    End If
Next
                   
EncontrarCant = 50
   Exit Function
fallo:
Call LogError("ENCONTRARCANT" & Err.number & " D: " & Err.Description)
                
End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
On Error GoTo fallo

Dim i As Integer

Npclist(NpcIndex).Invent.NroItems = 0

For i = 1 To MAX_INVENTORY_SLOTS
   Npclist(NpcIndex).Invent.Object(i).ObjIndex = 0
   Npclist(NpcIndex).Invent.Object(i).Amount = 0
Next i

Npclist(NpcIndex).InvReSpawn = 0
Exit Sub
fallo:
Call LogError("RESETNPCINV" & Err.number & " D: " & Err.Description)

End Sub

Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

On Error GoTo fallo

Dim ObjIndex As Integer
ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex

    'Quita un Obj
    If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Crucial = 0 Then
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    Else
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
            
            If Not QuedanItems(NpcIndex, ObjIndex) Then
                   
                   Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = ObjIndex
                   Npclist(NpcIndex).Invent.Object(Slot).Amount = EncontrarCant(NpcIndex, ObjIndex)
                   Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
            
            End If
            
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    
    
    
    End If
Exit Sub
fallo:
Call LogError("QUITARNPCINVITEM" & Err.number & " D: " & Err.Description)


End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
On Error GoTo fallo
'Vuelve a cargar el inventario del npc NpcIndex
Dim loopc As Integer
Dim ln As String

Dim npcfile As String

If Npclist(NpcIndex).numero > 499 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "NPCs.dat"
End If

Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & Npclist(NpcIndex).numero, "NROITEMS"))

For loopc = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).numero, "Obj" & loopc)
    Npclist(NpcIndex).Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))
    
Next loopc
Exit Sub
fallo:
Call LogError("CARGAR INVENT" & Err.number & " D: " & Err.Description)

End Sub


