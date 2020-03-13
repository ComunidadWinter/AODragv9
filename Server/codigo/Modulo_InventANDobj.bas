Attribute VB_Name = "InvNpc"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


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
Public Function TirarItemAlPiso(Pos As WorldPos, Obj As Obj, Optional NotPirata As Boolean = True) As WorldPos
On Error GoTo Errhandler
    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    
    Tilelibre Pos, NuevaPos, Obj, NotPirata, True
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call MakeObj(Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
    End If
    TirarItemAlPiso = NuevaPos

Exit Function
Errhandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc, ByVal UserIndex As Integer)
'TIRA TODOS LOS ITEMS DEL NPC
On Error Resume Next
Dim Prob As Byte
Dim cant As Integer

If npc.Invent.NroItems > 0 Then
    
    Dim i As Byte
    Dim MiObj As Obj
    
    For i = 1 To MAX_INVENTORY_SLOTS
        If npc.Invent.Object(i).ObjIndex > 0 Then
            Prob = RandomNumber(1, 100)
            
            If Prob <= npc.Invent.Object(i).ProbTirar Then
                MiObj.Amount = npc.Invent.Object(i).Amount
                MiObj.ObjIndex = npc.Invent.Object(i).ObjIndex
                
                '¿Estaba trabajando?
                If UserTrabajando = True Then
                    'Le damos un porcentaje en funcion a sus skill y un randomnumber
                    cant = RandomNumber(Porcentaje(MiObj.Amount, UserList(UserIndex).Stats.UserSkills(SkillTrabajando)), 1)
                    
                    'Nunca podra dar 0 del material
                    If cant = 0 Then
                        MiObj.Amount = 1
                    Else
                        MiObj.Amount = cant
                    End If
                    
                    UserTrabajando = False
                    SkillTrabajando = 0
                End If
                    
                Call TirarItemAlPiso(npc.Pos, MiObj)
                
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.DROPPROB)
            End If
        End If
    Next i

End If

End Sub

Function QuedanItems(ByVal NPCIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error Resume Next
'Call LogTarea("Function QuedanItems npcindex:" & NpcIndex & " objindex:" & ObjIndex)

Dim i As Integer
If NPCList(NPCIndex).Invent.NroItems > 0 Then
    For i = 1 To MAX_INVENTORY_SLOTS
        If NPCList(NPCIndex).Invent.Object(i).ObjIndex = ObjIndex Then
            QuedanItems = True
            Exit Function
        End If
    Next
End If
QuedanItems = False
End Function

''
' Gets the amount of a certain item that an npc has.
'
' @param npcIndex Specifies reference to npcmerchant
' @param ObjIndex Specifies reference to object
' @return   The amount of the item that the npc has
' @remarks This function reads the Npc.dat file
Function EncontrarCant(ByVal NPCIndex As Integer, ByVal ObjIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 03/09/08
'Last Modification By: Marco Vanotti (Marco)
' - 03/09/08 EncontrarCant now returns 0 if the npc doesn't have it (Marco)
'***************************************************
On Error Resume Next
'Devuelve la cantidad original del obj de un npc

Dim ln As String, npcfile As String
Dim i As Integer

npcfile = DatPath & "NPCs.dat"
 
For i = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(npcfile, "NPC" & NPCList(NPCIndex).Numero, "Obj" & i)
    If ObjIndex = val(ReadField(1, ln, 45)) Then
        EncontrarCant = val(ReadField(2, ln, 45))
        Exit Function
    End If
Next
                   
EncontrarCant = 0

End Function

Sub ResetNpcInv(ByVal NPCIndex As Integer)
On Error Resume Next

Dim i As Integer

NPCList(NPCIndex).Invent.NroItems = 0

For i = 1 To MAX_INVENTORY_SLOTS
   NPCList(NPCIndex).Invent.Object(i).ObjIndex = 0
   NPCList(NPCIndex).Invent.Object(i).Amount = 0
Next i

NPCList(NPCIndex).InvReSpawn = 0

End Sub

''
' Removes a certain amount of items from a slot of an npc's inventory
'
' @param npcIndex Specifies reference to npcmerchant
' @param Slot Specifies reference to npc's inventory's slot
' @param antidad Specifies amount of items that will be removed
Sub QuitarNpcInvItem(ByVal NPCIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 03/09/08
'Last Modification By: Marco Vanotti (Marco)
' - 03/09/08 Now this sub checks that te npc has an item before respawning it (Marco)
'***************************************************
Dim ObjIndex As Integer
Dim iCant As Integer

With NPCList(NPCIndex)

    ObjIndex = .Invent.Object(Slot).ObjIndex
    
        'Quita un Obj
        If ObjData(.Invent.Object(Slot).ObjIndex).Crucial = 0 And ObjData(.Invent.Object(Slot).ObjIndex).NoLimpiar = 0 Then
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            
            If .Invent.Object(Slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.Object(Slot).Amount = 0
                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                   Call CargarInvent(NPCIndex) 'Reponemos el inventario
                End If
            End If
        Else
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            
            If .Invent.Object(Slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.Object(Slot).Amount = 0
                
                If Not QuedanItems(NPCIndex, ObjIndex) Then
                    'Check if the item is in the npc's dat.
                    iCant = EncontrarCant(NPCIndex, ObjIndex)
                    If iCant Then
                        .Invent.Object(Slot).ObjIndex = ObjIndex
                        .Invent.Object(Slot).Amount = iCant
                        .Invent.NroItems = .Invent.NroItems + 1
                    End If
                End If
                
                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                   Call CargarInvent(NPCIndex) 'Reponemos el inventario
                End If
            End If
        End If
   End With
End Sub

Sub CargarInvent(ByVal NPCIndex As Integer)

'Vuelve a cargar el inventario del npc NpcIndex
Dim LoopC As Integer
Dim ln As String
Dim npcfile As String

npcfile = DatPath & "NPCs.dat"

NPCList(NPCIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NPCList(NPCIndex).Numero, "NROITEMS"))

For LoopC = 1 To NPCList(NPCIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & NPCList(NPCIndex).Numero, "Obj" & LoopC)
    NPCList(NPCIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    NPCList(NPCIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
Next LoopC

End Sub


