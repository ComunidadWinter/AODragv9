Attribute VB_Name = "modUsuariosInv"
Option Explicit

Public Sub MoveItem(ByVal UserIndex As Integer, ByVal originalSlot As Integer, ByVal newSlot As Integer) ' 0.13.3
'***************************************************
'Author: Unknownn
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************

Dim tmpObj As UserOBJ
Dim newObjIndex As Integer, originalObjIndex As Integer
If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub

With UserList(UserIndex)
    If (originalSlot > MAX_INVENTORY_SLOTS) Or (newSlot > MAX_INVENTORY_SLOTS) Then Exit Sub

    tmpObj = .Invent.Object(originalSlot)
    .Invent.Object(originalSlot) = .Invent.Object(newSlot)
    .Invent.Object(newSlot) = tmpObj
    
    If .Invent.AnilloEqpSlot = originalSlot Then
        .Invent.AnilloEqpSlot = newSlot
    ElseIf .Invent.AnilloEqpSlot = newSlot Then
        .Invent.AnilloEqpSlot = originalSlot
    End If
    
    If .Invent.ArmourEqpSlot = originalSlot Then
        .Invent.ArmourEqpSlot = newSlot
    ElseIf .Invent.ArmourEqpSlot = newSlot Then
        .Invent.ArmourEqpSlot = originalSlot
    End If
    
    If .Invent.BarcoSlot = originalSlot Then
        .Invent.BarcoSlot = newSlot
    ElseIf .Invent.BarcoSlot = newSlot Then
        .Invent.BarcoSlot = originalSlot
    End If
    
    If .Invent.CascoEqpSlot = originalSlot Then
         .Invent.CascoEqpSlot = newSlot
    ElseIf .Invent.CascoEqpSlot = newSlot Then
         .Invent.CascoEqpSlot = originalSlot
    End If
    
    If .Invent.EscudoEqpSlot = originalSlot Then
        .Invent.EscudoEqpSlot = newSlot
    ElseIf .Invent.EscudoEqpSlot = newSlot Then
        .Invent.EscudoEqpSlot = originalSlot
    End If
    
    If .Invent.MunicionEqpSlot = originalSlot Then
        .Invent.MunicionEqpSlot = newSlot
    ElseIf .Invent.MunicionEqpSlot = newSlot Then
        .Invent.MunicionEqpSlot = originalSlot
    End If
    
    If .Invent.WeaponEqpSlot = originalSlot Then
        .Invent.WeaponEqpSlot = newSlot
    ElseIf .Invent.WeaponEqpSlot = newSlot Then
        .Invent.WeaponEqpSlot = originalSlot
    End If

    Call UpdateUserInv(False, UserIndex, originalSlot)
    Call UpdateUserInv(False, UserIndex, newSlot)
End With
End Sub

