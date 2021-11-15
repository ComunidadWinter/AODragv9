Attribute VB_Name = "ModMakro"
'********************************Modulo Makro*********************************
'Author: Matías Ignacio Rojo (MaxTus)
'Last Modification: 02/12/2011
'Control asistido de trabajo.
'******************************************************************************

Option Explicit

'******************************************************************************
'Requisitos para pescar
'******************************************************************************

Public Function PuedePescar(ByVal UserIndex As Integer) As Boolean

    Dim DummyINT As Integer

    With UserList(UserIndex)
    
        DummyINT = .Invent.WeaponEqpObjIndex
                
        If DummyINT = 0 Then
            .flags.Makro = 0
            Call WriteConsoleMsg(UserIndex, "Necesitas una caña o una red para atrapar peces. Dejas de trabajar...", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
            PuedePescar = False
            Exit Function
        End If
        
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.invisible)
            .flags.Makro = 0
            Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
            PuedePescar = False
            Exit Function
        End If
        
        If .Stats.UserSkills(eSkill.pesca) < 5 Then
            Call WriteConsoleMsg(UserIndex, "¡No tienes conocimientos en esa profesion!", FontTypeNames.FONTTYPE_INFOBOLD)
            .flags.Makro = 0
            Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
            PuedePescar = False
            Exit Function
        End If
        
        If .Stats.MinSta <= 5 Then
            Call WriteConsoleMsg(UserIndex, "Te encuentras demasiado cansado. Dejas de trabajar...", FontTypeNames.FONTTYPE_INFOBOLD)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
            .flags.Makro = 0
            Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
            PuedePescar = False
            Exit Function
        End If
                
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFOBOLD)
            .flags.Makro = 0
            Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
            PuedePescar = False
            Exit Function
        End If
        
        PuedePescar = True
        
    End With
    
End Function


'******************************************************************************
'Requisitos para lingotear
'******************************************************************************

Public Function PuedeLingotear(ByVal UserIndex As Integer) As Boolean

With UserList(UserIndex)

    If .flags.QueMontura Then
        Call WriteConsoleMsg(UserIndex, "No puedes fundir minerales estando montado... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFOBOLD)
        .flags.Makro = 0
        Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
        PuedeLingotear = False
        Exit Function
    End If
    
    If .flags.invisible = 1 Or .flags.Oculto = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.invisible)
        .flags.Makro = 0
        Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
        PuedeLingotear = False
        Exit Function
    End If
    
    If .Stats.UserSkills(eSkill.Mineria) < 5 Then
        Call WriteConsoleMsg(UserIndex, "¡No tienes conocimientos en esa profesion!", FontTypeNames.FONTTYPE_INFOBOLD)
        .flags.Makro = 0
        Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
        PuedeLingotear = False
        Exit Function
    End If
                
    'Check there is a proper item there
    If .flags.targetObj > 0 Then
        If ObjData(.flags.targetObj).OBJType = eOBJType.otFragua Then
        'Validate other items
            If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MAX_INVENTORY_SLOTS Then
                Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en tu inventario... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFOBOLD)
                .flags.Makro = 0
                Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
                PuedeLingotear = False
                Exit Function
            End If
                        
            ''chequeamos que no se zarpe duplicando oro
            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                    .flags.Makro = 0
                    Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
                    PuedeLingotear = False
                    Exit Function
                End If
                            
                            ''FUISTE
                Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Function
            End If
            
            'Puede trabajar ;)
            PuedeLingotear = True
            
        Else
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fragua allí... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFOBOLD)
            .flags.Makro = 0
            Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
            PuedeLingotear = False
            Exit Function
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "No hay ninguna fragua allí... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFOBOLD)
        .flags.Makro = 0
        Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
        PuedeLingotear = False
        Exit Function
    End If
            
End With

End Function

'******************************************************************************
'Inicia la actividad
'******************************************************************************

Public Sub MakroTrabajo(ByVal UserIndex As Integer, ByRef Tarea As eMakro)
        
        Select Case Tarea
            Case eMakro.PESCAR
                If PuedePescar(UserIndex) Then
                    DoPescar UserIndex
                End If
                
            Case eMakro.PescarRed
                If PuedePescar(UserIndex) Then
                    DoPescarRed UserIndex
                End If
                
            Case eMakro.Lingotear
                If PuedeLingotear(UserIndex) Then
                    FundirMineral UserIndex
                End If
        End Select
        
End Sub

