Attribute VB_Name = "modHechizos"
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

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 13/02/2009
'13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
'***************************************************
If NPCList(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

' Si no se peude usar magia en el mapa, no le deja hacerlo.
If MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto > 0 Then Exit Sub

NPCList(NpcIndex).CanAttack = 0
Dim daño As Integer

If Hechizos(Spell).SubeHP = 1 Then

    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + daño
    If UserList(UserIndex).Stats.MinHP > VidaMaxima(UserIndex) Then UserList(UserIndex).Stats.MinHP = VidaMaxima(UserIndex)
    
    Call WriteConsoleMsg(UserIndex, NPCList(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteUpdateUserStats(UserIndex)

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        'Casco
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    
        'Armadura
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then _
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        
        'Arma
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).DefensaMagicaMax)
        
        'Escudo
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).DefensaMagicaMax)
        
        'Anillo
        If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then _
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        
        '<Edurne> 'En montura...
        If UserList(UserIndex).flags.QueMontura Then _
            daño = daño - UserList(UserIndex).flags.Montura(UserList(UserIndex).flags.QueMontura).DefMagia
        
        If daño < 0 Then daño = 0
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
    
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
                
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, daño, COLOR_DAÑO))
        
        Call WriteConsoleMsg(UserIndex, NPCList(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteUpdateUserStats(UserIndex)
        
        'Muere
        If UserList(UserIndex).Stats.MinHP < 1 Then
            UserList(UserIndex).Stats.MinHP = 0
            If NPCList(NpcIndex).NPCType = eNPCType.GuardiaReal Then
                RestarCriminalidad (UserIndex)
            End If
            Call UserDie(UserIndex)
            '[Barrin 1-12-03]
            If NPCList(NpcIndex).MaestroUser > 0 Then
                'Store it!
                Call Statistics.StoreFrag(NPCList(NpcIndex).MaestroUser, UserIndex)
                
                Call ContarMuerte(UserIndex, NPCList(NpcIndex).MaestroUser)
                Call ActStats(UserIndex, NPCList(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Or Hechizos(Spell).Inmoviliza = 1 Then
    If UserList(UserIndex).flags.Paralizado = 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
          
        If UserList(UserIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
            Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If Hechizos(Spell).Inmoviliza = 1 Then
            UserList(UserIndex).flags.Inmovilizado = 1
        End If
          
        UserList(UserIndex).flags.Paralizado = 1
        UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
          
        Call WriteParalizeOK(UserIndex)
    End If
End If

If Hechizos(Spell).Estupidez = 1 Then   ' turbacion
     If UserList(UserIndex).flags.Estupidez = 0 Then
          Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
          Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
          
            If UserList(UserIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
          
          UserList(UserIndex).flags.Estupidez = 1
          UserList(UserIndex).Counters.Ceguera = 80
                  
        Call WriteDumb(UserIndex)
     End If
End If

End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If NPCList(NpcIndex).CanAttack = 0 Then Exit Sub
NPCList(NpcIndex).CanAttack = 0

Dim daño As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV, NPCList(TargetNPC).Pos.X, NPCList(TargetNPC).Pos.Y))
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(NPCList(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        
        NPCList(TargetNPC).Stats.MinHP = NPCList(TargetNPC).Stats.MinHP - daño
        
        'Muere
        If NPCList(TargetNPC).Stats.MinHP < 1 Then
            NPCList(TargetNPC).Stats.MinHP = 0
            If NPCList(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, NPCList(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

On Error GoTo errHandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errHandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call WriteConsoleMsg(UserIndex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub DecirPalabrasMagicas(ByVal SpellWords As String, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 25/07/2009
'25/07/2009: ZaMa - Invisible admins don't say any word when casting a spell
'***************************************************
On Error Resume Next
    If UserList(UserIndex).flags.AdminInvisible <> 1 Then _
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(SpellWords, UserList(UserIndex).Char.CharIndex, vbCyan))
    Exit Sub
End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 06/11/09
'Last Modification By: Torres Patricio (Pato)
' - 06/11/09 Corregida la bonificación de maná del mimetismo en el druida con flauta mágica equipada.
'***************************************************
Dim DruidManaBonus As Single

    If UserList(UserIndex).flags.Muerto Then
        Call WriteMultiMessage(UserIndex, eMessages.Muerto)
        PuedeLanzar = False
        Exit Function
    End If

    '¿Está trabajando?
    If UserList(UserIndex).flags.Makro <> 0 Then
        Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
        
    If UserList(UserIndex).Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
        Call WriteConsoleMsg(UserIndex, "No tenes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
        If UserList(UserIndex).genero = eGenero.Hombre Then
            Call WriteConsoleMsg(UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "Estás muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        End If
        PuedeLanzar = False
        Exit Function
    End If

    DruidManaBonus = 1

    
    If UserList(UserIndex).Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido * DruidManaBonus Then
        Call WriteMultiMessage(UserIndex, eMessages.NoMana)
        PuedeLanzar = False
        Exit Function
    End If
        
    PuedeLanzar = True
End Function

Sub HechizoporArea(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    '// HECHIZOS POR ÁREA //
    For TempX = PosCasteadaX - 1 To PosCasteadaX + 1
        For TempY = PosCasteadaY - 1 To PosCasteadaY + 1
            If MapData(PosCasteadaM, TempX, TempY).NpcIndex > 0 Then
                Call HechizoPropNPC(H, MapData(PosCasteadaM, TempX, TempY).NpcIndex, UserIndex, False)
            End If
           
            'Lo pongo así porque este sub da ASCO
            If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                Call HechizoPropUsuario(UserIndex, False)
            End If
        Next TempY
    Next TempX
    '// HECHIZOS POR ÁREA //
    b = True
End Sub

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(UserIndex)
    End If

End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Author: Uknown
'Last modification: 06/15/2008 (NicoNZ)
'Sale del sub si no hay una posición valida.
'***************************************************
If UserList(UserIndex).NroMascotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call WriteConsoleMsg(UserIndex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As worldPos


TargetPos.Map = UserList(UserIndex).flags.TargetMap
TargetPos.X = UserList(UserIndex).flags.TargetX
TargetPos.Y = UserList(UserIndex).flags.TargetY

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).cant
    
    If UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1
            
            Index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(Index) = ind
            UserList(UserIndex).MascotasType(Index) = NPCList(ind).Numero
            
            NPCList(ind).MaestroUser = UserIndex
            NPCList(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            NPCList(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        Else
            Exit Sub
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(UserIndex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/01/08
'
'***************************************************

Dim b As Boolean

Select Case Hechizos(uh).tipo
    Case TipoHechizo.uInvocacion '
        Call HechizoInvocacion(UserIndex, b)
    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(UserIndex, b)
    Case TipoHechizo.uArea
        Call HechizoporArea(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(UserIndex)
End If

End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/01/08
'
'***************************************************
Dim b As Boolean

Select Case Hechizos(uh).tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(UserIndex, b)
    
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(UserIndex, b)
       
    Case TipoHechizo.uArea
        Call HechizoporArea(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    'Agregado para que los druidas, al tener equipada la flauta magica, el coste de mana de mimetismo es de 50% menos.
    If UserList(UserIndex).clase = eClass.Druid And Hechizos(uh).Mimetiza = 1 Then
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido * 0.5
    Else
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    End If
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(UserIndex)
    Call WriteUpdateUserStats(UserList(UserIndex).flags.targetUser)
    UserList(UserIndex).flags.targetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/02/2009
'13/02/2009: ZaMa - Agregada 50% bonificacion en coste de mana a mimetismo para druidas
'***************************************************
Dim b As Boolean

Select Case Hechizos(uh).tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)
    Case TipoHechizo.uArea
        Call HechizoporArea(UserIndex, b)
End Select


If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNPC = 0
    
    ' Bonificación para druidas.
    If UserList(UserIndex).clase = eClass.Druid And Hechizos(uh).Mimetiza = 1 Then
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido * 0.5
    Else
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    End If

    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(UserIndex)
End If

End Sub


Sub LanzarHechizo(Index As Integer, UserIndex As Integer)

On Error GoTo errHandler

Dim uh As Integer

uh = UserList(UserIndex).Stats.UserHechizos(Index)

If PuedeLanzar(UserIndex, uh) Then
    Select Case Hechizos(uh).Target
        Case TargetType.uUsuarios
            If UserList(UserIndex).flags.targetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.targetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Este hechizo actúa solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uNPC
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(NPCList(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uUsuariosYnpc
            If UserList(UserIndex).flags.targetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.targetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                End If
            ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(NPCList(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call WriteMultiMessage(UserIndex, eMessages.Lejos)
                End If
            End If
        
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
    
End If

If UserList(UserIndex).Counters.Trabajando Then _
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

If UserList(UserIndex).Counters.Ocultando Then _
    UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    
Exit Sub

errHandler:
    Dim UserNick As String
    
    If UserIndex > 0 Then UserNick = UserList(UserIndex).Name

    Call LogError("Error en LanzarHechizo. Error " & Err.Number & " : " & Err.Description & " UserIndex: " & UserIndex & " Nick: " & UserNick)
    
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 13/02/2009
'Handles the Spells that afect the Stats of an User
'24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
'26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
'26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
'02/01/2008 Marcos (ByVal) - Curar Veneno no permitido en usuarios muertos.
'06/28/2008 NicoNZ - Agregué que se le de valor al flag Inmovilizado.
'17/11/2008: NicoNZ - Agregado para quitar la penalización de vida en el ring y cambio de ecuacion.
'13/02/2009: ZaMa - Arreglada ecuacion para quitar vida tras resucitar en rings.
'***************************************************


Dim H As Integer, tU As Integer
Dim UserProtected As Boolean

    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    tU = UserList(UserIndex).flags.targetUser

    UserProtected = Not IntervaloPermiteSerAtacado(tU) And UserList(tU).flags.NoPuedeSerAtacado
    
    If UserProtected Then b = False: Exit Sub

    If Hechizos(H).Invisibilidad = 1 Then
       
        If UserList(tU).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        If UserList(tU).Counters.Saliendo Then
            If UserIndex <> tU Then
                Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "¡No puedes ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
                b = False
                Exit Sub
            End If
        End If
        
        'No usar invi mapas InviSinEfecto
        '12/12/2018 Irongete: Comprobar si la zona permite la invisibilidad
        If permiso_en_zona(UserIndex) And permiso_zona.no_invisibilidad Then
          Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona en esta zona!", FontTypeNames.FONTTYPE_INFO)
          b = False
          Exit Sub
        End If

        
        'Para poder tirar invi a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If criminal(tU) And Not criminal(UserIndex) Then
                If esArmada(UserIndex) Or esLegion(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real y Legion no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)
                End If
            End If
        End If
        
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
            If Not UserList(tU).flags.Privilegios And PlayerType.User Then
                Exit Sub
            End If
        End If
       
        UserList(tU).flags.invisible = 1
        Call SetInvisible(tU, UserList(tU).Char.CharIndex, True)
        'Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))
    
        Call InfoHechizo(UserIndex)
        b = True
    End If
    
    If Hechizos(H).Mimetiza = 1 Then
        If UserList(tU).flags.Muerto = 1 Then
            Exit Sub
        End If
        
        If UserList(tU).flags.Navegando = 1 Then
            Exit Sub
        End If
        If UserList(UserIndex).flags.Navegando = 1 Then
            Exit Sub
        End If
        
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
            If Not UserList(tU).flags.Privilegios And PlayerType.User Then
                Exit Sub
            End If
        End If
        
        If UserList(UserIndex).flags.Mimetizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
        
        'copio el char original al mimetizado
        
        With UserList(UserIndex)
            .CharMimetizado.body = .Char.body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
            .flags.Mimetizado = 1
            
            'ahora pongo local el del enemigo
            .Char.body = UserList(tU).Char.body
            .Char.Head = UserList(tU).Char.Head
            .Char.CascoAnim = UserList(tU).Char.CascoAnim
            .Char.ShieldAnim = UserList(tU).Char.ShieldAnim
            .Char.WeaponAnim = UserList(tU).Char.WeaponAnim
        
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End With
       
       Call InfoHechizo(UserIndex)
       b = True
    End If
    
    If Hechizos(H).Envenena = 1 Then
        If UserIndex = tU Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).flags.Envenenado = 1
        Call InfoHechizo(UserIndex)
        ' No podras pasar de mapa por un rato
        Call IntervaloPermiteCambiardeMapa(UserIndex, True)
        Call IntervaloPermiteCambiardeMapa(tU, True)
        b = True
    End If
    
    If Hechizos(H).CuraVeneno = 1 Then
    
        'Verificamos que el usuario no este muerto
        If UserList(tU).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'Para poder tirar curar veneno a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If criminal(tU) And Not criminal(UserIndex) Then
                If esArmada(UserIndex) Or esLegion(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Los Armadas y Legión no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If
            End If
        End If
            
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
            If Not UserList(tU).flags.Privilegios And PlayerType.User Then
                Exit Sub
            End If
        End If
            
        UserList(tU).flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        b = True
    End If
    
    If Hechizos(H).Maldicion = 1 Then
        If UserIndex = tU Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        ' No podras pasar de mapa por un rato
        Call IntervaloPermiteCambiardeMapa(UserIndex, True)
        Call IntervaloPermiteCambiardeMapa(tU, True)
        b = True
    End If
    
    If Hechizos(H).RemoverMaldicion = 1 Then
            UserList(tU).flags.Maldicion = 0
            Call InfoHechizo(UserIndex)
            b = True
    End If
    
    If Hechizos(H).Bendicion = 1 Then
            UserList(tU).flags.Bendicion = 1
            Call InfoHechizo(UserIndex)
            b = True
    End If
    
    If Hechizos(H).Paraliza = 1 Or Hechizos(H).Inmoviliza = 1 Then
        If UserIndex = tU Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
         If UserList(tU).flags.Paralizado = 0 Then
                If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
                
                If UserIndex <> tU Then
                    Call UsuarioAtacadoPorUsuario(UserIndex, tU)
                End If
                
                Call InfoHechizo(UserIndex)
                b = True
                If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                    Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                    Call FlushBuffer(tU)
                    Exit Sub
                End If
                
                If Hechizos(H).Inmoviliza = 1 Then UserList(tU).flags.Inmovilizado = 1
                UserList(tU).flags.Paralizado = 1
                UserList(tU).Counters.Paralisis = IntervaloParalizado
                
                Call WriteParalizeOK(tU)
                ' No podras pasar de mapa por un rato
                Call IntervaloPermiteCambiardeMapa(UserIndex, True)
                Call IntervaloPermiteCambiardeMapa(tU, True)
                
                Call FlushBuffer(tU)
          
        End If
    End If
    
    
    If Hechizos(H).RemoverParalisis = 1 Then
        If UserList(tU).flags.Paralizado = 1 Then
            'Para poder tirar remo a un pk en el ring
            If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
                If criminal(tU) And Not criminal(UserIndex) Then
                    If esArmada(UserIndex) Or esLegion(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "Los Armadas y Legión no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                    If UserList(UserIndex).flags.Seguro = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    Else
                        Call VolverCriminal(UserIndex)
                    End If
                End If
            End If
            
            UserList(tU).flags.Inmovilizado = 0
            UserList(tU).flags.Paralizado = 0
            'no need to crypt this
            Call WriteParalizeOK(tU)
            Call InfoHechizo(UserIndex)
            b = True
        End If
    End If
    
    If Hechizos(H).RemoverEstupidez = 1 Then
        If UserList(tU).flags.Estupidez = 1 Then
            'Para poder tirar remo estu a un pk en el ring
            If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
                If criminal(tU) And Not criminal(UserIndex) Then
                    If esArmada(UserIndex) Or esLegion(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "Los Armadas o Legión no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                    If UserList(UserIndex).flags.Seguro = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    Else
                        Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If
                End If
            End If
        
            UserList(tU).flags.Estupidez = 0
            'no need to crypt this
            Call WriteDumbNoMore(tU)
            Call FlushBuffer(tU)
            Call InfoHechizo(UserIndex)
            b = True
        End If
    End If
    
    
    If Hechizos(H).Revivir = 1 Then
        If UserList(tU).flags.Muerto = 1 Then
        
            'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
            If UserList(tU).flags.SeguroResu Then
                Call WriteConsoleMsg(UserIndex, "¡El espíritu no tiene intenciones de regresar al mundo de los vivos!", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            
            'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
            If UserList(tU).flags.SeguroResu Then
                Call WriteConsoleMsg(UserIndex, "¡El espíritu no tiene intenciones de regresar al mundo de los vivos!", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        
            'No usar resu en mapas con ResuSinEfecto
            If MapInfo(UserList(tU).Pos.Map).ResuSinEfecto > 0 Then
                Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            
            'No podemos resucitar si nuestra barra de energía no está llena. (GD: 29/04/07)
            If UserList(UserIndex).Stats.MaxSta <> UserList(UserIndex).Stats.MinSta Then
                Call WriteConsoleMsg(UserIndex, "No puedes resucitar si no tienes tu barra de energía llena.", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            
            'Para poder tirar revivir a un pk en el ring
            If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
                If criminal(tU) And Not criminal(UserIndex) Then
                    If esArmada(UserIndex) Or esLegion(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "Los Armadas y Legión no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                    If UserList(UserIndex).flags.Seguro = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    Else
                        Call VolverCriminal(UserIndex)
                    End If
                End If
            End If
    
            '16/02/2016 Lorwik> Si es CAOS no gana nobleza es malo...
            If esCaos(UserIndex) = False Then
                Dim EraCriminal As Boolean
                EraCriminal = criminal(UserIndex)
                If Not criminal(tU) Then
                    If tU <> UserIndex Then
                        UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + 500
                        If UserList(UserIndex).Reputacion.NobleRep > MAXREP Then _
                            UserList(UserIndex).Reputacion.NobleRep = MAXREP
                        Call WriteConsoleMsg(UserIndex, "¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
                If EraCriminal And Not criminal(UserIndex) Then
                    Call RefreshCharStatus(UserIndex)
                End If
            End If
            
            'Pablo Toxic Waste (GD: 29/04/07)
            UserList(tU).Stats.MinAGU = 0
            UserList(tU).flags.Sed = 1
            UserList(tU).Stats.MinHam = 0
            UserList(tU).flags.Hambre = 1
            Call WriteUpdateHungerAndThirst(tU)
            Call InfoHechizo(UserIndex)
            UserList(tU).Stats.MinMAN = 0
            UserList(tU).Stats.MinSta = 0
            
            'Agregado para quitar la penalización de vida en el ring y cambio de ecuacion. (NicoNZ)
            If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
                'Solo saco vida si es User. no quiero que exploten GMs por ahi.
                If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
                    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP * (1 - UserList(tU).Stats.ELV * 0.015)
                End If
            End If
            
            If (UserList(UserIndex).Stats.MinHP <= 0) Then
                Call UserDie(UserIndex)
                Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande", FontTypeNames.FONTTYPE_INFO)
                b = False
            Else
                Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado", FontTypeNames.FONTTYPE_INFO)
                b = True
            End If
            
            Call RevivirUsuario(tU)
        Else
            b = False
        End If
    
    End If
    
    If Hechizos(H).Ceguera = 1 Then
        If UserIndex = tU Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
            If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            If UserIndex <> tU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
            UserList(tU).flags.Ceguera = 1
            UserList(tU).Counters.Ceguera = 80
    
            ' No podras pasar de mapa por un rato
            Call IntervaloPermiteCambiardeMapa(UserIndex, True)
            Call IntervaloPermiteCambiardeMapa(tU, True)
        
            Call WriteBlind(tU)
            Call FlushBuffer(tU)
            Call InfoHechizo(UserIndex)
            b = True
    End If
    
    If Hechizos(H).Estupidez = 1 Then
        If UserIndex = tU Then
            Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
            If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            If UserIndex <> tU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
            If UserList(tU).flags.Estupidez = 0 Then
                UserList(tU).flags.Estupidez = 1
                UserList(tU).Counters.Ceguera = 80
            End If
            ' No podras pasar de mapa por un rato
            Call IntervaloPermiteCambiardeMapa(UserIndex, True)
            Call IntervaloPermiteCambiardeMapa(tU, True)
            
            Call WriteDumb(tU)
            Call FlushBuffer(tU)
    
            Call InfoHechizo(UserIndex)
            b = True
    End If

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 07/07/2008
'Handles the Spells that afect the Stats of an NPC
'04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
'removidos por users de su misma faccion.
'07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
'***************************************************
If Hechizos(hIndex).Invisibilidad = 1 Then
    Call InfoHechizo(UserIndex)
    NPCList(NpcIndex).flags.invisible = 1
    b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, UserIndex)
    Call InfoHechizo(UserIndex)
    NPCList(NpcIndex).flags.Envenenado = 1
    b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
    Call InfoHechizo(UserIndex)
    NPCList(NpcIndex).flags.Envenenado = 0
    b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, UserIndex)
    Call InfoHechizo(UserIndex)
    NPCList(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
    Call InfoHechizo(UserIndex)
    NPCList(NpcIndex).flags.Maldicion = 0
    b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
    Call InfoHechizo(UserIndex)
    NPCList(NpcIndex).flags.Bendicion = 1
    b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If NPCList(NpcIndex).flags.AfectaParalisis = 0 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        Call InfoHechizo(UserIndex)
        NPCList(NpcIndex).flags.Paralizado = 1
        NPCList(NpcIndex).flags.Inmovilizado = 0
        NPCList(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        ' No podras pasar de mapa por un rato
        Call IntervaloPermiteCambiardeMapa(UserIndex, True)
        b = True
    Else
        Call WriteConsoleMsg(UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
End If

If Hechizos(hIndex).RemoverParalisis = 1 Then
    If NPCList(NpcIndex).flags.Paralizado = 1 Or NPCList(NpcIndex).flags.Inmovilizado = 1 Then
        If NPCList(NpcIndex).MaestroUser = UserIndex Then
            Call InfoHechizo(UserIndex)
            NPCList(NpcIndex).flags.Paralizado = 0
            NPCList(NpcIndex).Contadores.Paralisis = 0
            b = True
        Else
            If NPCList(NpcIndex).NPCType = eNPCType.GuardiaReal Then
                If esArmada(UserIndex) Then
                    Call InfoHechizo(UserIndex)
                    NPCList(NpcIndex).flags.Paralizado = 0
                    NPCList(NpcIndex).Contadores.Paralisis = 0
                    b = True
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                
                Call WriteConsoleMsg(UserIndex, "Solo puedes Remover la Parálisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                If NPCList(NpcIndex).NPCType = eNPCType.Guardiascaos Then
                    If esCaos(UserIndex) Then
                        Call InfoHechizo(UserIndex)
                        NPCList(NpcIndex).flags.Paralizado = 0
                        NPCList(NpcIndex).Contadores.Paralisis = 0
                        b = True
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                End If
            End If
        End If
   Else
      Call WriteConsoleMsg(UserIndex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFO)
      b = False
      Exit Sub
   End If
End If
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If NPCList(NpcIndex).flags.AfectaParalisis = 0 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        NPCList(NpcIndex).flags.Inmovilizado = 1
        NPCList(NpcIndex).flags.Paralizado = 0
        NPCList(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(UserIndex)
        ' No podras pasar de mapa por un rato
        Call IntervaloPermiteCambiardeMapa(UserIndex, True)
        b = True
    Else
        Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

If Hechizos(hIndex).Mimetiza = 1 Then
    
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
    
        
    If UserList(UserIndex).clase = eClass.Druid Then
        'copio el char original al mimetizado
        With UserList(UserIndex)
            .CharMimetizado.body = .Char.body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
            .flags.Mimetizado = 1
            
            'ahora pongo lo del NPC.
            .Char.body = NPCList(NpcIndex).Char.body
            .Char.Head = NPCList(NpcIndex).Char.Head
            .Char.CascoAnim = NingunCasco
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
        
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End With
    Else
        Call WriteConsoleMsg(UserIndex, "Solo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

   Call InfoHechizo(UserIndex)
   b = True
End If
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean, Optional ByVal xArea As Boolean = False)
    Dim daño As Long
    Dim i As Byte
    Dim n As Long

    'LORWIK> PUERTA DEL CASTILLO
    '(SE PUEDE HACER MEJOR, PERO ASI NOS VALE)
    'CODIGO ORIGINAL: AODRAG 7
    If NPCList(NpcIndex).NPCType = 10 Then
    
        Dim nPos As worldPos
        nPos.Map = NPCList(NpcIndex).Pos.Map
        nPos.X = NPCList(NpcIndex).Pos.X
        nPos.Y = NPCList(NpcIndex).Pos.Y
        'pluto:6.0A-----------------
        If Hechizos(hIndex).SubeHP = 1 And nPos.Y > UserList(UserIndex).Pos.Y Then
            Call WriteConsoleMsg(UserIndex, "No puedes restaurar la puerta desde este lado.", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
        '----------------------------
        
    End If

    If NPCList(NpcIndex).NPCType = Arbol Or NPCList(NpcIndex).NPCType = Yacimiento Then
        Call WriteConsoleMsg(UserIndex, "No puedes trabajar de ese modo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    'Salud
    If Hechizos(hIndex).SubeHP = 1 Then
        daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
        daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
        Call InfoHechizo(UserIndex, xArea)
        NPCList(NpcIndex).Stats.MinHP = NPCList(NpcIndex).Stats.MinHP + daño
        If NPCList(NpcIndex).Stats.MinHP > NPCList(NpcIndex).Stats.MaxHP Then _
            NPCList(NpcIndex).Stats.MinHP = NPCList(NpcIndex).Stats.MaxHP
        Call WriteConsoleMsg(UserIndex, Hechizos(hIndex).nombre & " cura " & daño & " a " & NPCList(NpcIndex).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateRenderValue(NPCList(NpcIndex).Pos.X, NPCList(NpcIndex).Pos.Y, Abs(daño), COLOR_CURACION))
        b = True
    
    ElseIf Hechizos(hIndex).SubeHP = 2 Then
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, UserIndex)
    
    If NPCList(NpcIndex).flags.SoloParty = 1 And UserList(UserIndex).PartyId = 0 Then
        Call WriteConsoleMsg(UserIndex, "Para atacar a esta criatura necesitas pertenecer a una Party.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    'cascos antimagia
    If (UserList(UserIndex).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DañoMagico
    End If
    
    'anillos
    If (UserList(UserIndex).Invent.AnilloEqpObjIndex > 0) Then
        daño = daño - ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DañoMagico
    End If
    
    'Escudo
    If (UserList(UserIndex).Invent.EscudoEqpObjIndex > 0) Then
        daño = daño - ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).DañoMagico
    End If
    
    'Arma
    If (UserList(UserIndex).Invent.WeaponEqpObjIndex > 0) Then
        daño = daño - ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).DañoMagico
    End If
    
    'Armadura
    If (UserList(UserIndex).Invent.ArmourEqpObjIndex > 0) Then
        daño = daño - ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DañoMagico
    End If
    
    '<Edurne>   En montura...
    If UserList(UserIndex).flags.QueMontura Then _
        daño = daño + UserList(UserIndex).flags.Montura(UserList(UserIndex).flags.QueMontura).AtMagia
    '</Edurne>

    Call InfoHechizo(UserIndex, xArea)
    b = True
    
    If NPCList(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(NPCList(NpcIndex).flags.Snd2, NPCList(NpcIndex).Pos.X, NPCList(NpcIndex).Pos.Y))
    End If
    
    'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
    daño = daño - NPCList(NpcIndex).Stats.defM
    If daño < 0 Then daño = 0
    
    'Irongete: Si el jugador está invisible hace la mitad de daño
    If UserList(UserIndex).flags.invisible = 1 Then
        daño = daño * 0.5
    End If
    
    '<Edurne>
    If NPCList(NpcIndex).Numero = NPCReyCastle Then
        If UserList(UserIndex).GuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, MSG_ATK_CASTILLO_NOCLAN, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            Call ReyEsAtacado(UserIndex, NpcIndex, daño)
        End If
    End If
    '/<Edurne>
    
    '16/11/2015 Irongete: Está atacando la puerta del castillo
    If NPCList(NpcIndex).NPCType = 10 Then
        If UserList(UserIndex).GuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, MSG_ATK_CASTILLO_NOCLAN, FontTypeNames.FONTTYPE_INFO)
        Else
            Call PuertaEsAtacada(UserIndex, NpcIndex, daño)
        End If
    End If
    
    '26/02/2016 Irongete: Está atacando el defensor de la fortaleza
    If NPCList(NpcIndex).Numero = NPCDefensorFortaleza Then
        If UserList(UserIndex).GuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, MSG_ATK_CASTILLO_NOCLAN, FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            Call DefensorEsAtacado(UserIndex, NpcIndex, daño)
        End If
    End If
    
    
    Call EventosDaño(UserIndex, NpcIndex, daño)
    
    If GranPoder = UserIndex Then daño = daño * MultiplicadorGPN
    
    NPCList(NpcIndex).Stats.MinHP = NPCList(NpcIndex).Stats.MinHP - daño
    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateRenderValue(NPCList(NpcIndex).Pos.X, NPCList(NpcIndex).Pos.Y, daño, COLOR_DAÑO))
    
    ' No podras pasar de mapa por un rato
    Call IntervaloPermiteCambiardeMapa(UserIndex, True)
    '
    ' 27/10/2015 Irongete: Cambio el formato de los mensajes de los hechizos:
    '   Antes: ¡Le has causado 1234 puntos de daño a la criatura! -> Call WriteConsoleMsg(userIndex, "¡Le has causado " & Hechizos(hIndex).N & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
    '   Ahora: 'Tu Misil Mágico causa 1234 puntos de daño a Uruk (28766/30000)!
    Call WriteConsoleMsg(UserIndex, Hechizos(hIndex).nombre & " causa " & daño & " de daño " & NPCList(NpcIndex).Name & " (" & NPCList(NpcIndex).Stats.MinHP & "/" & NPCList(NpcIndex).Stats.MaxHP & ")", FontTypeNames.FONTTYPE_FIGHT)

    If NPCList(NpcIndex).Stats.MinHP < 1 Then
        NPCList(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, UserIndex)
    End If
End If

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer, Optional ByVal xArea As Boolean = False)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 25/07/2009
'25/07/2009: ZaMa - Code improvements.
'25/07/2009: ZaMa - Now invisible admins magic sounds are not sent to anyone but themselves
'***************************************************
    Dim SpellIndex As Integer
    Dim tUser As Integer
    Dim tNPC As Integer
    
    With UserList(UserIndex)
        SpellIndex = .Stats.UserHechizos(.flags.Hechizo)
        If xArea = 1 Then
            tUser = 0
            tNPC = 0
        Else
            tUser = .flags.targetUser
            tNPC = .flags.TargetNPC
        End If
        
        Call DecirPalabrasMagicas(Hechizos(SpellIndex).PalabrasMagicas, UserIndex)
        
        If tUser > 0 Then
            ' Los admins invisibles no producen sonidos ni fx's
            If .flags.AdminInvisible = 1 And UserIndex = tUser Then
                Call EnviarDatosASlot(UserIndex, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)
            End If
        ElseIf tNPC > 0 Then
            Call SendData(SendTarget.ToNPCArea, tNPC, PrepareMessageCreateFX(NPCList(tNPC).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
            Call SendData(SendTarget.ToNPCArea, tNPC, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, NPCList(tNPC).Pos.X, NPCList(tNPC).Pos.Y))
        End If
        
        If tUser > 0 Then
            If UserIndex <> tUser Then
                If .showName Then
                    'Call WriteConsoleMsg(userIndex, Hechizos(SpellIndex).HechizeroMsg & " " & UserList(tUser).name, FontTypeNames.FONTTYPE_FIGHT)
                Else
                    'Call WriteConsoleMsg(userIndex, Hechizos(SpellIndex).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT)
                End If
                'Call WriteConsoleMsg(tUser, .name & " " & Hechizos(SpellIndex).targetMSG, FontTypeNames.FONTTYPE_FIGHT)
            Else
                'Call WriteConsoleMsg(userIndex, Hechizos(SpellIndex).PropioMsg, FontTypeNames.FONTTYPE_FIGHT)
            End If
        ElseIf tNPC > 0 Then
            'Call WriteConsoleMsg(userIndex, Hechizos(SpellIndex).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With

End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean, Optional ByVal xArea As Boolean = False)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 02/01/2008
'02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
'***************************************************

Dim H As Integer
Dim daño As Long
Dim tempChr As Integer
Dim UserProtected As Boolean

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
tempChr = UserList(UserIndex).flags.targetUser
      
UserProtected = Not IntervaloPermiteSerAtacado(tempChr) And UserList(tempChr).flags.NoPuedeSerAtacado
    
If UserProtected Then b = False: Exit Sub
      
If UserList(tempChr).flags.Muerto Then
    Call WriteConsoleMsg(UserIndex, "No podés lanzar ese hechizo a un muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
      
'Hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(UserIndex, xArea)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + daño
    If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then _
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
        
    UserList(tempChr).flags.Hambre = 0
        
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call WriteUpdateHungerAndThirst(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    Else
        Exit Sub
    End If
    
    Call InfoHechizo(UserIndex, xArea)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
    Call WriteUpdateHungerAndThirst(tempChr)

    ' No podras pasar de mapa por un rato
    Call IntervaloPermiteCambiardeMapa(UserIndex, True)
    Call IntervaloPermiteCambiardeMapa(tempChr, True)
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(UserIndex, xArea)
    
    daño = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + daño
    If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then _
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
    
    UserList(tempChr).flags.Sed = 0
    
    Call WriteUpdateHungerAndThirst(tempChr)
         
    If UserIndex <> tempChr Then
      Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
      Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    Else
      Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex, xArea)
    
    daño = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
        UserList(tempChr).Stats.MinAGU = 0
        UserList(tempChr).flags.Sed = 1
    End If
    
    Call WriteUpdateHungerAndThirst(tempChr)
    
    ' No podras pasar de mapa por un rato
    Call IntervaloPermiteCambiardeMapa(UserIndex, True)
    Call IntervaloPermiteCambiardeMapa(tempChr, True)
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    
    'Para poder tirar cl a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
        If criminal(tempChr) And Not criminal(UserIndex) Then
            If esArmada(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            If UserList(UserIndex).flags.Seguro = 1 Then
                Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
            End If
        End If
    End If
    
    Call InfoHechizo(UserIndex, xArea)
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
    If UserList(UserIndex).clase = Bard Then
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) + 15) Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) + 15)
    Else
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) + 13) Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) + 13)
    End If
    
    UserList(tempChr).flags.TomoPocion = True
    Call WriteUpdateDexterity(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex, xArea)
    
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    Call WriteUpdateDexterity(tempChr)
    
    ' No podras pasar de mapa por un rato
    Call IntervaloPermiteCambiardeMapa(UserIndex, True)
    Call IntervaloPermiteCambiardeMapa(tempChr, True)
    b = True
    
End If

' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
    'Para poder tirar fuerza a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
        If criminal(tempChr) And Not criminal(UserIndex) Then
            If esArmada(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            If UserList(UserIndex).flags.Seguro = 1 Then
                Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
            End If
        End If
    End If
    
    Call InfoHechizo(UserIndex, xArea)
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño
    
    If UserList(UserIndex).clase = Bard Then
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) + 15) Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) + 15)
    Else
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) + 13) Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) + 13)
    End If
    
    UserList(tempChr).flags.TomoPocion = True
    Call WriteUpdateStrenght(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex, xArea)
    
    UserList(tempChr).flags.TomoPocion = True
    
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    Call WriteUpdateStrenght(tempChr)
    ' No podras pasar de mapa por un rato
    Call IntervaloPermiteCambiardeMapa(UserIndex, True)
    Call IntervaloPermiteCambiardeMapa(tempChr, True)
    b = True
    
End If

'Salud
If Hechizos(H).SubeHP = 1 Then
    
    'Verifica que el usuario no este muerto
    If UserList(tempChr).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'Para poder tirar curar a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
        If criminal(tempChr) And Not criminal(UserIndex) Then
            If esArmada(UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            If UserList(UserIndex).flags.Seguro = 1 Then
                Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
            End If
        End If
    End If
       
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex, xArea)

    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + daño
    If UserList(tempChr).Stats.MinHP > VidaMaxima(tempChr) Then _
        UserList(tempChr).Stats.MinHP = VidaMaxima(tempChr)
    
    Call WriteUpdateHP(tempChr)
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
    If UserIndex = tempChr Then
        Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    '<Edurne> 'En montura...
    If UserList(UserIndex).flags.QueMontura Then _
        daño = daño + UserList(UserIndex).flags.Montura(UserList(UserIndex).flags.QueMontura).AtMagia
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
    'anillos
    If (UserList(tempChr).Invent.AnilloEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
    End If
    
    'Escudo
    If (UserList(tempChr).Invent.EscudoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).DefensaMagicaMax)
    End If
    
    'Arma
    If (UserList(tempChr).Invent.WeaponEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.WeaponEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.WeaponEqpObjIndex).DefensaMagicaMax)
    End If
    
    'Armadura
    If (UserList(tempChr).Invent.ArmourEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
    End If
    
    '<Edurne>    'En montura...
    If UserList(tempChr).flags.QueMontura Then _
        daño = daño - UserList(tempChr).flags.Montura(UserList(tempChr).flags.QueMontura).DefMagia
    
    If daño < 0 Then daño = 0
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex, xArea)
    
    If GranPoder = UserIndex Then daño = daño * MultiplicadorGP
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    
    Call WriteUpdateHP(tempChr)
    Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageCreateRenderValue(UserList(tempChr).Pos.X, UserList(tempChr).Pos.Y, daño, COLOR_DAÑO))
    Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
    Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    
    ' No podras pasar de mapa por un rato
    Call IntervaloPermiteCambiardeMapa(UserIndex, True)
    Call IntervaloPermiteCambiardeMapa(tempChr, True)
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        'Store it!
        Call Statistics.StoreFrag(UserIndex, tempChr)
        'Lorwik> Comprobamos si esta en torneo
        Call MuerteEnTorneo(UserIndex, tempChr)
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        Call UserDie(tempChr, UserIndex)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(UserIndex, xArea)
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + daño
    If UserList(tempChr).Stats.MinMAN > ManaMaxima(tempChr) Then _
        UserList(tempChr).Stats.MinMAN = ManaMaxima(tempChr)
    
    Call WriteUpdateMana(tempChr)
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex, xArea)
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    
    Call WriteUpdateMana(tempChr)
    
    ' No podras pasar de mapa por un rato
    Call IntervaloPermiteCambiardeMapa(UserIndex, True)
    Call IntervaloPermiteCambiardeMapa(tempChr, True)
    
    b = True
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(UserIndex, xArea)
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + daño
    If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then _
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
    
    Call WriteUpdateSta(tempChr)
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de energia a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeSta = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex, xArea)
    
    If UserIndex <> tempChr Then
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de energia a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    
    Call WriteUpdateSta(tempChr)
    
    ' No podras pasar de mapa por un rato
    Call IntervaloPermiteCambiardeMapa(UserIndex, True)
    Call IntervaloPermiteCambiardeMapa(tempChr, True)
    
    b = True
End If

Call FlushBuffer(tempChr)

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(UserIndex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo


  If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
      
      Call WriteChangeSpellSlot(UserIndex, Slot)
  
  Else
  
      Call WriteChangeSpellSlot(UserIndex, Slot)
  
  End If


End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : MoveSpell
' Autor         : Facundo Ortega (GoDKeR)
' Fecha         : 27/12/2013
' Propósito     : Movemos el slot del Spell
'---------------------------------------------------------------------------------------
'
Sub MoveSpell(ByVal UserIndex As Integer, ByVal originalSlot As Byte, ByVal newSlot As Byte)
    
'#FABULOUS

    If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub
    
    Dim tmpSpell As Integer
    
    With UserList(UserIndex)
        
        If (originalSlot > 30) Or (newSlot > 30) Then Exit Sub
        
        tmpSpell = .Stats.UserHechizos(originalSlot)
        
        .Stats.UserHechizos(originalSlot) = .Stats.UserHechizos(newSlot)
        .Stats.UserHechizos(newSlot) = tmpSpell
    End With
    
WriteChangeSpellSlot UserIndex, originalSlot
WriteChangeSpellSlot UserIndex, newSlot

End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If (Dire <> 1 And Dire <> -1) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(UserIndex).flags.Hechizo > 0 Then
            UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo - 1
        End If
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(UserIndex).flags.Hechizo > 0 Then
            UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo + 1
        End If
    End If
End If
End Sub


Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)
'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos
    Dim EraCriminal As Boolean
    EraCriminal = criminal(UserIndex)
    
    'Si estamos en la arena no hacemos nada
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
    
If UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
    'pierdo nobleza...
    UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep - NoblePts
    If UserList(UserIndex).Reputacion.NobleRep < 0 Then
        UserList(UserIndex).Reputacion.NobleRep = 0
    End If
    
    'gano bandido...
    UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + BandidoPts
    If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then _
        UserList(UserIndex).Reputacion.BandidoRep = MAXREP
    Call WriteNobilityLost(UserIndex)
    If criminal(UserIndex) Then If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End If
    
    If Not EraCriminal And criminal(UserIndex) Then
        Call RefreshCharStatus(UserIndex)
    End If
End Sub
