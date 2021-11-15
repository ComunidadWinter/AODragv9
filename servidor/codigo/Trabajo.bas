Attribute VB_Name = "Trabajo"
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

Private Const ENERGIA_TRABAJO_HERRERO As Byte = 2


Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 28/01/2007
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'********************************************************

UserList(UserIndex).Counters.TiempoOculto = UserList(UserIndex).Counters.TiempoOculto - 1
If UserList(UserIndex).Counters.TiempoOculto <= 0 Then
    
    UserList(UserIndex).Counters.TiempoOculto = IntervaloOculto
    If UserList(UserIndex).clase = eClass.Hunter And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
        If UserList(UserIndex).Invent.ArmourEqpObjIndex = 648 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 360 Then
            Exit Sub
        End If
    End If
    UserList(UserIndex).Counters.TiempoOculto = 0
    UserList(UserIndex).flags.Oculto = 0
    If UserList(UserIndex).flags.invisible = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
End If

Exit Sub

errHandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
On Error GoTo errHandler

Dim Suerte As Double
Dim res As Integer
Dim Skill As Integer

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse)

Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100

res = RandomNumber(1, 100)

If res <= Suerte Then

    UserList(UserIndex).flags.Oculto = 1
    Suerte = (-0.000001 * (100 - Skill) ^ 3)
    Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
    Suerte = Suerte + (-0.0088 * (100 - Skill))
    Suerte = Suerte + (0.9571)
    Suerte = Suerte * IntervaloOculto
    UserList(UserIndex).Counters.TiempoOculto = Suerte
  
    Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, True)
    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))

    Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
    Call SubirSkill(UserIndex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 4 Then
        Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando + 1

Exit Sub

errHandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)
'*************************************************
'Author: Lorwik
'Ultima modificacion: 16/12/2018
'Descripción: Cambia el estado del usuario a navegando y le aplica las propiedades.
'Changelog:
'- Añadido modificador de velocidad
'*************************************************
    Dim ModNave As Long
    
    With UserList(UserIndex)
        ModNave = ModNavegacion(.clase)
        
        If .Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion. Consulta el manual del juego en http://www.aodrag.es/wiki/ para saber cómo conseguirlos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
        .Invent.BarcoSlot = Slot
        
        If .flags.Navegando = 0 Then
            
            .Char.Head = 0
            
            If .flags.Muerto = 0 Then
                '(Nacho)
                If .Faccion.ArmadaReal = 1 Or .Faccion.Legion = 1 Then
                    .Char.body = iFragataReal
                ElseIf .Faccion.FuerzasCaos = 1 Then
                    .Char.body = iFragataCaos
                Else
                    If criminal(UserIndex) Then
                        If Barco.Ropaje = iBarca Then .Char.body = iBarcaPk
                        If Barco.Ropaje = iGalera Then .Char.body = iGaleraPk
                        If Barco.Ropaje = iGaleon Then .Char.body = iGaleonPk
                    Else
                        If Barco.Ropaje = iBarca Then .Char.body = iBarcaCiuda
                        If Barco.Ropaje = iGalera Then .Char.body = iGaleraCiuda
                        If Barco.Ropaje = iGaleon Then .Char.body = iGaleonCiuda
                    End If
                End If
            Else
                .Char.body = iFragataFantasmal
            End If
            
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .flags.Navegando = 1
            .flags.Speed = ObjData(.Invent.Object(Slot).ObjIndex).Speed
            
        Else
            
            .flags.Navegando = 0
            
            If .flags.Muerto = 0 Then
                .Char.Head = .OrigChar.Head
                
                If .Invent.ArmourEqpObjIndex > 0 Then
                    .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                Else
                    Call DarCuerpoDesnudo(UserIndex)
                End If
                
                If .Invent.EscudoEqpObjIndex > 0 Then _
                    .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
                If .Invent.WeaponEqpObjIndex > 0 Then _
                    .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
                If .Invent.CascoEqpObjIndex > 0 Then _
                    .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
            Else
                .Char.body = iCuerpoMuerto
                .Char.Head = iCabezaMuerto
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            End If
            .flags.Speed = 0
        End If
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteNavigateToggle(UserIndex)
        Call WriteChangeSpeed(UserIndex, .Char.CharIndex, .flags.Speed)
    End With
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)

On Error GoTo errHandler

If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList(UserIndex).clase) Then
        Call DoLingotes(UserIndex)
   Else
        Call WriteConsoleMsg(UserIndex, "No tenes conocimientos de mineria suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
   End If

End If

Exit Sub

errHandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.Description)

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(UserIndex).Invent.Object(i).Amount
    End If
Next i

If cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/08/09
'05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
'***************************************************

'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    With UserList(UserIndex).Invent.Object(i)
        If .ObjIndex = ItemIndex Then
            If .Amount <= cant And .Equipped = 1 Then Call Desequipar(UserIndex, i)
            
            .Amount = .Amount - cant
            If .Amount <= 0 Then
                cant = Abs(.Amount)
                UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                .Amount = 0
                .ObjIndex = 0
            Else
                cant = 0
            End If
            
            Call UpdateUserInv(False, UserIndex, i)
            
            If cant = 0 Then Exit Sub
        End If
    End With
Next i

End Sub

Sub QuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Integer)
Dim i As Byte
    If ObjData(ItemIndex).CantMateriales > 0 Then
        For i = 1 To ObjData(ItemIndex).CantMateriales
        Debug.Print ObjData(ItemIndex).Material(i).Material
            If ObjData(ItemIndex).Material(i).CantMaterial > 0 Then Call QuitarObjetos(ObjData(ItemIndex).Material(i).Material, ObjData(ItemIndex).Material(i).CantMaterial * cantidad, UserIndex)
        Next i
    End If
End Sub
 
Function TieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Integer) As Boolean
    Dim i As Byte
    Debug.Print ObjData(ItemIndex).CantMateriales
    If ObjData(ItemIndex).CantMateriales > 0 Then
    Debug.Print "hola"
        For i = 1 To ObjData(ItemIndex).CantMateriales
            If Not TieneObjetos(ObjData(ItemIndex).Material(i).Material, ObjData(ItemIndex).Material(i).CantMaterial * cantidad, UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "No tenes suficientes materiales.", FontTypeNames.FONTTYPE_INFO)
                    TieneMateriales = False
                Exit Function
            End If
        Next i
    End If
    
    TieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 24/08/2009
'24/08/2008: ZaMa - Validates if the player has the required skill
'***************************************************
PuedeConstruir = TieneMateriales(UserIndex, ItemIndex, cantidad) And _
                    Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) >= ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function


Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Integer)

If PuedeConstruir(UserIndex, ItemIndex, cantidad) And PuedeConstruirHerreria(ItemIndex) Then
    
    Dim EnergiaFinal As Long
    
    'Chequeamos que tenga los puntos antes de sacarselos
    If UserList(UserIndex).Stats.MinSta >= ENERGIA_TRABAJO_HERRERO * cantidad Then
        EnergiaFinal = UserList(UserIndex).Stats.MinSta - (ENERGIA_TRABAJO_HERRERO * cantidad)
        If EnergiaFinal < 0 Then EnergiaFinal = 0
        UserList(UserIndex).Stats.MinSta = EnergiaFinal
        Call WriteUpdateSta(UserIndex)
    Else
        Call WriteMultiMessage(UserIndex, eMessages.NoEnergia)
        Exit Sub
    End If
    
    Call QuitarMateriales(UserIndex, ItemIndex, cantidad)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call WriteConsoleMsg(UserIndex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
        Call WriteConsoleMsg(UserIndex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
        Call WriteConsoleMsg(UserIndex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
        Call WriteConsoleMsg(UserIndex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.Amount = cantidad
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
    If ObjData(MiObj.ObjIndex).Log = 1 Then
        Call LogDesarrollo(UserList(UserIndex).Name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)
    End If

    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
End If
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 24/08/2009
'24/08/2008: ZaMa - Validates if the player has the required skill
'***************************************************
Dim EnergiaFinal As Long

If TieneMateriales(UserIndex, ItemIndex, cantidad) And _
   Round(UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase), 0) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(UserIndex).Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
   
    'Chequeamos que tenga los puntos antes de sacarselos
    If UserList(UserIndex).Stats.MinSta >= (ENERGIA_TRABAJO_HERRERO * cantidad) Then
        EnergiaFinal = UserList(UserIndex).Stats.MinSta - (ENERGIA_TRABAJO_HERRERO * cantidad)
        If EnergiaFinal < 0 Then EnergiaFinal = 0
        UserList(UserIndex).Stats.MinSta = EnergiaFinal
        Call WriteUpdateSta(UserIndex)
    Else
        Call WriteMultiMessage(UserIndex, eMessages.NoEnergia)
        Exit Sub
    End If
    
    Call QuitarMateriales(UserIndex, ItemIndex, cantidad)
    Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = cantidad
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
    If ObjData(MiObj.ObjIndex).Log = 1 Then
        Call LogDesarrollo(UserList(UserIndex).Name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)
    End If
    
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End If
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 10
        Case iMinerales.PlataCruda
            MineralesParaLingote = 20
        Case iMinerales.OroCrudo
            MineralesParaLingote = 30
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)
'    Call LogTarea("Sub DoLingotes")
    Dim Slot As Integer
    Dim obji As Integer

    Slot = UserList(UserIndex).flags.TargetObjInvSlot
    obji = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    
    If UserList(UserIndex).Invent.Object(Slot).Amount < MineralesParaLingote(obji) Or _
        ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(UserIndex, "No tienes mas minerales para fundir... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFOBOLD)
            UserList(UserIndex).flags.Makro = 0
            Call WriteMultiMessage(UserIndex, eMessages.NoTrabaja)
            Exit Sub
    End If
    
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - MineralesParaLingote(obji)
    If UserList(UserIndex).Invent.Object(Slot).Amount < 1 Then
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    End If
    
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call UpdateUserInv(False, UserIndex, Slot)
    Call WriteConsoleMsg(UserIndex, "¡Has obtenido un lingote!", FontTypeNames.FONTTYPE_INFOBOLD)

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
End Sub

Function ModNavegacion(ByVal clase As eClass) As Single

    ModNavegacion = 1

End Function


Function ModFundicion(ByVal clase As eClass) As Single

    ModFundicion = 1

End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer

    ModCarpinteria = 1

End Function

Function ModHerreriA(ByVal clase As eClass) As Single

    ModHerreriA = 1

End Function

Function ModDomar(ByVal clase As eClass) As Integer
    Select Case clase
        Case eClass.Druid
            ModDomar = 6
        Case eClass.Hunter
            ModDomar = 6
        Case eClass.Cleric
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 02/03/09
'02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
'***************************************************
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Nacho (Integer)
'Last Modification: 02/03/2009
'12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
'02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
'***************************************************

On Error GoTo errHandler

Dim puntosDomar As Integer
Dim puntosRequeridos As Integer
Dim CanStay As Boolean
Dim petType As Integer
Dim NroPets As Integer


If NPCList(NpcIndex).MaestroUser = UserIndex Then
    Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
    
    If NPCList(NpcIndex).MaestroNpc > 0 Or NPCList(NpcIndex).MaestroUser > 0 Then
        Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
        Call WriteConsoleMsg(UserIndex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    puntosDomar = CInt(UserList(UserIndex).Stats.UserSkills(eSkill.Domar))
    puntosRequeridos = NPCList(NpcIndex).flags.Domable
    
    If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
        Dim Index As Integer
        UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1
        Index = FreeMascotaIndex(UserIndex)
        UserList(UserIndex).MascotasIndex(Index) = NpcIndex
        UserList(UserIndex).MascotasType(Index) = NPCList(NpcIndex).Numero
        
        NPCList(NpcIndex).MaestroUser = UserIndex
        
        Call FollowAmo(NpcIndex)
        Call ReSpawnNpc(NPCList(NpcIndex))
        
        Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        
        ' Es zona segura?
        CanStay = (MapInfo(UserList(UserIndex).Pos.Map).Pk = True)
        
        If Not CanStay Then
            petType = NPCList(NpcIndex).Numero
            NroPets = UserList(UserIndex).NroMascotas
            
            Call QuitarNPC(NpcIndex)
            
            UserList(UserIndex).MascotasType(Index) = petType
            UserList(UserIndex).NroMascotas = NroPets
            
            Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        End If

    Else
        If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
            Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 5
        End If
    End If
    
Else
    Call WriteConsoleMsg(UserIndex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
End If

Exit Sub

errHandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.Description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'This function checks how many NPCs of the same type have
'been tamed by the user.
'Returns True if that amount is less than two.
'***************************************************
    Dim i As Long
    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) = NPCList(NpcIndex).Numero Then
            numMascotas = numMascotas + 1
        End If
    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/07/2009
'Makes an admin invisible o visible.
'13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
'***************************************************
    
    With UserList(UserIndex)
        If .flags.AdminInvisible = 0 Then
            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0
            End If
            
            .flags.AdminInvisible = 1
            .flags.invisible = 1
            .flags.Oculto = 1
            .flags.OldBody = .Char.body
            .flags.OldHead = .Char.Head
            .Char.body = 0
            .Char.Head = 0
            
            ' Solo el admin sabe que se hace invi
            Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            .Char.body = .flags.OldBody
            .Char.Head = .flags.OldHead
            
            'Borramos el personaje en del cliente del GM
            Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
            'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
            Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
        End If
    End With
    
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

Dim Obj As Obj
Dim posMadera As worldPos

If Not LegalPos(Map, X, Y) Then Exit Sub

With posMadera
    .Map = Map
    .X = X
    .Y = Y
End With

If MapData(Map, X, Y).ObjInfo.ObjIndex <> 58 Then
    Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
    Call WriteMultiMessage(UserIndex, eMessages.Lejos)
    Exit Sub
End If

If UserList(UserIndex).flags.Muerto = 1 Then
    Call WriteMultiMessage(UserIndex, eMessages.Muerto)
    Exit Sub
End If

If MapData(Map, X, Y).ObjInfo.Amount < 3 Then
    Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Obj.ObjIndex = FOGATA_APAG
Obj.Amount = MapData(Map, X, Y).ObjInfo.Amount \ 3
    
Call WriteConsoleMsg(UserIndex, "Has hecho " & Obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
Call MakeObj(Obj, Map, X, Y)
    
'Seteamos la fogata como el nuevo TargetObj del user
UserList(UserIndex).flags.targetObj = FOGATA_APAG

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
On Error GoTo errHandler

Dim Suerte As Integer
Dim res As Integer

Call QuitarSta(UserIndex, EsfuerzoPescarPescador)

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.pesca)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

'Si esta en zona segura le cuesta el doble
If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then Suerte = Suerte * 2

res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    
    MiObj.Amount = RandomNumber(1, 5)
    MiObj.ObjIndex = Pescado
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call WriteConsoleMsg(UserIndex, "¡Has pescado un lindo pez!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
      Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
      UserList(UserIndex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

errHandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " : " & Err.Description)
End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)
On Error GoTo errHandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer

Call QuitarSta(UserIndex, EsfuerzoPescarPescador)

iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

'Si esta en zona segura le cuesta el doble
If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then Suerte = Suerte * 2

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 4) As Integer
        
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO2
        PecesPosibles(3) = PESCADO3
        PecesPosibles(4) = PESCADO4
        
        MiObj.Amount = RandomNumber(5, 10)
        
        MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call WriteConsoleMsg(UserIndex, "¡Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
        
    Else
        Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
    End If

End If
        
Exit Sub

errHandler:
    Call LogError("Error en DoPescarRed")
End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 24/07/028
'Last Modification By: Marco Vanotti (MarKoxX)
' - 24/07/08 Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
'*************************************************

On Error GoTo errHandler

If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub

If UserList(LadrOnIndex).flags.Seguro = 1 Then
    Call WriteConsoleMsg(LadrOnIndex, "Debes quitar el seguro para robar", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If

If UserList(LadrOnIndex).PartyId <> UserList(VictimaIndex).PartyId Then
    Call WriteConsoleMsg(LadrOnIndex, "¡No puedes robar a miembros de tu party.", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If

If EsNewbie(LadrOnIndex) = True Then
    Call WriteConsoleMsg(LadrOnIndex, "Los Newbies no pueden robar.", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If

If EsNewbie(VictimaIndex) = True Then
    Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a los Newbies.", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If

If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de las fuerzas del caos", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If


Call QuitarSta(LadrOnIndex, 15)

Dim GuantesHurto As Boolean
'Tiene los Guantes de Hurto equipados?
GuantesHurto = True
If UserList(LadrOnIndex).Invent.AnilloEqpObjIndex = 0 Then
    GuantesHurto = False
Else
    If ObjData(UserList(LadrOnIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then GuantesHurto = False
    If ObjData(UserList(LadrOnIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then GuantesHurto = False
End If


If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) < 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
                        Suerte = 7
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) = 100 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
        
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim n As Integer
                
                'Si no tine puestos los guantes de hurto roba un 20% menos. Pablo (ToxicWaste)
                If GuantesHurto Then
                    n = RandomNumber(100, 1000)
                Else
                    n = RandomNumber(80, 800)
                End If
                If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                
                UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + n
                If UserList(LadrOnIndex).Stats.GLD > MAXORO Then _
                    UserList(LadrOnIndex).Stats.GLD = MAXORO
                
                Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_oro)
                Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                
                Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                Call FlushBuffer(VictimaIndex)
            Else
                Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    Else
        Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(LadrOnIndex).Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(LadrOnIndex).Name & " es un criminal!", FontTypeNames.FONTTYPE_INFO)
        Call FlushBuffer(VictimaIndex)
    End If

    If Not criminal(LadrOnIndex) Then
        Call VolverCriminal(LadrOnIndex)
    End If
    
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)
    'If UserList(LadrOnIndex).Faccion.Legion = 1 Then Call expulsarfaccionrelegion(LadrOnIndex)
    
    UserList(LadrOnIndex).Reputacion.LadronesRep = UserList(LadrOnIndex).Reputacion.LadronesRep + vlLadron
    If UserList(LadrOnIndex).Reputacion.LadronesRep > MAXREP Then _
        UserList(LadrOnIndex).Reputacion.LadronesRep = MAXREP
End If

Exit Sub

errHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.Description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

    Dim OI As Integer

    OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex
    
    ObjEsRobable = _
    ObjData(OI).OBJType <> eOBJType.otLlaves And _
    UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
    ObjData(OI).Real = 0 And _
    ObjData(OI).Caos = 0 And _
    ObjData(OI).OBJType <> eOBJType.otMonturas And _
    ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
Else
    Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
End If

'If exiting, cancel de quien es robado
Call CancelExit(VictimaIndex)

End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

If Not UserList(UserIndex).clase = eClass.Assasin Then Exit Sub

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)

If RandomNumber(0, 100) < Suerte Then
    If VictimUserIndex <> 0 Then
        daño = Round(daño * 1.4, 0)

        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateRenderValue(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, daño, COLOR_DAÑO))
        Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_APUÑALA)
        
        Call FlushBuffer(VictimUserIndex)
    Else
        NPCList(VictimNpcIndex).Stats.MinHP = NPCList(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageCreateRenderValue(NPCList(VictimNpcIndex).Pos.X, NPCList(VictimNpcIndex).Pos.Y, Int(daño * 2), COLOR_DAÑO))
        Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño * 2), FontTypeNames.FONTTYPE_FIGHT)
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_APUÑALA)
        Call SubirSkill(UserIndex, Apuñalar)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
End If

End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name <> "Espada Vikinga" Then Exit Sub


Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling)

Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0493) * 100)

If RandomNumber(0, 100) < Suerte Then
    daño = Int(daño * 0.5)
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateRenderValue(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, daño, COLOR_DAÑO))
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado críticamente por " & daño, FontTypeNames.FONTTYPE_FIGHT)
    Else
        NPCList(VictimNpcIndex).Stats.MinHP = NPCList(VictimNpcIndex).Stats.MinHP - daño
        Call SendData(SendTarget.ToPCArea, VictimNpcIndex, PrepareMessageCreateRenderValue(UserList(VictimNpcIndex).Pos.X, UserList(VictimNpcIndex).Pos.Y, daño, COLOR_DAÑO))
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
    End If
End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal cantidad As Integer)

On Error GoTo errHandler

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateSta(UserIndex)
    
Exit Sub

errHandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.Description)
    
End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim cant As Integer

If UserList(UserIndex).Counters.bPuedeMeditar = False Then
    UserList(UserIndex).Counters.bPuedeMeditar = True
End If
    
If UserList(UserIndex).Stats.MinMAN >= ManaMaxima(UserIndex) Then
    Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFOBOLD)
    Call WriteMeditateToggle(UserIndex)
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) < 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 7
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    
    cant = Porcentaje(ManaMaxima(UserIndex), PorcentajeRecuperoMana)
    If cant <= 0 Then cant = 1
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + cant
    If UserList(UserIndex).Stats.MinMAN > ManaMaxima(UserIndex) Then _
        UserList(UserIndex).Stats.MinMAN = ManaMaxima(UserIndex)

    Call WriteUpdateMana(UserIndex)
    Call SubirSkill(UserIndex, Meditar)
End If

End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 28/01/2007
'Implements the pick pocket skill of the Bandit :)
'***************************************************
'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
'Uso el slot de los anillos para "equipar" los guantes.
'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
If UserList(UserIndex).Invent.AnilloEqpObjIndex = 0 Then
    Exit Sub
Else
    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then Exit Sub
    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then Exit Sub
End If

Dim res As Integer
res = RandomNumber(1, 100)
If (res < 20) Then
    If TieneObjetosRobables(VictimaIndex) Then
        Call RobarObjeto(UserIndex, VictimaIndex)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 17/02/2007
'Implements the special Skill of the Thief
'***************************************************
If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    
'una vez más, la forma de reconocer los guantes es medio patética.
If UserList(UserIndex).Invent.AnilloEqpObjIndex = 0 Then
    Exit Sub
Else
    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then Exit Sub
    If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then Exit Sub
End If

    
Dim res As Integer
res = RandomNumber(0, 100)
If res < (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
    UserList(VictimaIndex).flags.Paralizado = 1
    UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
    Call WriteParalizeOK(VictimaIndex)
    Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 91 Then
                    Suerte = 7
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        If UserList(VictimIndex).Stats.ELV < 20 Then
            Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
        End If
        Call FlushBuffer(VictimIndex)
    End If
End Sub

Public Sub EsTrabajo(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    'NUEVO SISTEMA DE EXTRACCION
    If NPCList(NpcIndex).NPCType = Arbol Then
        If Not UserList(UserIndex).Invent.WeaponEqpObjIndex = HACHA_LEÑADOR Then
            Call WriteConsoleMsg(UserIndex, "Necesitas un hacha de leñador para talar.", FontTypeNames.FONTTYPE_INFO)
        Else
            SkillTrabajando = eSkill.talar
            UserTrabajando = True
            Exit Sub
        End If
    End If
    
    If NPCList(NpcIndex).NPCType = Yacimiento Then
        If Not UserList(UserIndex).Invent.WeaponEqpObjIndex = PIQUETE_MINERO Then
            Call WriteConsoleMsg(UserIndex, "Necesitas un pico de minero para minar.", FontTypeNames.FONTTYPE_INFO)
        Else
            SkillTrabajando = eSkill.Mineria
            UserTrabajando = True
            Exit Sub
        End If
    End If
    
    UserTrabajando = False
End Sub
