Attribute VB_Name = "SistemaCombate"
'Argentum Online 0.9.83
'Copyright (C) 2001 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based in Baronsoft's VB6 Online RPG
'Engine 9/08/2000 http://www.baronsoft.com/
'aaron@baronsoft.com
'
'Contact info:
'Pablo Ignacio Márquez
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900


Option Explicit

Function ModificadorEvasion(ByVal Clase As String) As Single

Select Case UCase(Clase)
    Case "GUERRERO"
        ModificadorEvasion = 1
    Case "CAZADOR"
        ModificadorEvasion = 0.9
    Case "PALADIN"
        ModificadorEvasion = 0.9
    Case "BANDIDO"
        ModificadorEvasion = 0.9
    Case "ASESINO"
        ModificadorEvasion = 1.1
    Case "PIRATA"
        ModificadorEvasion = 0.9
    Case "LADRON"
        ModificadorEvasion = 1.1
    Case "BARDO"
        ModificadorEvasion = 1.1
    Case Else
        ModificadorEvasion = 0.8
End Select
End Function

Function ModificadorPoderAtaqueArmas(ByVal Clase As String) As Single
Select Case UCase(Clase)
    Case "GUERRERO"
        ModificadorPoderAtaqueArmas = 1
    Case "CAZADOR"
        ModificadorPoderAtaqueArmas = 0.8
    Case "PALADIN"
        ModificadorPoderAtaqueArmas = 0.85
    Case "ASESINO"
        ModificadorPoderAtaqueArmas = 0.85
    Case "PIRATA"
        ModificadorPoderAtaqueArmas = 0.8
    Case "LADRON"
        ModificadorPoderAtaqueArmas = 0.75
    Case "BANDIDO"
        ModificadorPoderAtaqueArmas = 0.75
    Case "CLERIGO"
        ModificadorPoderAtaqueArmas = 0.7
    Case "BARDO"
        ModificadorPoderAtaqueArmas = 0.7
    Case "DRUIDA"
        ModificadorPoderAtaqueArmas = 0.7
    Case "PESCADOR"
        ModificadorPoderAtaqueArmas = 0.6
    Case "LEÑADOR"
        ModificadorPoderAtaqueArmas = 0.6
    Case "MINERO"
        ModificadorPoderAtaqueArmas = 0.6
    Case "HERRERO"
        ModificadorPoderAtaqueArmas = 0.6
    Case "CARPINTERO"
        ModificadorPoderAtaqueArmas = 0.6
    Case Else
        ModificadorPoderAtaqueArmas = 0.5
End Select
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal Clase As String) As Single
Select Case UCase(Clase)
    Case "GUERRERO"
        ModificadorPoderAtaqueProyectiles = 0.8
    Case "CAZADOR"
        ModificadorPoderAtaqueProyectiles = 1
    Case "PALADIN"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "ASESINO"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "PIRATA"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "LADRON"
        ModificadorPoderAtaqueProyectiles = 0.8
    Case "BANDIDO"
        ModificadorPoderAtaqueProyectiles = 0.8
    Case "CLERIGO"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case "BARDO"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case "DRUIDA"
        ModificadorPoderAtaqueProyectiles = 0.75
    Case "PESCADOR"
        ModificadorPoderAtaqueProyectiles = 0.65
    Case "LEÑADOR"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case "MINERO"
        ModificadorPoderAtaqueProyectiles = 0.65
    Case "HERRERO"
        ModificadorPoderAtaqueProyectiles = 0.65
    Case "CARPINTERO"
        ModificadorPoderAtaqueProyectiles = 0.7
    Case Else
        ModificadorPoderAtaqueProyectiles = 0.5
End Select
End Function

Function ModicadorDañoClaseArmas(ByVal Clase As String) As Single
Select Case UCase(Clase)
    Case "GUERRERO"
        ModicadorDañoClaseArmas = 1.1
    Case "CAZADOR"
        ModicadorDañoClaseArmas = 0.8
    Case "PALADIN"
        ModicadorDañoClaseArmas = 0.8
    Case "ASESINO"
        ModicadorDañoClaseArmas = 0.8
    Case "LADRON"
        ModicadorDañoClaseArmas = 0.7
    Case "PIRATA"
        ModicadorDañoClaseArmas = 0.7
    Case "BANDIDO"
        ModicadorDañoClaseArmas = 0.7
    Case "CLERIGO"
        ModicadorDañoClaseArmas = 0.65
    Case "BARDO"
        ModicadorDañoClaseArmas = 0.6
    Case "DRUIDA"
        ModicadorDañoClaseArmas = 0.6
    Case "PESCADOR"
        ModicadorDañoClaseArmas = 0.5
    Case "LEÑADOR"
        ModicadorDañoClaseArmas = 0.6
    Case "MINERO"
        ModicadorDañoClaseArmas = 0.7
    Case "HERRERO"
        ModicadorDañoClaseArmas = 0.7
    Case "CARPINTERO"
        ModicadorDañoClaseArmas = 0.6
    Case Else
        ModicadorDañoClaseArmas = 0.4
End Select
End Function

Function ModicadorDañoClaseProyectiles(ByVal Clase As String) As Single
Select Case UCase(Clase)
    Case "GUERRERO"
        ModicadorDañoClaseProyectiles = 1
    Case "CAZADOR"
        ModicadorDañoClaseProyectiles = 1.1
    Case "PALADIN"
        ModicadorDañoClaseProyectiles = 0.75
    Case "ASESINO"
        ModicadorDañoClaseProyectiles = 0.75
    Case "LADRON"
        ModicadorDañoClaseProyectiles = 0.7
    Case "PIRATA"
        ModicadorDañoClaseProyectiles = 0.7
    Case "BANDIDO"
        ModicadorDañoClaseProyectiles = 0.7
    Case "CLERIGO"
        ModicadorDañoClaseProyectiles = 0.6
    Case "BARDO"
        ModicadorDañoClaseProyectiles = 0.6
    Case "DRUIDA"
        ModicadorDañoClaseProyectiles = 0.65
    Case "PESCADOR"
        ModicadorDañoClaseProyectiles = 0.5
    Case "LEÑADOR"
        ModicadorDañoClaseProyectiles = 0.6
    Case "MINERO"
        ModicadorDañoClaseProyectiles = 0.5
    Case "HERRERO"
        ModicadorDañoClaseProyectiles = 0.5
    Case "CARPINTERO"
        ModicadorDañoClaseProyectiles = 0.6
    Case Else
        ModicadorDañoClaseProyectiles = 0.4
End Select
End Function

Function ModEvasionDeEscudoClase(ByVal Clase As String) As Integer

Select Case UCase(Clase)
Case "GUERRERO"
        ModEvasionDeEscudoClase = 1
    Case "CAZADOR"
        ModEvasionDeEscudoClase = 0.8
    Case "PALADIN"
        ModEvasionDeEscudoClase = 1
    Case "ASESINO"
        ModEvasionDeEscudoClase = 0.8
    Case "LADRON"
        ModEvasionDeEscudoClase = 0.7
    Case "BANDIDO"
        ModEvasionDeEscudoClase = 0.8
    Case "PIRATA"
        ModEvasionDeEscudoClase = 0.75
    Case "CLERIGO"
        ModEvasionDeEscudoClase = 0.9
    Case "BARDO"
        ModEvasionDeEscudoClase = 0.75
    Case "DRUIDA"
        ModEvasionDeEscudoClase = 0.75
    Case "PESCADOR"
        ModEvasionDeEscudoClase = 0.7
    Case "LEÑADOR"
        ModEvasionDeEscudoClase = 0.7
    Case "MINERO"
        ModEvasionDeEscudoClase = 0.7
    Case "HERRERO"
        ModEvasionDeEscudoClase = 0.7
    Case "CARPINTERO"
        ModEvasionDeEscudoClase = 0.7
    Case Else
        ModEvasionDeEscudoClase = 0.6
End Select

End Function
Function Minimo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Minimo = b
    Else: Minimo = a
End If
End Function

Function Maximo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Maximo = a
    Else: Maximo = b
End If
End Function

Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long

PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(Defensa) * _
ModEvasionDeEscudoClase(UserList(UserIndex).Clase)) / 2

End Function

Function PoderEvasion(ByVal UserIndex As Integer) As Long
Dim PoderEvasionTemp As Long

If UserList(UserIndex).Stats.UserSkills(Tacticas) < 31 Then
    PoderEvasionTemp = (UserList(UserIndex).Stats.UserSkills(Tacticas) * _
    ModificadorEvasion(UserList(UserIndex).Clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Tacticas) < 61 Then
        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(Tacticas) + _
        UserList(UserIndex).Stats.UserAtributos(Agilidad)) * _
        ModificadorEvasion(UserList(UserIndex).Clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Tacticas) < 91 Then
        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(Tacticas) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
        ModificadorEvasion(UserList(UserIndex).Clase))
Else
        PoderEvasionTemp = ((UserList(UserIndex).Stats.UserSkills(Tacticas) + _
        (3 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
        ModificadorEvasion(UserList(UserIndex).Clase))
End If

PoderEvasion = (PoderEvasionTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))

End Function

Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(Armas) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(Armas) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Armas) < 61 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Armas) + _
    UserList(UserIndex).Stats.UserAtributos(Agilidad)) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Armas) < 91 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Armas) + _
    (2 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase))
Else
   PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Armas) + _
   (3 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
   ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase))
End If

PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))
End Function

Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(Proyectiles) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(Proyectiles) * _
    ModificadorPoderAtaqueProyectiles(UserList(UserIndex).Clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Proyectiles) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Proyectiles) + _
        UserList(UserIndex).Stats.UserAtributos(Agilidad)) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).Clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Proyectiles) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Proyectiles) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).Clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Proyectiles) + _
      (3 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
      ModificadorPoderAtaqueProyectiles(UserList(UserIndex).Clase))
End If

PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))

End Function

Function PoderAtaqueWresterling(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(Wresterling) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(Wresterling) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Wresterling) + _
        UserList(UserIndex).Stats.UserAtributos(Agilidad)) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase))
ElseIf UserList(UserIndex).Stats.UserSkills(Wresterling) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Wresterling) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(Wresterling) + _
       (3 * UserList(UserIndex).Stats.UserAtributos(Agilidad))) * _
       ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase))
End If

PoderAtaqueWresterling = (PoderAtaqueTemp + (2.5 * Maximo(UserList(UserIndex).Stats.ELV - 12, 0)))

End Function


Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma > 0 Then 'Usando un arma
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(UserIndex)
    Else
        PoderAtaque = PoderAtaqueArma(UserIndex)
    End If
Else 'Peleando con puños
    PoderAtaque = PoderAtaqueWresterling(UserIndex)
End If


ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

If UserImpactoNpc Then
    If Arma <> 0 Then
       If proyectil Then
            Call SubirSkill(UserIndex, Proyectiles)
       Else
            Call SubirSkill(UserIndex, Armas)
       End If
    Else
        Call SubirSkill(UserIndex, Wresterling)
    End If
End If


End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long
Dim NpcPoderAtaque As Long
Dim PoderEvasioEscudo As Long
Dim SkillTacticas As Long
Dim SkillDefensa As Long

UserEvasion = PoderEvasion(UserIndex)
NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)

SkillTacticas = UserList(UserIndex).Stats.UserSkills(Tacticas)
SkillDefensa = UserList(UserIndex).Stats.UserSkills(Defensa)

'Esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
   If NpcImpacto = False Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
         Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_ESCUDO)
         Call SendData(ToIndex, UserIndex, 0, "7")
         Call SubirSkill(UserIndex, Defensa)
      End If
   End If
End If

End Function


Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Long
Dim proyectil As ObjData
Dim DañoMaxArma As Long

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    
    
    ' Ataca a un npc?
    If NpcIndex > 0 Then
        
        'Usa la mata dragones?
        If Arma.SubTipo = MATADRAGONES Then ' Usa la matadragones?
            If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca dragon?
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
            Else ' Sino es dragon daño es 1
                DañoArma = 1
                DañoMaxArma = 1
            End If
        Else ' daño comun
           If Arma.proyectil = 1 Then
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    
    Else ' Ataca usuario
        If Arma.SubTipo = MATADRAGONES Then
            DañoArma = 1 ' Si usa la espada matadragones daño es 1
            DañoMaxArma = 1
        Else
           If Arma.proyectil = 1 Then
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    End If
End If

DañoUsuario = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHIT)
ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).Clase)
CalcularDaño = (((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(Fuerza) - 15))) + DañoUsuario) * ModifClase)

End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
Dim daño As Long



daño = CalcularDaño(UserIndex, NpcIndex)

'esta navegando? si es asi le sumamos el daño del barco
If UserList(UserIndex).Flags.Navegando = 1 Then _
        daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)

daño = daño - Npclist(NpcIndex).Stats.Def

If daño < 0 Then daño = 0

Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño

Call SendData(ToIndex, UserIndex, 0, "U2" & daño)

If Npclist(NpcIndex).Stats.MinHP > 0 Then
    'Trata de apuñalar por la espalda al enemigo
    If PuedeApuñalar(UserIndex) Then
       Call DoApuñalar(UserIndex, NpcIndex, 0, daño)
       Call SubirSkill(UserIndex, Apuñalar)
    End If
End If

 
If Npclist(NpcIndex).Stats.MinHP <= 0 Then
          
          ' Si era un Dragon perdemos la espada matadragones
          If Npclist(NpcIndex).NPCtype = DRAGON Then Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
          
          ' Para que las mascotas no sigan intentando luchar y
          ' comiencen a seguir al amo
         
          Dim j As Integer
          For j = 1 To MAXMASCOTAS
                If UserList(UserIndex).MascotasIndex(j) > 0 Then
                    If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = NpcIndex Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = 0
                    Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = SIGUE_AMO
                End If
          Next j
  
          Call MuereNpc(NpcIndex, UserIndex)
End If

End Sub


Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

Dim daño As Integer, Lugar As Integer, absorbido As Integer, npcfile As String
Dim antdaño As Integer, defbarco As Integer
Dim Obj As ObjData



daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
antdaño = daño

If UserList(UserIndex).Flags.Navegando = 1 Then
    Obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If


Lugar = RandomNumber(1, 6)

Select Case Lugar
  Case bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
End Select

Call SendData(ToIndex, UserIndex, 0, "N2" & Lugar & "," & daño)

If UserList(UserIndex).Flags.Privilegios <> 3 Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño

'Muere el usuario
If UserList(UserIndex).Stats.MinHP <= 0 Then

    Call SendData(ToIndex, UserIndex, 0, "6") ' Le informamos que ha muerto ;)
    
    'Si lo mato un guardia
    If Criminal(UserIndex) And Npclist(NpcIndex).NPCtype = 2 Then
        If UserList(UserIndex).Reputacion.AsesinoRep > 0 Then
             UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep - vlASESINO / 4
             If UserList(UserIndex).Reputacion.AsesinoRep < 0 Then UserList(UserIndex).Reputacion.AsesinoRep = 0
        ElseIf UserList(UserIndex).Reputacion.BandidoRep > 0 Then
             UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep - vlASALTO / 4
             If UserList(UserIndex).Reputacion.BandidoRep < 0 Then UserList(UserIndex).Reputacion.BandidoRep = 0
        ElseIf UserList(UserIndex).Reputacion.LadronesRep > 0 Then
             UserList(UserIndex).Reputacion.LadronesRep = UserList(UserIndex).Reputacion.LadronesRep - vlCAZADOR / 3
             If UserList(UserIndex).Reputacion.LadronesRep < 0 Then UserList(UserIndex).Reputacion.LadronesRep = 0
        End If
    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        'Al matarlo no lo sigue mas
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).Flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).Flags.OldHostil
                    Npclist(NpcIndex).Flags.AttackedBy = ""
        End If
    End If
    
    Call UserDie(UserIndex)

End If

End Sub
Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
       If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
        If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = 0 Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNpc = NpcIndex
        'Npclist(UserList(UserIndex).MascotasIndex(j)).Flags.OldMovement = Npclist(UserList(UserIndex).MascotasIndex(j)).Movement
        Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = NPC_ATACA_NPC
       End If
    End If
Next j

End Sub
Public Sub AllFollowAmo(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
        Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
    End If
Next j
End Sub

Public Sub NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
' El npc puede atacar ???
If Npclist(NpcIndex).CanAttack = 1 Then
    Call CheckPets(NpcIndex, UserIndex)
    
    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex
    
    If UserList(UserIndex).Flags.AtacadoPorNpc = 0 And _
       UserList(UserIndex).Flags.AtacadoPorUser = 0 Then UserList(UserIndex).Flags.AtacadoPorNpc = NpcIndex
Else
    Exit Sub
End If

Npclist(NpcIndex).CanAttack = 0
    
    
   
If Npclist(NpcIndex).Flags.Snd1 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).Flags.Snd1)
        
     
    
If NpcImpacto(NpcIndex, UserIndex) Then
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_IMPACTO)
    
    If UserList(UserIndex).Flags.Navegando = 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)

    Call NpcDaño(NpcIndex, UserIndex)
    '¿Puede envenenar?
    If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
Else
    Call SendData(ToIndex, UserIndex, 0, "N1")
End If

'-----Tal vez suba los skills------
Call SubirSkill(UserIndex, Tacticas)

Call SendUserStatsBox(val(UserIndex))
'Controla el nivel del usuario
Call CheckUserLevel(UserIndex)

End Sub

Function NpcImpactoNpc(ByVal atacante As Integer, ByVal Victima As Integer) As Boolean
Dim PoderAtt As Long, PoderEva As Long, dif As Long
Dim ProbExito As Long

PoderAtt = Npclist(atacante).PoderAtaque
PoderEva = Npclist(Victima).PoderEvasion
ProbExito = Maximo(10, Minimo(90, 50 + _
            ((PoderAtt - PoderEva) * 0.4)))
NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

End Function

Public Sub NpcDañoNpc(ByVal atacante As Integer, ByVal Victima As Integer)
Dim daño As Integer
Dim ANpc As Npc, DNpc As Npc
ANpc = Npclist(atacante)

daño = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)
Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - daño

If Npclist(Victima).Stats.MinHP < 1 Then
        
        If Npclist(atacante).Flags.AttackedBy <> "" Then
            Npclist(atacante).Movement = Npclist(atacante).Flags.OldMovement
            Npclist(atacante).Hostile = Npclist(atacante).Flags.OldHostil
        Else
            Npclist(atacante).Movement = Npclist(atacante).Flags.OldMovement
        End If
        
        Call FollowAmo(atacante)
        Call MuereNpc(Victima, Npclist(atacante).MaestroUser)
End If

End Sub

Public Sub NpcAtacaNpc(ByVal atacante As Integer, ByVal Victima As Integer)

' El npc puede atacar ???
If Npclist(atacante).CanAttack = 1 Then
       Npclist(atacante).CanAttack = 0
       Npclist(Victima).TargetNpc = atacante
       Npclist(Victima).Movement = NPC_ATACA_NPC
Else
    Exit Sub
End If

If Npclist(atacante).Flags.Snd1 > 0 Then Call SendData(ToNPCArea, atacante, Npclist(atacante).Pos.Map, "TW" & Npclist(atacante).Flags.Snd1)


If NpcImpactoNpc(atacante, Victima) Then
    
    If Npclist(Victima).Flags.Snd2 > 0 Then
        Call SendData(ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & Npclist(Victima).Flags.Snd2)
    Else
        Call SendData(ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & SND_IMPACTO2)
    End If

    If Npclist(atacante).MaestroUser > 0 Then
        Call SendData(ToNPCArea, atacante, Npclist(atacante).Pos.Map, "TW" & SND_IMPACTO)
    Else
        Call SendData(ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & SND_IMPACTO)
    End If
    Call NpcDañoNpc(atacante, Victima)
    
Else
    If Npclist(atacante).MaestroUser > 0 Then
        Call SendData(ToNPCArea, atacante, Npclist(atacante).Pos.Map, "TW" & SOUND_SWING)
    Else
        Call SendData(ToNPCArea, Victima, Npclist(Victima).Pos.Map, "TW" & SOUND_SWING)
    End If
End If

End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 7 Then
   Call SendData(ToIndex, UserIndex, 0, "||Estás muy lejos para disparar." & FONTTYPE_FIGHT)
   Exit Sub
End If

Call NpcAtacado(NpcIndex, UserIndex)

If UserImpactoNpc(UserIndex, NpcIndex) Then
    
    If Npclist(NpcIndex).Flags.Snd2 > 0 Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).Flags.Snd2)
    Else
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_IMPACTO2)
    End If
    
    
    
    
    Call UserDañoNpc(UserIndex, NpcIndex)
   
Else
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SWING)
    Call SendData(ToIndex, UserIndex, 0, "U1")
End If

End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)

If UserList(UserIndex).Flags.PuedeAtacar = 1 Then
    
    'Quitamos stamina
    If UserList(UserIndex).Stats.MinSta >= 10 Then
        Call QuitarSta(UserIndex, RandomNumber(1, 10))
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(UserIndex).Flags.PuedeAtacar = 0
    
    Dim AttackPos As WorldPos
    AttackPos = UserList(UserIndex).Pos
    Call HeadtoPos(UserList(UserIndex).Char.Heading, AttackPos)
    
    'Exit if not legal
    If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SWING)
        Exit Sub
    End If
    
    Dim Index As Integer
    Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
    
    'Look for user
    If Index > 0 Then
        Call UsuarioAtacaUsuario(UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex)
        Call SendUserStatsBox(UserIndex)
        Call SendUserStatsBox(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex)
        Exit Sub
    End If
    
    'Look for NPC
    If MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then
    
        If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Attackable Then
            
            If Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).MaestroUser > 0 And _
               MapInfo(Npclist(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex).Pos.Map).Pk = False Then
                    Call SendData(ToIndex, UserIndex, 0, "||No podés atacar mascotas en zonas seguras" & FONTTYPE_FIGHT)
                    Exit Sub
            End If
               
            Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex)
            
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No podés atacar a este NPC" & FONTTYPE_FIGHT)
        End If
        
        Call SendUserStatsBox(UserIndex)
        
        Exit Sub
    End If
    
    
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_SWING)
    Call SendUserStatsBox(UserIndex)
End If


End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

Dim ProbRechazo As Long
Dim Rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim UserPoderEvasionEscudo As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim SkillTacticas As Long
Dim SkillDefensa As Long

SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(Tacticas)
SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(Defensa)

Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
proyectil = ObjData(Arma).proyectil = 1

'Calculamos el poder de evasion...
UserPoderEvasion = PoderEvasion(VictimaIndex)

If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
   UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
Else
    UserPoderEvasionEscudo = 0
End If

'Esta usando un arma ???
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
    Else
        PoderAtaque = PoderAtaqueArma(AtacanteIndex)
    End If
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
   
Else
    PoderAtaque = PoderAtaqueWresterling(AtacanteIndex)
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
    
End If
UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
    
    'Fallo ???
    If UsuarioImpacto = False Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If Rechazo = True Then
      'Se rechazo el ataque con el escudo
              Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "TW" & SND_ESCUDO)
              Call SendData(ToIndex, AtacanteIndex, 0, "8")
              Call SendData(ToIndex, VictimaIndex, 0, "7")
              Call SubirSkill(VictimaIndex, Defensa)
      End If
    End If
End If
    
If UsuarioImpacto Then
   If Arma > 0 Then
           If Not proyectil Then
                  Call SubirSkill(AtacanteIndex, Armas)
           Else
                  Call SubirSkill(AtacanteIndex, Proyectiles)
           End If
   Else
        Call SubirSkill(AtacanteIndex, Wresterling)
   End If
End If

End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

If Distancia(UserList(AtacanteIndex).Pos, UserList(VictimaIndex).Pos) > 7 Then
   Call SendData(ToIndex, AtacanteIndex, 0, "||Estás muy lejos para disparar." & FONTTYPE_FIGHT)
   Exit Sub
End If


Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
    Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "TW" & SND_IMPACTO)
    
    If UserList(VictimaIndex).Flags.Navegando = 0 Then Call SendData(ToPCArea, VictimaIndex, UserList(VictimaIndex).Pos.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    
    Call UserDañoUser(AtacanteIndex, VictimaIndex)
Else
    Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).Pos.Map, "TW" & SOUND_SWING)
    Call SendData(ToIndex, AtacanteIndex, 0, "U1")
    Call SendData(ToIndex, VictimaIndex, 0, "U3" & UserList(AtacanteIndex).Name)
End If

End Sub

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim daño As Long, antdaño As Integer
Dim Lugar As Integer, absorbido As Long
Dim defbarco As Integer

Dim Obj As ObjData

daño = CalcularDaño(AtacanteIndex)
antdaño = daño

If UserList(AtacanteIndex).Flags.Navegando = 1 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
     daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
End If

If UserList(VictimaIndex).Flags.Navegando = 1 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If


Lugar = RandomNumber(1, 6)

Select Case Lugar
  
  Case bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 0 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
           Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 0 Then daño = 1
        End If
End Select

Call SendData(ToIndex, AtacanteIndex, 0, "N5" & Lugar & "," & daño & "," & UserList(VictimaIndex).Name)
Call SendData(ToIndex, VictimaIndex, 0, "N4" & Lugar & "," & daño & "," & UserList(AtacanteIndex).Name)

UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - daño

If UserList(AtacanteIndex).Flags.Hambre = 0 And UserList(AtacanteIndex).Flags.Sed = 0 Then
        'Si usa un arma quizas suba "Combate con armas"
        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call SubirSkill(AtacanteIndex, Armas)
        Else
        'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, Wresterling)
        End If
        
        Call SubirSkill(AtacanteIndex, Tacticas)
        
        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(AtacanteIndex) Then
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño)
                Call SubirSkill(AtacanteIndex, Apuñalar)
        End If
End If


If UserList(VictimaIndex).Stats.MinHP <= 0 Then
     
     Call ContarMuerte(VictimaIndex, AtacanteIndex)
     
     ' Para que las mascotas no sigan intentando luchar y
     ' comiencen a seguir al amo
     Dim j As Integer
     For j = 1 To MAXMASCOTAS
        If UserList(AtacanteIndex).MascotasIndex(j) > 0 Then
            If Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = VictimaIndex Then Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = 0
            Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))
        End If
     Next j

     Call ActStats(VictimaIndex, AtacanteIndex)
End If
        

'Controla el nivel del usuario
Call CheckUserLevel(AtacanteIndex)


End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

If UserList(AttackerIndex).GuildInfo.GuildName = "" Or UserList(VictimIndex).GuildInfo.GuildName = "" Then

    If Not Criminal(AttackerIndex) And Not Criminal(VictimIndex) Then
            Call VolverCriminal(AttackerIndex)
    End If
    
    If Not Criminal(VictimIndex) Then
          Call AddtoVar(UserList(AttackerIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
    Else
          Call AddtoVar(UserList(AttackerIndex).Reputacion.NobleRep, vlNoble, MAXREP)
    End If
    
    
Else 'Tiene clan

    If Not UserList(AttackerIndex).GuildRef.IsEnemy(UserList(VictimIndex).GuildInfo.GuildName) Then
        
            If Not Criminal(AttackerIndex) And Not Criminal(VictimIndex) Then
                    Call VolverCriminal(AttackerIndex)
            End If
            
            If Not Criminal(VictimIndex) Then
                  Call AddtoVar(UserList(AttackerIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
            Else
                  Call AddtoVar(UserList(AttackerIndex).Reputacion.NobleRep, vlNoble, MAXREP)
            End If
            
    Else
            
            If Not Criminal(VictimIndex) Then
                  Call AddtoVar(UserList(AttackerIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
            Else
                  Call AddtoVar(UserList(AttackerIndex).Reputacion.NobleRep, vlNoble, MAXREP)
            End If
            
            'Call GiveGuildPoints(1, AttackerIndex, False)
    
    End If
    

End If

Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)


End Sub

Sub AllMascotasAtacanUser(ByVal Victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS
    If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Flags.AttackedBy = UserList(Victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
    End If
Next iCount

End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

If MapInfo(UserList(VictimIndex).Pos.Map).Pk = False Then
    Call SendData(ToIndex, AttackerIndex, 0, "||Esta es una zona segura, aqui no podes atacar otros usuarios." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).Trigger = 4 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes pelear aqui." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

If Not Criminal(VictimIndex) And UserList(AttackerIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||Los soldados del Ejercito Real tienen prohibido atacar ciudadanos." & FONTTYPE_WARNING)
    PuedeAtacar = False
    Exit Function
End If

'Se asegura que la victima no es un GM
If UserList(VictimIndex).Flags.Privilegios >= 2 Then
    SendData ToIndex, AttackerIndex, 0, "||¡¡No podes atacar a los administradores del juego!! " & FONTTYPE_WARNING
    PuedeAtacar = False
    Exit Function
End If

If UserList(VictimIndex).Flags.Muerto = 1 Then
    SendData ToIndex, AttackerIndex, 0, "||No podes atacar a un espiritu" & FONTTYPE_INFO
    PuedeAtacar = False
    Exit Function
End If

If UserList(AttackerIndex).Flags.Muerto = 1 Then
    SendData ToIndex, AttackerIndex, 0, "||No podes atacar porque estas muerto" & FONTTYPE_INFO
    PuedeAtacar = False
    Exit Function
End If

If UserList(AttackerIndex).Flags.Seguro Then
        If Not Criminal(VictimIndex) Then
                Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar ciudadanos, para hacerlo debes desactivar el seguro apretando la tecla S" & FONTTYPE_FIGHT)
                Exit Function
        End If
End If
   

PuedeAtacar = True

End Function


