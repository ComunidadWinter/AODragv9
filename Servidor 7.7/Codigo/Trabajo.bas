Attribute VB_Name = "Trabajo"
Option Explicit

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= -1 Then
                    suerte = 135
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 21 Then
                    suerte = 130
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 41 Then
                    suerte = 128
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 61 Then
                    suerte = 124
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 81 Then
                    suerte = 122
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 120 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 101 Then
                    suerte = 120
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 140 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 121 Then
                    suerte = 118
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 160 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 141 Then
                    suerte = 116
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 180 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 161 Then
                    suerte = 113
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 200 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 181 Then
                    suerte = 110
End If

If UserList(UserIndex).Stats.UserSkills(Ocultarse) = 200 Then suerte = 107

If UCase$(UserList(UserIndex).clase) <> "LADRON" Then suerte = suerte + 10

res = RandomNumber(1, suerte)

If res > 103 Then
   UserList(UserIndex).flags.Oculto = 0
   UserList(UserIndex).flags.Invisible = 0
   UserList(UserIndex).Counters.Invisibilidad = 0
   Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",0")
   Call SendData(ToIndex, UserIndex, 0, "E3")
End If

Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")
End Sub
Public Sub DoOcultarse(ByVal UserIndex As Integer)

On Error GoTo errhandler

If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then
    Exit Sub
End If

Dim suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= -1 Then
                    suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 21 Then
                    suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 41 Then
                    suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 61 Then
                    suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 81 Then
                    suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 120 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 101 Then
                    suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 140 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 121 Then
                    suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 160 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 141 Then
                    suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 180 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 161 Then
                    suerte = 12
ElseIf UserList(UserIndex).Stats.UserSkills(Ocultarse) <= 200 _
   And UserList(UserIndex).Stats.UserSkills(Ocultarse) >= 181 Then
                    suerte = 9
End If
If UserList(UserIndex).Stats.UserSkills(Ocultarse) = 200 Then suerte = 7
If UCase$(UserList(UserIndex).clase) <> "LADRON" Then suerte = suerte + 30

res = RandomNumber(1, suerte)

If res <= 5 Then
   UserList(UserIndex).flags.Oculto = 1
   UserList(UserIndex).flags.Invisible = 1
   Call SendData2(ToMap, 0, UserList(UserIndex).Pos.Map, 17, UserList(UserIndex).Char.CharIndex & ",1")
   Call SendData(ToIndex, UserIndex, 0, "E4")
   Call SubirSkill(UserIndex, Ocultarse)
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData)
On Error GoTo fallo
Dim X As Integer
Dim Y As Integer

'PLUTO:2.4
If UserList(UserIndex).flags.Montura > 0 Or UserList(UserIndex).flags.Angel > 0 Or UserList(UserIndex).flags.Morph > 0 Or UserList(UserIndex).flags.Demonio > 0 Then Exit Sub


Dim ModNave As Long
ModNave = ModNavegacion(UserList(UserIndex).clase)
If UserList(UserIndex).Stats.UserSkills(Navegacion) / ModNave < Barco.MinSkill Then
    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes conocimientos para usar este barco." & "´" & FontTypeNames.FONTTYPE_info)
    Call SendData(ToIndex, UserIndex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & "´" & FontTypeNames.FONTTYPE_info)
    Exit Sub
End If

If UserList(UserIndex).flags.Navegando = 0 Then
    
    UserList(UserIndex).Char.Head = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.Body = Barco.Ropaje
    Else
        UserList(UserIndex).Char.Body = iFragataFantasmal
    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    '[GAU]
    UserList(UserIndex).Char.Botas = NingunBota
    '[GAU]
    UserList(UserIndex).flags.Navegando = 1
      'pluto:6.0A------------
If UserList(UserIndex).Invent.BarcoObjIndex = 474 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + 100
ElseIf UserList(UserIndex).Invent.BarcoObjIndex = 475 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + 300
ElseIf UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax + 500
End If

'-----------------------
Else

'PLUTO:2.4
If HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y) And HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1) And HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1) And HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y) Then
Call SendData(ToIndex, UserIndex, 0, "||No Puedes bajar del barco." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

    UserList(UserIndex).flags.Navegando = 0
      'pluto:6.0A------------
If UserList(UserIndex).Invent.BarcoObjIndex = 474 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax - 100
ElseIf UserList(UserIndex).Invent.BarcoObjIndex = 475 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax - 300
ElseIf UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
UserList(UserIndex).Stats.PesoMax = UserList(UserIndex).Stats.PesoMax - 500
End If

'-----------------------
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
            
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    '[GAU]
        If UserList(UserIndex).Invent.BotaEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.Botas = ObjData(UserList(UserIndex).Invent.BotaEqpObjIndex).Botas
'[GAU]
    
    Else
        If Not Criminal(UserIndex) Then UserList(UserIndex).Char.Body = iCuerpoMuerto Else UserList(UserIndex).Char.Body = iCuerpoMuerto2
        If Not Criminal(UserIndex) Then UserList(UserIndex).Char.Head = iCabezaMuerto Else UserList(UserIndex).Char.Head = iCabezaMuerto2
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    '[GAU]
        UserList(UserIndex).Char.Botas = NingunBota
        '[GAU]
    End If

End If
'[GAU] Agregamo UserList(UserIndex).Char.Botas
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendData2(ToIndex, UserIndex, 0, 6)
'pluto:6.0A------------------------
Call SendUserStatsPeso(UserIndex)
'-----------------------------------
Exit Sub
fallo:
Call LogError("donavega " & Err.number & " D: " & Err.Description)

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
  'pluto:2.14
  If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType <> 23 Then
  'PLUTO:6.3---------------
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'--------------------------
  'Call LogError(" en Jugador:" & UserList(UserIndex).Name & " Bug Fundir " & "Ip: " & UserList(UserIndex).ip & "HD: " & UserList(UserIndex).Serie & " Objeto: " & UserList(UserIndex).flags.TargetObjInvIndex)
'pluto:2.18-------------------
'Dim Tindex As Integer
'Tindex = NameIndex("AoDraGBoT")
'If Tindex <= 0 Then Exit Sub
'Call SendData(ToIndex, Tindex, 0, "|| Jugador: " & UserList(UserIndex).Name & " -> Bug Fundir metal." & "´" & FontTypeNames.FONTTYPE_talk)
'--------------------------

'CloseUser (UserIndex)
Exit Sub
End If
   '-------------
   
   If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(Mineria) / ModFundicion(UserList(UserIndex).clase) Then
        Call DoLingotes(UserIndex)
   Else
        Call SendData(ToIndex, UserIndex, 0, "||No tenes conocimientos de mineria suficientes para trabajar este mineral." & "´" & FontTypeNames.FONTTYPE_info)
   End If

End If
Exit Sub
fallo:
Call LogError("fundirmineral " & Err.number & " D: " & Err.Description)

End Sub
Function TieneObjetos(ByVal itemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
Dim i As Integer
Dim total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = itemIndex Then
        total = total + UserList(UserIndex).Invent.Object(i).Amount
    End If
Next i

If Cant <= total Then
    TieneObjetos = True
    Exit Function
End If
 'pluto:2.10
 TieneObjetos = False
 Exit Function
fallo:
Call LogError("tieneobjetos " & Err.number & " D: " & Err.Description)
       
End Function

Function QuitarObjetos(ByVal itemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
On Error GoTo fallo
Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = itemIndex Then
        
        'pluto:6.0A quito weaponindex=1 no entiendo pq estaba...(lo pongo pq da error al remortear el desequipar recupera atributos bases)
        If UserList(UserIndex).Invent.WeaponEqpObjIndex = 1 Then Call Desequipar(UserIndex, i)
        'Call Desequipar(UserIndex, i)
        
        UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - Cant
'pluto:2.4
UserList(UserIndex).Stats.Peso = UserList(UserIndex).Stats.Peso - (ObjData(UserList(UserIndex).Invent.Object(i).ObjIndex).Peso * Cant)
'pluto:2.4.5
        If UserList(UserIndex).Stats.Peso < 0.001 Then UserList(UserIndex).Stats.Peso = 0

Call SendUserStatsPeso(UserIndex)
        
        Cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
  
       
         'pluto:2-3-04
        If UserList(UserIndex).Invent.Object(i).Amount > 0 Then
        Call UpdateUserInv(False, UserIndex, i)
        Exit Function
        End If

        If UserList(UserIndex).Invent.Object(i).Amount = 0 Then
                UserList(UserIndex).Invent.Object(i).Amount = 0
                UserList(UserIndex).Invent.Object(i).ObjIndex = 0
                QuitarObjetos = True
                 'pluto:hoy
                Call UpdateUserInv(False, UserIndex, i)
                
                Exit Function
        End If
        
        If UserList(UserIndex).Invent.Object(i).Amount < 1 Then
                UserList(UserIndex).Invent.Object(i).Amount = 0
                UserList(UserIndex).Invent.Object(i).ObjIndex = 0
        
        End If
        
        Call UpdateUserInv(False, UserIndex, i)
        
        
    End If
Next i

Exit Function
fallo:
Call LogError("quitarobjetos " & Err.number & " D: " & Err.Description)

End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer)
 On Error GoTo fallo
    If ObjData(itemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(itemIndex).LingH, UserIndex)
    If ObjData(itemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(itemIndex).LingP, UserIndex)
    If ObjData(itemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(itemIndex).LingO, UserIndex)
'pluto:2.10
If ObjData(itemIndex).LingH < 1 And ObjData(itemIndex).LingP < 1 And ObjData(itemIndex).LingO < 1 Then
Call LogCasino("Jugador:" & UserList(UserIndex).Name & "Herrero materiales cero (c) OBJ: " & itemIndex & "Ip: " & UserList(UserIndex).ip)
Exit Sub
End If


Exit Sub
fallo:
Call LogError("herreroquitarmateriales " & Err.number & " D: " & Err.Description)

End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer)
  On Error GoTo fallo
    
If ObjData(itemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(itemIndex).Madera, UserIndex)

 'pluto:2.10
If ObjData(itemIndex).Madera < 1 Then
Call LogCasino("Jugador:" & UserList(UserIndex).Name & "Carpinterp materiales cero(c) OBJ: " & itemIndex & "Ip: " & UserList(UserIndex).ip)
Exit Sub
End If


Exit Sub
fallo:
Call LogError("carpinteroquitarmateriales " & Err.number & " D: " & Err.Description)

End Sub

'[MerLiNz:6]
Sub ermitanoQuitarMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer)
    On Error GoTo fallo
    If ObjData(itemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(itemIndex).Madera, UserIndex)
    'pluto:2.4.5
    If ObjData(itemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(itemIndex).LingH, UserIndex)

    If ObjData(itemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(itemIndex).LingO, UserIndex)
    If ObjData(itemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(itemIndex).LingP, UserIndex)
    If ObjData(itemIndex).Gemas > 0 Then Call QuitarObjetos(GemaI, ObjData(itemIndex).Gemas, UserIndex)
    If ObjData(itemIndex).Diamantes > 0 Then Call QuitarObjetos(Diamante, ObjData(itemIndex).Diamantes, UserIndex)

'pluto:2.10
If ObjData(itemIndex).Madera < 1 And ObjData(itemIndex).LingH < 1 And ObjData(itemIndex).LingP < 1 And ObjData(itemIndex).LingO < 1 And ObjData(itemIndex).Gemas < 1 And ObjData(itemIndex).Diamantes < 1 Then
Call LogCasino("Jugador:" & UserList(UserIndex).Name & " Ermitaño materiales cero OBJ: " & itemIndex & "Ip: " & UserList(UserIndex).ip)
Exit Sub
End If

Exit Sub
fallo:
Call LogError("ermitañoquitarmateriales " & Err.number & " D: " & Err.Description)

End Sub


'[MerLiNz:6]
Function ermitanoTieneMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer) As Boolean
   On Error GoTo fallo
    If ObjData(itemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(itemIndex).Madera, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes madera." & "´" & FontTypeNames.FONTTYPE_info)
                    ermitanoTieneMateriales = False
                    Exit Function
            End If
    End If
    'pluto:2.4.5
    If ObjData(itemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(itemIndex).LingH, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes Hierro." & "´" & FontTypeNames.FONTTYPE_info)
                    ermitanoTieneMateriales = False
                    Exit Function
            End If
    End If
    
    If ObjData(itemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(itemIndex).LingP, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes plata." & "´" & FontTypeNames.FONTTYPE_info)
                    ermitanoTieneMateriales = False
                    Exit Function
            End If
    End If
    
    If ObjData(itemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(itemIndex).LingO, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes oro." & "´" & FontTypeNames.FONTTYPE_info)
                    ermitanoTieneMateriales = False
                    Exit Function
            End If
    End If
    
    If ObjData(itemIndex).Gemas > 0 Then
            If Not TieneObjetos(GemaI, ObjData(itemIndex).Gemas, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes gemas." & "´" & FontTypeNames.FONTTYPE_info)
                    ermitanoTieneMateriales = False
                    Exit Function
            End If
    End If
    
    If ObjData(itemIndex).Diamantes > 0 Then
            If Not TieneObjetos(Diamante, ObjData(itemIndex).Diamantes, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes diamantes." & "´" & FontTypeNames.FONTTYPE_info)
                    ermitanoTieneMateriales = False
                    Exit Function
            End If
    End If
  
  'pluto:2.10
If ObjData(itemIndex).Madera < 1 And ObjData(itemIndex).LingH < 1 And ObjData(itemIndex).LingP < 1 And ObjData(itemIndex).LingO < 1 And ObjData(itemIndex).Gemas < 1 And ObjData(itemIndex).Diamantes < 1 Then
Call LogCasino("Jugador:" & UserList(UserIndex).Name & "Ermitaño materiales cero (b) OBJ: " & itemIndex & "Ip: " & UserList(UserIndex).ip)
ermitanoTieneMateriales = False
Exit Function
End If

  
ermitanoTieneMateriales = True
'[\END]
Exit Function
fallo:
Call LogError("ermitañotienemateriales " & Err.number & " D: " & Err.Description)

End Function



Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer) As Boolean
    On Error GoTo fallo
    If ObjData(itemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(itemIndex).Madera, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes madera." & "´" & FontTypeNames.FONTTYPE_info)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
   'pluto:2.10
If ObjData(itemIndex).Madera < 1 Then
Call LogCasino("Jugador:" & UserList(UserIndex).Name & "Carpintero materiales cero (A) OBJ: " & itemIndex & "Ip: " & UserList(UserIndex).ip)
CarpinteroTieneMateriales = False
Exit Function
End If
    
    
    CarpinteroTieneMateriales = True
Exit Function
fallo:
Call LogError("carpinterotienemateriales " & Err.number & " D: " & Err.Description)

End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer) As Boolean
    On Error GoTo fallo
    If ObjData(itemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(itemIndex).LingH, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes lingotes de hierro." & "´" & FontTypeNames.FONTTYPE_info)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(itemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(itemIndex).LingP, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes lingotes de plata." & "´" & FontTypeNames.FONTTYPE_info)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(itemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(itemIndex).LingO, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes lingotes de oro." & "´" & FontTypeNames.FONTTYPE_info)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    'pluto:2.10
If ObjData(itemIndex).LingH < 1 And ObjData(itemIndex).LingP < 1 And ObjData(itemIndex).LingO < 1 Then
Call LogCasino("Jugador:" & UserList(UserIndex).Name & "Herrero materiales cero (b) OBJ: " & itemIndex & "Ip: " & UserList(UserIndex).ip)
HerreroTieneMateriales = False
Exit Function
End If
    
    HerreroTieneMateriales = True
Exit Function
fallo:
Call LogError("herrerotienemateriales " & Err.number & " D: " & Err.Description)

End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal itemIndex As Integer) As Boolean
On Error GoTo fallo
PuedeConstruir = HerreroTieneMateriales(UserIndex, itemIndex) And UserList(UserIndex).Stats.UserSkills(Herreria) >= _
 ObjData(itemIndex).SkHerreria

Exit Function
fallo:
Call LogError("puedeconstruir " & Err.number & " D: " & Err.Description)

End Function



Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal itemIndex As Integer)
On Error GoTo fallo

If PuedeConstruir(UserIndex, itemIndex) Then
    Call HerreroQuitarMateriales(UserIndex, itemIndex)
    ' AGREGAR FX
    If ObjData(itemIndex).OBJType = OBJTYPE_WEAPON Then
        Call SendData(ToIndex, UserIndex, 0, "||Has construido el arma!." & "´" & FontTypeNames.FONTTYPE_info)
    ElseIf ObjData(itemIndex).OBJType = OBJTYPE_ESCUDO Then
        Call SendData(ToIndex, UserIndex, 0, "||Has construido el escudo!." & "´" & FontTypeNames.FONTTYPE_info)
    ElseIf ObjData(itemIndex).OBJType = OBJTYPE_CASCO Then
        Call SendData(ToIndex, UserIndex, 0, "||Has construido el casco!." & "´" & FontTypeNames.FONTTYPE_info)
    ElseIf ObjData(itemIndex).OBJType = OBJTYPE_ARMOUR Then
        Call SendData(ToIndex, UserIndex, 0, "||Has construido la armadura!." & "´" & FontTypeNames.FONTTYPE_info)
    '[GAU]
    ElseIf ObjData(itemIndex).OBJType = OBJTYPE_BOTA Then
        Call SendData(ToIndex, UserIndex, 0, "||Has construido las botas!." & "´" & FontTypeNames.FONTTYPE_info)
'[GAU]
    End If

    'PLUTO:6.0a
If ObjData(itemIndex).ParaHerre = 0 Then
Call LogNpcFundidor("Nombre: " & UserList(UserIndex).Name & " intenta fabricar Obj: " & itemIndex & " con herrero.")
Exit Sub
End If
    
    Dim MiObj As obj
    MiObj.Amount = 1
    MiObj.ObjIndex = itemIndex
    
'pluto:6.0A
Call LogNpcFundidor("Nombre: " & UserList(UserIndex).Name & " fabrica Obj: " & itemIndex)

    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    Call LogCasino("Jugador:" & UserList(UserIndex).Name & " fabrica herrero inventario lleno(A) " & itemIndex & "Ip: " & UserList(UserIndex).ip)
UserList(UserIndex).Alarma = 1
UserList(UserIndex).ObjetosTirados = UserList(UserIndex).ObjetosTirados + 1

    End If
    'pluto.2.4.1
UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + (CInt((UserList(UserIndex).Stats.ELV / 10) + 1) * MiObj.Amount)
Call CheckUserLevel(UserIndex)

    Call SubirSkill(UserIndex, Herreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MARTILLOHERRERO)
    
End If
Exit Sub
fallo:
Call LogError("herreroconstruyeitem " & Err.number & " D: " & Err.Description)

End Sub
Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal itemIndex As Integer)
On Error GoTo fallo
Dim MiObj As obj
Dim X As Integer
If CarpinteroTieneMateriales(UserIndex, itemIndex) And _
   UserList(UserIndex).Stats.UserSkills(Carpinteria) >= _
   ObjData(itemIndex).SkCarpinteria Then


'PLUTO:6.0a
If ObjData(itemIndex).ParaCarpin = 0 Then
Call LogNpcFundidor("Nombre: " & UserList(UserIndex).Name & " intenta fabricar Obj: " & itemIndex & " con carpintero.")
Exit Sub
End If



'pluto:2.14---------------------------
If (ObjData(itemIndex).OBJType = OBJTYPE_FLECHAS) Then
For X = 1 To UserList(UserIndex).Stats.ELV
    If CarpinteroTieneMateriales(UserIndex, itemIndex) And _
        UserList(UserIndex).Stats.UserSkills(Carpinteria) >= _
        ObjData(itemIndex).SkCarpinteria Then
       
        Call CarpinteroQuitarMateriales(UserIndex, itemIndex)
    Else
        Exit For
    End If 'tienematerial
Next X
        MiObj.Amount = X - 1
        MiObj.ObjIndex = itemIndex
If MiObj.Amount > 0 Then
            
If Not MeterItemEnInventario(UserIndex, MiObj) Then
            'pluto:2.9.0
UserList(UserIndex).ObjetosTirados = UserList(UserIndex).ObjetosTirados + 1
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
Call LogCasino("Jugador:" & UserList(UserIndex).Name & " fabrica flecha carpintero inventario lleno(C) " & itemIndex & "Ip: " & UserList(UserIndex).ip)
UserList(UserIndex).Alarma = 1
End If 'meter invent
End If 'amount>0
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)

Else 'no flechas

'----------------------------------


    Call CarpinteroQuitarMateriales(UserIndex, itemIndex)
    Call SendData(ToIndex, UserIndex, 0, "E5")
    
    'Dim MiObj As obj
    MiObj.Amount = 1
    MiObj.ObjIndex = itemIndex
'pluto:6.0A
If itemIndex <> 163 And itemIndex <> 960 Then
Call LogNpcFundidor("Nombre: " & UserList(UserIndex).Name & " fabrica Obj: " & itemIndex)
End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    Call LogCasino("Jugador:" & UserList(UserIndex).Name & " fabrica carpintero inventario lleno(A) " & itemIndex & "Ip: " & UserList(UserIndex).ip)
UserList(UserIndex).ObjetosTirados = UserList(UserIndex).ObjetosTirados + 1
UserList(UserIndex).Alarma = 1
    End If
    'pluto.2.4.1
UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + (CInt((UserList(UserIndex).Stats.ELV / 10) + 1) * MiObj.Amount)
Call CheckUserLevel(UserIndex)

    Call SubirSkill(UserIndex, Carpinteria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
End If

End If 'flechas

Exit Sub
fallo:
Call LogError("carpinteroconstruyeitem " & Err.number & " D: " & Err.Description)

End Sub

'[MerLiNz:6]
Public Sub ermitanoConstruirItem(ByVal UserIndex As Integer, ByVal itemIndex As Integer)
On Error GoTo fallo
Dim MiObj As obj
Dim X As Integer
Dim cons As Boolean
cons = False
'pluto:2.10
 If ermitanoTieneMateriales(UserIndex, itemIndex) And _
        UserList(UserIndex).Stats.UserSkills(Carpinteria) >= _
        ObjData(itemIndex).SkCarpinteria And UserList(UserIndex).Stats.UserSkills(Herreria) >= _
        ObjData(itemIndex).SkHerreria Then

'PLUTO:6.0a
If ObjData(itemIndex).ParaErmi = 0 Then
Call LogNpcFundidor("Nombre: " & UserList(UserIndex).Name & " intenta fabricar Obj: " & itemIndex & " con ermitaño.")
Exit Sub
End If

If (ObjData(itemIndex).OBJType = OBJTYPE_FLECHAS) Then
    For X = 1 To 10
        If ermitanoTieneMateriales(UserIndex, itemIndex) And _
        UserList(UserIndex).Stats.UserSkills(Carpinteria) >= _
        ObjData(itemIndex).SkCarpinteria And UserList(UserIndex).Stats.UserSkills(Herreria) >= _
        ObjData(itemIndex).SkHerreria Then
        cons = True
        Call ermitanoQuitarMateriales(UserIndex, itemIndex)
        Else
        Exit For
        End If
    Next X
        MiObj.Amount = X - 1
        MiObj.ObjIndex = itemIndex
        If MiObj.Amount > 0 Then
            
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
            'pluto:2.9.0
UserList(UserIndex).ObjetosTirados = UserList(UserIndex).ObjetosTirados + 1
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
Call LogCasino("Jugador:" & UserList(UserIndex).Name & " fabrica ermitaño inventario lleno(C) " & itemIndex & "Ip: " & UserList(UserIndex).ip)
UserList(UserIndex).Alarma = 1

                End If 'meter
        End If 'amount>0
        If (ObjData(itemIndex).SkCarpinteria > 0) Then Call SubirSkill(UserIndex, Carpinteria)
        If (ObjData(itemIndex).SkHerreria > 0) Then Call SubirSkill(UserIndex, Herreria)
        Call UpdateUserInv(True, UserIndex, 0)
Else ' no flechas
    If ermitanoTieneMateriales(UserIndex, itemIndex) And _
        UserList(UserIndex).Stats.UserSkills(Carpinteria) >= _
        ObjData(itemIndex).SkCarpinteria And UserList(UserIndex).Stats.UserSkills(Herreria) >= _
        ObjData(itemIndex).SkHerreria Then

        Call ermitanoQuitarMateriales(UserIndex, itemIndex)
        Call SendData(ToIndex, UserIndex, 0, "E5")
        'pluto:6.0A
        Call LogNpcFundidor("Nombre: " & UserList(UserIndex).Name & " fabrica Obj: " & itemIndex)
        
        MiObj.Amount = 1
        MiObj.ObjIndex = itemIndex
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
        'Call Encarcelar(UserIndex, 10)
        Call LogCasino("Jugador:" & UserList(UserIndex).Name & " fabrica ermitaño inventario lleno(B) " & itemIndex & "Ip: " & UserList(UserIndex).ip)
        'Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
        'pluto:2.9.0
        UserList(UserIndex).ObjetosTirados = UserList(UserIndex).ObjetosTirados + 1
        UserList(UserIndex).Alarma = 1

        End If
        
        If (ObjData(itemIndex).SkCarpinteria > 0) Then Call SubirSkill(UserIndex, Carpinteria)
        If (ObjData(itemIndex).SkHerreria > 0) Then Call SubirSkill(UserIndex, Herreria)
        'pluto.2.4.1
UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + (CInt((UserList(UserIndex).Stats.ELV / 10) + 1) * MiObj.Amount)
Call CheckUserLevel(UserIndex)

        Call UpdateUserInv(True, UserIndex, 0)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
    End If
End If

If (cons = True) Then
    Call SendData(ToIndex, UserIndex, 0, "E5")
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
End If
'[\END]

End If ' pluto:2.10
Exit Sub
fallo:
Call LogError("ermitañoconstruyeitem " & Err.number & " D: " & Err.Description)

End Sub


Public Sub DoLingotes(ByVal UserIndex As Integer)
 On Error GoTo fallo
    'pluto:2.6.0 lingotes de 5 en 5 y 25 materiales a lo largo de todo el sub
    
    If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount < 25 Then
              Call SendData(ToIndex, UserIndex, 0, "||No tienes suficientes minerales para hacer lingotes." & "´" & FontTypeNames.FONTTYPE_info)
              Exit Sub
    End If

'pluto:6.7--------
If ObjData(UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex).OBJType <> 23 Then
Call SendData(ToIndex, UserIndex, 0, "||No tienes suficientes minerales para hacer lingotes." & "´" & FontTypeNames.FONTTYPE_info)
Call LogError("Posible Bug hacer lingotes en " & UserList(UserIndex).Name)
Exit Sub
End If
'-----------------
   'pluto:2.4  posibilidad hacer lingotes con skill suerte
    If RandomNumber(1, ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill) < 10 + CInt(UserList(UserIndex).Stats.UserSkills(suerte) / 10) + CInt(UserList(UserIndex).Stats.ELV / 2) Then
       
                UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount - 25
                If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount < 1 Then
                    UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0
                    UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex = 0
                End If
                Call SendData(ToIndex, UserIndex, 0, "E6")
                Dim nPos As WorldPos
                Dim MiObj As obj
                MiObj.Amount = 5
                MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
                End If
                Call UpdateUserInv(False, UserIndex, UserList(UserIndex).flags.TargetObjInvSlot)
'pluto.2.4.1
UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + (CInt((UserList(UserIndex).Stats.ELV / 10) + 1) * MiObj.Amount)
Call CheckUserLevel(UserIndex)

                'Call SendData(ToIndex, UserIndex, 0, "||¡Has obtenido cinco lingotes!" & FONTTYPENAMES.FONTTYPE_INFO)
    Else
        
        UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount - 25
        If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount < 1 Then
                UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0
                UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex = 0
        End If
        Call UpdateUserInv(False, UserIndex, UserList(UserIndex).flags.TargetObjInvSlot)
        Call SendData(ToIndex, UserIndex, 0, "E7")
    End If
    Exit Sub
fallo:
Call LogError("dolingotes " & Err.number & " D: " & Err.Description)

End Sub

Function ModNavegacion(ByVal clase As String) As Integer
On Error GoTo fallo
Select Case UCase$(clase)
    Case "PIRATA"
        ModNavegacion = 1
    Case "PESCADOR"
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select
Exit Function
fallo:
Call LogError("modnavegacion " & Err.number & " D: " & Err.Description)

End Function


Function ModFundicion(ByVal clase As String) As Integer
On Error GoTo fallo
Select Case UCase$(clase)
    Case "MINERO"
        ModFundicion = 1
    Case "HERRERO"
        ModFundicion = 1.2
    Case "ERMITAÑO"
        ModFundicion = 1.6
    Case Else
        ModFundicion = 3
End Select
Exit Function
fallo:
Call LogError("modfundicion " & Err.number & " D: " & Err.Description)

End Function

Function ModCarpinteria(ByVal clase As String) As Integer
On Error GoTo fallo
Select Case UCase$(clase)
    Case "CARPINTERO"
        ModCarpinteria = 1
    Case "ERMITAÑO"
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select
Exit Function
fallo:
Call LogError("modcarpinteria " & Err.number & " D: " & Err.Description)

End Function

Function ModHerreriA(ByVal clase As String) As Integer
On Error GoTo fallo
Select Case UCase$(clase)
    Case "HERRERO"
        ModHerreriA = 1
    Case "MINERO"
        ModHerreriA = 1.2
    Case "ERMITAÑO"
        ModHerreriA = 1
    Case Else
        ModHerreriA = 4
End Select
Exit Function
fallo:
Call LogError("modherreria " & Err.number & " D: " & Err.Description)

End Function
'pluto:2.4.5
Function ModMagia(ByVal clase As String) As Single
On Error GoTo fallo
Select Case UCase$(clase)
    Case "MAGO"
        ModMagia = 1
    Case "DRUIDA"
        ModMagia = 1
    Case "BARDO"
        ModMagia = 1
    Case "CLERIGO"
        ModMagia = 1
    Case "ASESINO"
        ModMagia = 1
    Case "PALADIN"
        ModMagia = 1
    Case "GUERRERO"
        ModMagia = 1
    Case "CAZADOR"
        ModMagia = 1
    Case "ARQUERO"
        ModMagia = 1
    Case Else
    'nati: cambio el modmagia 0.9 por 1 porque no se puede dividir con él.
        ModMagia = 1
End Select
Exit Function
fallo:
Call LogError("modmagia " & Err.number & " D: " & Err.Description)

End Function

Function ModDomar(ByVal clase As String) As Integer
On Error GoTo fallo
Select Case UCase$(clase)
    'pluto:2.3
    Case "DOMADOR"
        ModDomar = 8
    Case "DRUIDA"
        ModDomar = 12
    Case "CAZADOR"
        ModDomar = 12
    Case "CLERIGO"
        ModDomar = 14
    Case Else
        ModDomar = 20
End Select
Exit Function
fallo:
Call LogError("moddomar " & Err.number & " D: " & Err.Description)

End Function

Function CalcularPoderDomador(ByVal UserIndex As Integer) As Long
On Error GoTo fallo
CalcularPoderDomador = _
UserList(UserIndex).Stats.UserAtributos(Carisma) * _
(CInt(UserList(UserIndex).Stats.UserSkills(Domar) / 2) / ModDomar(UserList(UserIndex).clase)) _
+ RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Carisma) / 3) _
+ RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Carisma) / 3) _
+ RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Carisma) / 3)

Exit Function
fallo:
Call LogError("calcularpoderdomador " & Err.number & " D: " & Err.Description)

End Function
Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
On Error GoTo fallo
'Call LogTarea("Sub FreeMascotaIndex")
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) = 0 Then
        FreeMascotaIndex = j
        Exit Function
    End If
Next j
Exit Function
fallo:
Call LogError("freemascotaindex " & Err.number & " D: " & Err.Description)

End Function
Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'Call LogTarea("Sub DoDomar")

On Error GoTo fallo
Dim nPos As WorldPos
Dim MiObj As obj
Dim n As Byte
Dim tc As Integer
Dim userfile As String

'PLUTO:6.3---------------
If NpcIndex = 0 Then
'If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
Exit Sub
'End If
End If
If Npclist(NpcIndex).MaestroUser > 0 Then
'If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
Exit Sub
'End If
End If
'--------------------------

userfile = CharPath & Left$(UserList(UserIndex).Name, 1) & "\" & UserList(UserIndex).Name & ".chr"


If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call SendData(ToIndex, UserIndex, 0, "||La criatura ya te ha aceptado como su amo." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If
    
     If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||La criatura ya tiene amo." & "´" & FontTypeNames.FONTTYPE_info)
        Exit Sub
    End If

'pluto:6.0A
If UserList(UserIndex).Nmonturas > 2 Then
Call SendData(ToIndex, UserIndex, 0, "||No puedes tener más de 3 Mascotas." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If

'quitar esto
If UserList(UserIndex).flags.Privilegios > 2 Then GoTo domo
     
     'pluto:2.4.1
If Npclist(NpcIndex).NPCtype = 60 And Npclist(NpcIndex).flags.Domable <> 506 And (UCase$(UserList(UserIndex).clase) <> "DOMADOR" Or UserList(UserIndex).Stats.UserSkills(Domar) < Npclist(NpcIndex).SkillDomar Or UserList(UserIndex).Stats.ELV < 40) Then
Call SendData(ToIndex, UserIndex, 0, "P1")
Exit Sub
End If

If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(UserIndex) Or Npclist(NpcIndex).NPCtype = 60 Then

'pluto:2.4.1
 If Npclist(NpcIndex).NPCtype = 60 Then
 
 'pluto:2.18.Domable
 If UserList(UserIndex).Stats.UserSkills(Domar) < 200 And Npclist(NpcIndex).flags.Domable <> 506 Then
 Call SendData(ToIndex, UserIndex, 0, "P1")
 Exit Sub
 End If
 

 Dim aa As Integer
 aa = RandomNumber(1, (Npclist(NpcIndex).Stats.MaxHP * 5))
 
 'quitar esto
 'server secundario cambio <>20 por >10 para facilitar el domar
 'If aa <> 20 Then
'If ServerPrimario = 2 Then
 '   If aa > 10 Then
  '  Call SendData(ToIndex, UserIndex, 0, "P2")
   ' Exit Sub
    'End If
'Else
    If aa <> 20 Then
    Call SendData(ToIndex, UserIndex, 0, "P2")
    Exit Sub
    End If
'End If

'pluto:6.0A---------------------------
tc = Npclist(NpcIndex).flags.Domable + 387
MiObj.Amount = 1
MiObj.ObjIndex = tc
If TieneObjetos(tc, 1, UserIndex) Then
Call SendData(ToIndex, UserIndex, 0, "||Ya tienes esa clase de mascota." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
'miramos que no repita mascota
For n = 1 To 3
If val(GetVar(userfile, "MONTURA" & n, "TIPO")) = Npclist(NpcIndex).flags.Domable - 500 Then
Call SendData(ToIndex, UserIndex, 0, "||Ya tienes esa clase de mascota, ve a la cuidadora de mascotas en Banderbill a recuperarla." & "´" & FontTypeNames.FONTTYPE_info)
Exit Sub
End If
Next n
 '----------------------------------------------------------
 
 
 End If
'quitar esto
domo:
'pluto:2.4
Dim MinPc As npc
MinPc = Npclist(NpcIndex)
If Npclist(NpcIndex).NPCtype = 60 And MinPc.MaestroUser = 0 Then



Call DomarMontura(UserIndex, NpcIndex)
'pluto:6.5
If NoDomarMontura = True Then
NoDomarMontura = False
Exit Sub
End If

Dim CabalgaPos As WorldPos
Dim mapita As Integer
 Dim ini As Integer
    
      'evitamos respawn otro mapa del jabato
    If MinPc.flags.Domable = 506 Then
     MinPc.flags.Respawn = 0
     Call ReSpawnNpc(MinPc)
    Exit Sub
    End If
    
    
    CabalgaPos.X = 50
    CabalgaPos.Y = 50
a:
mapita = RandomNumber(1, 270)
CabalgaPos.Map = mapita
'If MapInfo(CabalgaPos.Map).Pk = False Or MapInfo(CabalgaPos.Map).BackUp = 1 Or MapInfo(CabalgaPos.Map).Terreno <> "BOSQUE" Then GoTo a:
If MapInfo(CabalgaPos.Map).Domar > 0 Then GoTo a:
ini = SpawnNpc(MinPc.numero, CabalgaPos, False, True)
If ini = MAXNPCS Then GoTo a
Call WriteVar(IniPath & "cabalgar.txt", MinPc.Name, "Mapa", val(mapita))
Exit Sub
End If
 '---fin pluto:2.4----
 
        Dim index As Integer
        UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
        index = FreeMascotaIndex(UserIndex)
  'pluto:2.4
  If index = 0 Then Exit Sub
  
  UserList(UserIndex).MascotasIndex(index) = NpcIndex
        UserList(UserIndex).MascotasType(index) = Npclist(NpcIndex).numero
        
        Npclist(NpcIndex).MaestroUser = UserIndex
        
        Call FollowAmo(NpcIndex)
        
        Call SendData(ToIndex, UserIndex, 0, "||La criatura te ha aceptado como su amo." & "´" & FontTypeNames.FONTTYPE_info)
        Call SubirSkill(UserIndex, Domar)
 'PLUTO:6.3
If UserList(UserIndex).flags.Macreanda > 0 Then
UserList(UserIndex).flags.ComproMacro = 0
UserList(UserIndex).flags.Macreanda = 0
Call SendData(ToIndex, UserIndex, 0, "O3")
End If
'---------------------------
        
        
        
    'pluto:2.4 respawn de los domados
    If Npclist(NpcIndex).NPCtype <> 60 Then
    Call ReSpawnNpc(MinPc)
    End If

   Else
        Call SendData(ToIndex, UserIndex, 0, "P3")
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||No podes controlar mas criaturas." & "´" & FontTypeNames.FONTTYPE_info)
End If
Exit Sub
fallo:
Call LogError("dodomar " & UserList(UserIndex).Name & " " & Err.number & " D: " & Err.Description)


End Sub

Sub DoAdminInvisible(ByVal UserIndex As Integer)

If UserList(UserIndex).flags.AdminInvisible = 0 Then

        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.Invisible = 1
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.Body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.Body = 0
        UserList(UserIndex).Char.Head = 0
        
    Else
        
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.Invisible = 0
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).Char.Body = UserList(UserIndex).flags.OldBody
        UserList(UserIndex).Char.Head = UserList(UserIndex).flags.OldHead
        
    End If
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "16" & UserList(UserIndex).Char.CharIndex & ",0")
End Sub
   

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error GoTo fallo
Dim suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim obj As obj

If Not LegalPos(Map, X, Y) Then Exit Sub

If MapData(Map, X, Y).OBJInfo.Amount < 3 Then
    Call SendData(ToIndex, UserIndex, 0, "K9")
    Exit Sub
End If


If UserList(UserIndex).Stats.UserSkills(Supervivencia) < 50 Then
            suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 50 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 120 Then
            suerte = 5
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 120 Then
            suerte = 2
End If

exito = RandomNumber(1, suerte)

If exito = 1 Then
    obj.ObjIndex = FOGATA_APAG
    obj.Amount = MapData(Map, X, Y).OBJInfo.Amount / 3
    
    If obj.Amount > 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||Has hecho " & obj.Amount & " fogatas." & "´" & FontTypeNames.FONTTYPE_info)
    Else
        Call SendData(ToIndex, UserIndex, 0, "K7")
    End If
    
    Call MakeObj(ToMap, 0, Map, obj, Map, X, Y)
    
    Dim Fogatita As New cGarbage
    Fogatita.Map = Map
    Fogatita.X = X
    Fogatita.Y = Y
    Call TrashCollector.Add(Fogatita)
    
Else
    Call SendData(ToIndex, UserIndex, 0, "K8")
End If

Call SubirSkill(UserIndex, Supervivencia)

Exit Sub
fallo:
Call LogError("tratarhacerfogata " & Err.number & " D: " & Err.Description)

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim suerte As Integer
Dim res As Integer
 'pluto:2.12
 UserList(UserIndex).Counters.IdleCount = 0
 
If UserList(UserIndex).clase = "Pescador" Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(Pesca) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= -1 Then
                    suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 21 Then
                    suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 41 Then
                    suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 61 Then
                    suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 81 Then
                    suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 120 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 101 Then
                    suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 140 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 121 Then
                    suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 160 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 141 Then
                    suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 180 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 161 Then
                    suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(Pesca) <= 200 _
   And UserList(UserIndex).Stats.UserSkills(Pesca) >= 181 Then
                    suerte = 10
End If
If UserList(UserIndex).Stats.UserSkills(Pesca) = 200 Then suerte = 7

res = RandomNumber(1, suerte)

'PLuto:2.4
Dim res2 As Integer
res2 = RandomNumber(1, 600)
    Dim nPos As WorldPos
    Dim MiObj As obj
  'pluto:2.4.1
  If res2 > 597 Then
  MiObj.Amount = 1
 MiObj.ObjIndex = 887
 
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
Exit Sub
End If
  
  If res < 6 Or res2 < UserList(UserIndex).Stats.UserSkills(1) Then
  'pluto:2.4.5
  If res > 5 Then res = 1
    If UserList(UserIndex).clase = "Pescador" Then
        MiObj.Amount = RandomNumber(1, CInt(UserList(UserIndex).Stats.ELV / 2))
    Else
        MiObj.Amount = 1
    End If
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 543 Then MiObj.Amount = MiObj.Amount * 2
   
    If MiObj.Amount < 1 Then MiObj.Amount = 1
   If res = 1 Then MiObj.ObjIndex = Pescado3
   If res = 2 Then MiObj.ObjIndex = Pescado2
   If res > 2 And res < 6 Then MiObj.ObjIndex = Pescado

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "G3")
'pluto.2.4.1
UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + (CInt((UserList(UserIndex).Stats.ELV / 10) + 1) * MiObj.Amount)
Call CheckUserLevel(UserIndex)
Else
    Call SendData(ToIndex, UserIndex, 0, "G4")
End If

Call SubirSkill(UserIndex, Pesca)


Exit Sub

errhandler:
    Call LogError("Error en DoPescar")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
On Error GoTo errhandler

If MapInfo(UserList(VictimaIndex).Pos.Map).Pk = False Then Exit Sub
'pluto:2.18
If MapInfo(UserList(VictimaIndex).Pos.Map).Terreno = "ALDEA" Or MapInfo(UserList(VictimaIndex).Pos.Map).Terreno = "TORNEO" Then Exit Sub

'pluto:6.2
If UserList(VictimaIndex).Name = "Jaba" Then Exit Sub

If UserList(VictimaIndex).Pos.Map = MapaSeguro Then Exit Sub
If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Exit Sub

If UserList(VictimaIndex).flags.Privilegios < 1 Then
    Dim suerte As Integer
    Dim res As Integer
       
    If UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= -1 Then
                        suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 21 Then
                        suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 41 Then
                        suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 61 Then
                        suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 81 Then
                        suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 120 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 101 Then
                        suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 140 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 121 Then
                        suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 160 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 141 Then
                        suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 180 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 161 Then
                        suerte = 11
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(Robar) <= 200 _
       And UserList(LadrOnIndex).Stats.UserSkills(Robar) >= 181 Then
                        suerte = 7
    End If
    If UserList(LadrOnIndex).Stats.UserSkills(Robar) = 200 Then suerte = 5
    
    res = RandomNumber(1, suerte)
    
    If res < 4 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 18) And (UCase$(UserList(LadrOnIndex).clase) = "LADRON") Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call SendData(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene objetos." & "´" & FontTypeNames.FONTTYPE_info)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim n As Integer
                
                n = RandomNumber(1, 100)
                If UCase$(UserList(LadrOnIndex).clase) = "LADRON" Then n = n + 1000
                If UCase$(UserList(LadrOnIndex).clase) = "BANDIDO" Then n = n + 2500
                If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                
                Call AddtoVar(UserList(LadrOnIndex).Stats.GLD, n, MAXORO)
                
                Call SendData(ToIndex, LadrOnIndex, 0, "||Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
            'pluto:2.4.5
            Call SendUserStatsOro(LadrOnIndex)
            Call SendUserStatsOro(VictimaIndex)
Else
                Call SendData(ToIndex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).Name & " no tiene oro." & "´" & FontTypeNames.FONTTYPE_info)
            End If
        End If
    Else
        Call SendData(ToIndex, LadrOnIndex, 0, "||¡No has logrado robar nada!" & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " ha intentado robarte!" & "´" & FontTypeNames.FONTTYPE_info)
        Call SendData(ToIndex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).Name & " es un criminal!" & "´" & FontTypeNames.FONTTYPE_info)
    End If

    If Not Criminal(LadrOnIndex) Then
            Call VolverCriminal(LadrOnIndex)
    End If
    
    'If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)
'If UserList(LadrOnIndex).Faccion.ArmadaReal = 2 Then Call ExpulsarFaccionlegion(LadrOnIndex)

    Call AddtoVar(UserList(LadrOnIndex).Reputacion.LadronesRep, vlLadron, MAXREP)
    Call SubirSkill(LadrOnIndex, Robar)

End If

'pluto:2.5.0
Exit Sub

errhandler:
    Call LogError("Error en DoRobar")

End Sub


Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
On Error GoTo fallo
Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).OBJType <> OBJTYPE_LLAVES And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).nocaer = 0 And _
ObjData(OI).Caos = 0

'pluto:robo barcos equipados
If ObjData(OI).OBJType = OBJTYPE_BARCOS And UserList(VictimaIndex).flags.Navegando = 1 Then ObjEsRobable = False
'pluto:roba ropas cabalgar equipados
If ObjData(OI).OBJType = 42 And UserList(VictimaIndex).flags.Montura > 0 Then ObjEsRobable = False

Exit Function
fallo:
Call LogError("objesrobable " & Err.number & " D: " & Err.Description)

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
On Error GoTo fallo
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
    Dim MiObj As obj
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
    
    Call SendData(ToIndex, LadrOnIndex, 0, "||Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & "´" & FontTypeNames.FONTTYPE_info)
Else
    Call SendData(ToIndex, LadrOnIndex, 0, "||No has robado nada" & "´" & FontTypeNames.FONTTYPE_info)
End If
Exit Sub
fallo:
Call LogError("robarobjeto " & Err.number & " D: " & Err.Description)

End Sub
Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
On Error GoTo fallo
Dim suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= -1 Then
                    suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 21 Then
                    suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 41 Then
                    suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 61 Then
                    suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 81 Then
                    suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 120 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 101 Then
                    suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 140 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 121 Then
                    suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 160 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 141 Then
                    suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 180 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 161 Then
                    suerte = 12
ElseIf UserList(UserIndex).Stats.UserSkills(Apuñalar) <= 200 _
   And UserList(UserIndex).Stats.UserSkills(Apuñalar) >= 181 Then
                    suerte = 9
End If

If UserList(UserIndex).Stats.UserSkills(Apuñalar) = 200 Then suerte = 7

If UCase$(UserList(UserIndex).clase) = "ASESINO" Then suerte = suerte - 4
res = RandomNumber(1, suerte)

If res = 2 Then
    If VictimUserIndex <> 0 Then
        If UserList(UserIndex).Char.Heading = UserList(VictimUserIndex).Char.Heading Then
            UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - (daño * 2)
            Call SendData(ToIndex, UserIndex, 0, "||Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & (daño * 2) & "´" & FontTypeNames.FONTTYPE_FIGHT)
            Call SendData(ToIndex, VictimUserIndex, 0, "||Te ha apuñalado " & UserList(UserIndex).Name & " por " & (daño * 2) & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Else
            UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
            Call SendData(ToIndex, UserIndex, 0, "||Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & daño & "´" & FontTypeNames.FONTTYPE_FIGHT)
            Call SendData(ToIndex, VictimUserIndex, 0, "||Te ha apuñalado " & UserList(UserIndex).Name & " por " & daño & "´" & FontTypeNames.FONTTYPE_FIGHT)
        End If
    Else
        If UserList(UserIndex).Char.Heading = Npclist(VictimNpcIndex).Char.Heading Then
            Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - (daño * 2)
            Call SendData(ToIndex, UserIndex, 0, "||Has apuñalado la criatura por " & (daño * 2) & "´" & FontTypeNames.FONTTYPE_FIGHT)
        Else
            Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - daño
            Call SendData(ToIndex, UserIndex, 0, "||Has apuñalado la criatura por " & daño & "´" & FontTypeNames.FONTTYPE_FIGHT)
        End If
        Call SubirSkill(UserIndex, Apuñalar)
    End If
    
Else
    Call SendData(ToIndex, UserIndex, 0, "||No has podido apuñalar a tu enemigo" & "´" & FontTypeNames.FONTTYPE_FIGHT)
End If
ako:

Exit Sub
fallo:
Call LogError("doapuñalar " & Err.number & " D: " & Err.Description)

End Sub
Public Sub DoDobleArma(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
On Error GoTo fallo
Dim suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(DobleArma) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= -1 Then
                    suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(DobleArma) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= 21 Then
                    suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(DobleArma) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= 41 Then
                    suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(DobleArma) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= 61 Then
                    suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(DobleArma) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= 81 Then
                    suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(DobleArma) <= 120 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= 101 Then
                    suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(DobleArma) <= 140 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= 121 Then
                    suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(DobleArma) <= 160 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= 141 Then
                    suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(DobleArma) <= 180 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= 161 Then
                    suerte = 12
ElseIf UserList(UserIndex).Stats.UserSkills(DobleArma) <= 200 _
   And UserList(UserIndex).Stats.UserSkills(DobleArma) >= 181 Then
                    suerte = 9
End If
If UserList(UserIndex).Stats.UserSkills(DobleArma) = 200 Then suerte = 7
'If UCase$(UserList(UserIndex).clase) = "ASESINO" Then suerte = suerte - 4
res = RandomNumber(1, suerte)

If res < 6 Then
    If VictimUserIndex <> 0 Then
UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - CInt(daño / 2)
Call SendData(ToIndex, UserIndex, 0, "||Golpeas con Segunda Arma a " & UserList(VictimUserIndex).Name & " por " & CInt(daño / 2) & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendData(ToIndex, VictimUserIndex, 0, "||Te ha Golpeado con su Segunda Arma " & UserList(UserIndex).Name & " por " & CInt(daño / 2) & "´" & FontTypeNames.FONTTYPE_FIGHT)
    Else
Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - CInt(daño / 2)
Call SendData(ToIndex, UserIndex, 0, "||Golpeas con segunda arma por " & CInt(daño / 2) & "´" & FontTypeNames.FONTTYPE_FIGHT)
Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(VictimNpcIndex).Name & ": " & Npclist(VictimNpcIndex).Stats.MinHP & "/" & Npclist(VictimNpcIndex).Stats.MaxHP & "´" & FontTypeNames.FONTTYPE_FIGHT)
End If
    
Call SubirSkill(UserIndex, DobleArma)
Else
Call SendData(ToIndex, UserIndex, 0, "||No has podido golpear con la Segunda Arma." & "´" & FontTypeNames.FONTTYPE_FIGHT)
End If
ako:

Exit Sub
fallo:
Call LogError("dodoblearma " & Err.number & " D: " & Err.Description)

End Sub
Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
On Error GoTo fallo
'pluto:6.8
If UserList(UserIndex).flags.Privilegios > 0 Then Exit Sub

UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0

If UserList(UserIndex).Stats.MinSta = 0 And UserList(UserIndex).flags.Angel > 0 Then
'[gau]
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).flags.Angel, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 1 & "," & 0)
UserList(UserIndex).flags.Angel = 0
UserList(UserIndex).flags.Sed = 0
UserList(UserIndex).flags.Hambre = 0
End If

If UserList(UserIndex).Stats.MinSta = 0 And UserList(UserIndex).flags.Demonio > 0 Then
'[gau]
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).flags.Demonio, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.Botas)
Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 1 & "," & 0)
UserList(UserIndex).flags.Demonio = 0
UserList(UserIndex).flags.Sed = 0
UserList(UserIndex).flags.Hambre = 0
End If
Exit Sub
fallo:
Call LogError("quitarstamina " & Err.number & " D: " & Err.Description)

End Sub

Public Sub DoTalar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim suerte As Integer
Dim res As Integer
 'pluto:2.12
 UserList(UserIndex).Counters.IdleCount = 0
 'pluto:2.11
'If MapInfo(UserList(UserIndex).pos.Map).Pk = False Then
'Call SendData(ToIndex, UserIndex, 0, "||Está Prohibido talar en Ciudad." & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
'End If

If UserList(UserIndex).clase = "Leñador" Then
    Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(Talar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= -1 Then
                    suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 21 Then
                    suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 41 Then
                    suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 61 Then
                    suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 81 Then
                    suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 120 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 101 Then
                    suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 140 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 121 Then
                    suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 160 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 141 Then
                    suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 180 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 161 Then
                    suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(Talar) <= 200 _
   And UserList(UserIndex).Stats.UserSkills(Talar) >= 181 Then
                    suerte = 10
End If

If UserList(UserIndex).Stats.UserSkills(Talar) = 200 Then suerte = 7

res = RandomNumber(1, suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As obj
    
    If UserList(UserIndex).clase = "Leñador" Then
        MiObj.Amount = RandomNumber(1, CInt(UserList(UserIndex).Stats.ELV * 2))
    Else
        MiObj.Amount = 1
    End If
        
    If MiObj.Amount < 1 Then MiObj.Amount = 1
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "G1")
  'pluto.2.4.1
UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + (CInt((UserList(UserIndex).Stats.ELV / 10) + 1) * MiObj.Amount)
Call CheckUserLevel(UserIndex)
  
Else
    'Call SendData(ToIndex, UserIndex, 0, "G2")
End If

Call SubirSkill(UserIndex, Talar)

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
On Error GoTo fallo
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Exit Sub

If UserList(UserIndex).flags.Privilegios = 0 Then
    UserList(UserIndex).Reputacion.BurguesRep = 0
    UserList(UserIndex).Reputacion.NobleRep = 0
    UserList(UserIndex).Reputacion.PlebeRep = 0
    'pluto:2.4
'UserCiu = UserCiu - 1
'UserCrimi = UserCrimi + 1
    
    Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
    'If UserList(UserIndex).Faccion.ArmadaReal = 2 Then Call ExpulsarFaccionlegion(UserIndex)

End If
Exit Sub
fallo:
Call LogError("volvercriminal " & Err.number & " D: " & Err.Description)

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
On Error GoTo fallo
'pluto:hoy
If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Exit Sub


'pluto:2.4
'UserCiu = UserCiu + 1
'UserCrimi = UserCrimi - 1
If UserList(UserIndex).flags.Privilegios = 0 Then
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.AsesinoRep = 0


Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlASALTO, MAXREP)
End If

Exit Sub
fallo:
Call LogError("volverciudadano " & Err.number & " D: " & Err.Description)

End Sub


Public Sub DoPlayInstrumento(ByVal UserIndex As Integer)

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim suerte As Integer
Dim res As Integer
Dim metal As Integer
 'pluto:2.12
 UserList(UserIndex).Counters.IdleCount = 0
 'pluto:2.11
'If MapInfo(UserList(UserIndex).pos.Map).Pk = False Then
'Call SendData(ToIndex, UserIndex, 0, "||Está Prohibido Minar en Ciudad." & FONTTYPENAMES.FONTTYPE_INFO)
'Exit Sub
'End If


If UserList(UserIndex).clase = "Minero" Then
    Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(Mineria) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= -1 Then
                    suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 21 Then
                    suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 41 Then
                    suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 61 Then
                    suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 81 Then
                    suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 120 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 101 Then
                    suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 140 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 121 Then
                    suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 160 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 141 Then
                    suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 180 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 161 Then
                    suerte = 12
ElseIf UserList(UserIndex).Stats.UserSkills(Mineria) <= 200 _
   And UserList(UserIndex).Stats.UserSkills(Mineria) >= 181 Then
                    suerte = 10
End If
If UserList(UserIndex).Stats.UserSkills(Mineria) = 200 Then suerte = 7

res = RandomNumber(1, suerte)
Dim res2 As Integer
    If res <= 5 Then
    Dim MiObj As obj
    Dim nPos As WorldPos
    
    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
    'objeto diamante
    res2 = RandomNumber(1, 100)
        If UserList(UserIndex).clase = "Minero" Then
        'nati: Si el usuario NO ESTÁ en minas fortaleza, minara por su nivel * 2.
        If Not UserList(UserIndex).Pos.Map = 186 Then
        MiObj.Amount = RandomNumber(1, CInt(UserList(UserIndex).Stats.ELV * 2))
        Else
        MiObj.Amount = RandomNumber(1, CInt(UserList(UserIndex).Stats.ELV))
        End If
       'nati:            FIN
    Else
        MiObj.Amount = 1
    End If
    
    If MiObj.Amount < 1 Then MiObj.Amount = 1
    
    'pluto:6.0A
    If res2 = 25 Then
    MiObj.ObjIndex = 695
    MiObj.Amount = 1
    ElseIf res2 = 26 Then
    MiObj.ObjIndex = 1170
    MiObj.Amount = 1
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    
    Call SendData(ToIndex, UserIndex, 0, "G5")
'pluto.2.4.1
UserList(UserIndex).Stats.exp = UserList(UserIndex).Stats.exp + (CInt((UserList(UserIndex).Stats.ELV / 10) + 1) * MiObj.Amount)
Call CheckUserLevel(UserIndex)
    
Else
    Call SendData(ToIndex, UserIndex, 0, "G6")
End If

Call SubirSkill(UserIndex, Mineria)


Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub



Public Sub DoMeditar(ByVal UserIndex As Integer)
On Error GoTo errhandler
UserList(UserIndex).Counters.IdleCount = 0

Dim suerte As Integer
Dim res As Integer
Dim Cant As Integer

If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
    Call SendData(ToIndex, UserIndex, 0, "G7")
    Call SendData2(ToIndex, UserIndex, 0, 54)
    Call SendData2(ToIndex, UserIndex, 0, 15, UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    'pluto:bug meditar
    Call SendData2(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 22, UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(Meditar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= -1 Then
                    suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 21 Then
                    suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 41 Then
                    suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 61 Then
                    suerte = 16
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 81 Then
                    suerte = 14
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 120 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 101 Then
                    suerte = 12
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 140 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 121 Then
                    suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 160 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 141 Then
                    suerte = 8
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 180 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 161 Then
                    suerte = 6
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) <= 200 _
   And UserList(UserIndex).Stats.UserSkills(Meditar) >= 181 Then
                    suerte = 4
ElseIf UserList(UserIndex).Stats.UserSkills(Meditar) = 200 Then
                    suerte = 3
End If

res = RandomNumber(1, suerte)

If res = 1 Then
    Cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, 3)
    Call AddtoVar(UserList(UserIndex).Stats.MinMAN, Cant, UserList(UserIndex).Stats.MaxMAN)
    Call SendData(ToIndex, UserIndex, 0, "V5" & Cant)
    Call SendUserStatsMana(UserIndex)
    Call SubirSkill(UserIndex, Meditar)
End If
'pluto:2.5.0
Exit Sub

errhandler:
    Call LogError("Error en Sub DoMeditar")

End Sub
