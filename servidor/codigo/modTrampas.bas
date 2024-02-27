Attribute VB_Name = "modTrampas"
Option Explicit

Public Const NumSaqueador As Integer = 517 'Numero del saqueador segun NPC.dat
Public SaqueadorIndex As Integer 'Numero del NPCindex del saqueador cuando esta en el mapa


Public Sub Gusano(ByVal UserIndex As Integer)
'Sistema del gusano creado por no se quien de AODrag y adaptado por Lorwik :P
On Error GoTo fallo
    Dim daño As Integer
    Dim lado As Integer
    'Calculamos el daño que le vamos a hacer
    daño = RandomNumber(5, 20)
    'Supongo que sera del lado en el que va a venir el bicho (?)
    'lado = RandomNumber(35, 36)
    'Mandamos el Wav y el FX
    'Call SendData2(ToPCArea, UserIndex, UserList(UserIndex).Pos.map, 22, UserList(UserIndex).Char.CharIndex & "," & lado & "," & 1)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 16, 0))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(218, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    'Le restamos el daño
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
    Call WriteUpdateHP(UserIndex)
    'Notificamos al usuario que el gusanito le ataco
    Call WriteConsoleMsg(UserIndex, "¡¡Un Gusano te causa " & daño & " de daño!!", FontTypeNames.FONTTYPE_FIGHT)
    If UserList(UserIndex).Stats.MinHP <= 0 Then Call UserDie(UserIndex)
    Exit Sub
fallo:
    Call LogError("GUSANO" & Err.Number & " D: " & Err.Description)

End Sub

Public Sub Trampa(ByVal UserIndex As Integer, Tipotrampa As Integer)
On Error GoTo fallo
    Dim daño As Integer
    
    'Calculamos el daño de la trampa
    daño = RandomNumber(5, 20)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Tipotrampa, 0))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(218, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
    Call WriteConsoleMsg(UserIndex, "¡¡Una trampa te causa " & daño & " de daño!!", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteUpdateHP(UserIndex)
    If UserList(UserIndex).Stats.MinHP <= 0 Then Call UserDie(UserIndex)
    Exit Sub
fallo:
    Call LogError("TRAMPA " & Err.Number & " D: " & Err.Description)

End Sub

Sub CasaEncantada(ByVal UserIndex As Integer)
'Creado por Pluto, adaptado y mejorado por Lorwik
'pluto:2.17
Dim X As Byte
Dim Y As Byte
Dim Map As Integer
Dim DadosCasa As Byte

    With UserList(UserIndex)
        Map = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
        
        'pluto:rayos puerta
        If (X = CasaRayoX1 Or X = CasaRayoX2) And Y = CasaRayoY And .flags.Muerto = 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(113, X, Y))
        
        If .Counters.Morph > 0 Then Exit Sub
        
        'pluto:sala sangre casa
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 11 Then
            Call WriteConsoleMsg(UserIndex, "¡¡La habitación de sangre te ha matado!!", FontTypeNames.FONTTYPE_FIGHT)
            Call UserDie(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(102, .Pos.X, .Pos.Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FxCASA, 0))
            Exit Sub
        End If
        
        'Lorwik:Espitirus
        DadosCasa = RandomNumber(1, 300)
        Select Case DadosCasa
        
            Case 1
                If .Stats.GLD >= 3000 Then
                    Call WriteConsoleMsg(UserIndex, "Los Espiritus de la Casa te hacen perder Oro.", FontTypeNames.FONTTYPE_FIGHT)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(102, .Pos.X, .Pos.Y))
                    Call TirarOro(3000, UserIndex)
                    Call WriteUpdateUserStats(UserIndex)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FxCASA, 0))
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 3, 0))
                    Exit Sub
                End If
            
            Case 30
                If .flags.Morph = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Los Espiritus de la Casa te transforman en Cerdo.", FontTypeNames.FONTTYPE_FIGHT)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(102, .Pos.X, .Pos.Y))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FxCASA, 0))
                    .flags.Morph = .Char.body
                    .Counters.Morph = IntervaloMorphPJ
                    Call ChangeUserChar(UserIndex, 6, 0, UserList(UserIndex).Char.heading, 2, 2, 2)
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 3, 0))
                    Exit Sub
                End If
            
            Case 49
                Call WriteConsoleMsg(UserIndex, "Los Espiritus de la Casa te hacen perder el inventario.", FontTypeNames.FONTTYPE_FIGHT)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(102, .Pos.X, .Pos.Y))
                Call TirarTodosLosItems(UserIndex)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FxCASA, 0))
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 3, 0))
                Exit Sub
            
            Case 53
                Call WriteConsoleMsg(UserIndex, "Los Espiritus de la Casa te teleportan fuera de ella.", FontTypeNames.FONTTYPE_FIGHT)
                Call PrepareMessagePlayWave(102, 0, 0)
                Call WarpUserChar(UserIndex, 5, 38, 36, True)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FxCASA, 0))
                Exit Sub
            
            Case 97
                Call WriteConsoleMsg(UserIndex, "Los Espiritus de la casa te han Paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Call PrepareMessagePlayWave(102, 0, 0)
                .flags.Paralizado = 1
                .Counters.Paralisis = IntervaloParalizado
                Call WriteParalizeOK(UserIndex)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(247, .Pos.X, .Pos.Y))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 8, 0))
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 3, 0))
                Exit Sub
        End Select
    End With
End Sub

Public Sub Saqueador(ByVal Spawn As Boolean)
'*****************************************************************************************************************
'Autor: Lorwik
'Fecha: 05/05/2016
'Descripción: Hace spawn a un NPC llamado saqueador que puede aparecer en 2 mapas en una posición aleatoria.
'O puede matar a un saqueador existente.
'*****************************************************************************************************************

    If Spawn = True Then
        Dim Pos As WorldPos
        
        Pos.Map = RandomNumber(31, 32) 'Puede aparecer en uno de estos 2 mapas.
        Pos.X = RandomNumber(15, 75) 'Puede aparecer en cualquiera de esta X
        Pos.Y = RandomNumber(17, 81) 'Puede aparecer en cualquiera de esta Y
        
        Call SpawnNpc(NumSaqueador, Pos, True, False)
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Un saqueador aparecio en el interior de la piramide.", FontTypeNames.FONTTYPE_DIOS))
    Else
        Call QuitarNPC(SaqueadorIndex)
    End If
End Sub

