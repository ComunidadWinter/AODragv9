Attribute VB_Name = "modMonturas"
Option Explicit

Public Const MAX_MONTURAS As Byte = 3 'Numero maximo de monturas por personaje
Public Const ELU_INICIAL As Integer = 30 'Experiencia requerida para pasar del nivel 1 al 2
Public Const MONTURA_MAX_LEVEL As Byte = 45 'Nivel maximo de la montura

Type tMontus
    id As Integer
    tipo As Byte
    nombre As String
    MonturaLevel As Byte
    Ataque As Byte
    Defensa As Byte
    AtMagia As Byte
    DefMagia As Byte
    Evasion As Byte
    Skills As Byte
    ELU As Integer
    Exp As Integer
    Speed As Byte
End Type

Public Sub DoEquita(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal SlotMontura As Byte)
'*************************************************
'Author: Lorwik
'Ultima modificacion: 16/12/2018
'Descripción: Cambia el estado del usuario a equitando y le aplica las propiedades.
'Changelog:
'- Añadido modificador de velocidad
'*************************************************
Dim ObjMontura As ObjData

    With UserList(UserIndex)
        ObjMontura = ObjData(.Invent.Object(Slot).ObjIndex) 'Lo pongo asi por que tira error :(
        .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
        .Invent.MonturaSlot = Slot
        
        If .flags.QueMontura = 0 Then
               .Char.Head = 0
               
                If .flags.Muerto = 0 Then
                   .Char.body = ObjMontura.Ropaje
                Else
                   .Char.body = iCuerpoMuerto
                   .Char.Head = iCabezaMuerto
                End If
               
                .Char.Head = UserList(UserIndex).OrigChar.Head
                .Char.ShieldAnim = .Char.ShieldAnim
                .Char.WeaponAnim = .Char.WeaponAnim
                .Char.CascoAnim = .Char.CascoAnim
               
               '.flags.Montando = 1
               .flags.QueMontura = SlotMontura 'Con esto anotamos que montura estoy usando de mi lista.
               .flags.Speed = .flags.Montura(SlotMontura).Speed
        Else
                '.flags.Montando = 0
                .flags.QueMontura = 0
                .flags.Speed = 0
            If .flags.Muerto = 0 Then
                  .Char.Head = UserList(UserIndex).OrigChar.Head
                  
                   If .Invent.ArmourEqpObjIndex > 0 Then
                      .Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
                   Else
                       Call DarCuerpoDesnudo(UserIndex)
                   End If
                   
                    If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
                    If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
                    If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
                 
            Else
                 .Char.body = iCuerpoMuerto
                 .Char.Head = iCabezaMuerto
                 .Char.ShieldAnim = NingunEscudo
                 .Char.WeaponAnim = NingunArma
                 .Char.CascoAnim = NingunCasco
            End If
        End If
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteMontateToggle(UserIndex)
        Call WriteChangeSpeed(UserIndex, .Char.CharIndex, .flags.Speed)
    End With
End Sub

Public Sub CheckMonturaLevel(ByVal UserIndex As Integer)
Dim MonturaSlot As Byte
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        'Primero ¿El usuario tiene montura?
        If .Stats.NUMMONTURAS = 0 Then Exit Sub
        
        'Ok, ¿Esta montado?
        If .flags.QueMontura = 0 Then Exit Sub
        
        MonturaSlot = .flags.QueMontura

        '<Edurne> Vamos a arreglar esto un poco xD
        With .flags.Montura(MonturaSlot)
            '¿Ya es nivel maximo?
            If MONTURA_MAX_LEVEL < .MonturaLevel Then Exit Sub
            
            '¿Llego a la experiencia requerida para pasar de nivel?
            Do While .Exp >= .ELU
            
                'Notificamos
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Call WriteConsoleMsg(UserIndex, "¡Tu montura subió de nivel!", FontTypeNames.FONTTYPE_INFO)
                
                'Deducimos la exp:
                .Exp = .Exp - .ELU
                
                'Subimos de nivel:
                .MonturaLevel = .MonturaLevel + 1
                
                'Nuevos Skills para asignar
                .Skills = .Skills + 1
                
                'Nuevo ELU
                Select Case .MonturaLevel
                    Case 1 To 10
                        .ELU = .ELU * 1.4
                    Case 11 To 20
                        .ELU = .ELU * 1.15
                    Case 21 To 30
                        .ELU = .ELU * 1.2
                    Case 31 To 40
                        .ELU = .ELU * 1.22
                    Case Else
                        .ELU = .ELU * 1.3
                End Select
                If .MonturaLevel = MONTURA_MAX_LEVEL Then .Exp = 0: .ELU = 0

                'Le decimos al usuario que su montura tiene nuevos puntos de skills para asignar:
                Call WriteConsoleMsg(UserIndex, "¡Tu montura tiene puntos de Skills libres para asignar!", FontTypeNames.FONTTYPE_INFO)
            Loop
        End With
        '</Edurne>
    End With
    Exit Sub
errHandler:
    LogError ("Error en la subrutina CheckMonturaLevel")
End Sub

Public Sub AsignacionSkillMontura(ByVal UserIndex As Integer, ByVal SlotMontura As Byte, ByVal SlotSkill As Integer)
    With UserList(UserIndex)
        '¿Tiene monturA?
        If .Stats.NUMMONTURAS = 0 Then Exit Sub
        
        '<Edurne>
        If (SlotSkill < 0 Or SlotSkill > 4) Or _
           (SlotMontura < 1 Or SlotMontura > 5) Then Exit Sub
        
        With .flags.Montura(SlotMontura)
            'Comprobamos si tiene skills libres
            If Not .Skills > 0 Then Exit Sub
            
            'No me gusta tal y como esta, quizas lo cambio
            Select Case SlotSkill
                Case 0
                    If .Ataque = 10 Then Exit Sub
                    .Ataque = .Ataque + 1
                Case 1
                    If .Defensa = 10 Then Exit Sub
                    .Defensa = .Defensa + 1
                Case 2
                    If .AtMagia = 10 Then Exit Sub
                    .AtMagia = .AtMagia + 1
                Case 3
                    If .DefMagia = 10 Then Exit Sub
                    .DefMagia = .DefMagia + 1
                Case 4
                    If .Evasion = 10 Then Exit Sub
                    .Evasion = .Evasion + 1
            End Select
            
            'Restamo Skills al contado
            .Skills = .Skills - 1
        End With
        '</Edurne>
        
        'Actualizamos la lista del cliente
        Call WriteInfoMontura(UserIndex, SlotMontura)
    End With
End Sub

