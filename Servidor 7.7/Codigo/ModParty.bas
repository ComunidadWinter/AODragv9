Attribute VB_Name = "ModParty"
Sub creaParty(ByVal UserIndex As Integer, privada As Byte)
    
    On Error GoTo errhandler
    Dim n As Integer
    Dim encontrado As Boolean
If UserList(UserIndex).flags.party = False Then
    If numPartys >= MAXPARTYS Then
        Call SendData(ToIndex, UserIndex, 0, "DD9A")
        'pluto:6.5
        Call LogParty(UserList(UserIndex).Name & ": Intenta crear con Max")
        'Call SendData(ToIndex, UserIndex, 0, "||No puedes crear partys en este momento." & FONTTYPENAMES.FONTTYPE_INFO)
    Else
        encontrado = False
        n = 0
        Do While (n <= MAXPARTYS And encontrado <> True)
            n = n + 1
            If partylist(n).lider = 0 Then
                encontrado = True
            End If
        Loop
        If encontrado = True Then
            UserList(UserIndex).flags.partyNum = n
            UserList(UserIndex).flags.party = True
            numPartys = numPartys + 1
            partylist(n).lider = UserIndex
            partylist(n).expAc = 0
            partylist(n).reparto = 1
            partylist(n).privada = privada
            partylist(n).numMiembros = 1
            partylist(n).miembros(1).ID = UserIndex
            partylist(n).miembros(1).privi = 100
            Call SendData(ToIndex, UserIndex, 0, "DD10")
        'pluto:6.5
        Call LogParty(UserList(UserIndex).Name & ": Crea Nº: " & n & " Numpartys: " & numPartys)
            'Call SendData(ToIndex, UserIndex, 0, "||Has creado una party!" & FONTTYPENAMES.FONTTYPE_INFO)
        End If
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "DD11")
            'pluto:6.5
        Call LogParty(UserList(UserIndex).Name & ": Intenta crear perteneciendo a otra.")
    'Call SendData(ToIndex, UserIndex, 0, "||Ya perteneces a una party!" & FONTTYPENAMES.FONTTYPE_INFO)
End If
Exit Sub
errhandler:
    Call LogError("Error en CreaPArty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " PRIV:" & privada & " N: " & Err.number & " D: " & Err.Description)
'    Call LogError("Error en creaParty")
    
End Sub
Sub quitParty(ByVal UserIndex As Integer)

    On Error GoTo errhandler


Dim lpp As Byte
Dim miembro As Integer
Dim partyid As Integer
If UserList(UserIndex).flags.party = False Then
    Call SendData(ToIndex, UserIndex, 0, "DD8A")
        'pluto:6.5
        Call LogParty(UserList(UserIndex).Name & ": Intenta Cerrar party sin estar en party")
    'Call SendData(ToIndex, UserIndex, 0, "||No estas en ninguna party" & FONTTYPENAMES.FONTTYPE_INFO)
    Exit Sub
End If
If UserIndex = partylist(UserList(UserIndex).flags.partyNum).lider And UserList(UserIndex).flags.party = True Then
        partyid = UserList(UserIndex).flags.partyNum
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": Finaliza party")
        lpp = MAXMIEMBROS
        Do While lpp > 0

            If partylist(partyid).miembros(lpp).ID <> 0 Then
                miembro = partylist(partyid).miembros(lpp).ID
                Call quitUserParty(miembro)
                partylist(partyid).miembros(lpp).ID = 0
                'UserList(miembro).flags.party = False
                'Call SendData(ToIndex, miembro, 0, "||Party finalizada!" & FONTTYPENAMES.FONTTYPE_INFO)
            End If
            lpp = lpp - 1
        Loop
        Call SendData(toParty, UserIndex, 0, "DD12")
        'Call SendData(toParty, miembro, 0, "||Party finalizada!" & FONTTYPENAMES.FONTTYPE_INFO)
        UserList(UserIndex).flags.party = False
        partylist(partyid).lider = 0
        partylist(partyid).expAc = 0
        partylist(partyid).numMiembros = 0
        numPartys = numPartys - 1
 'pluto:6.5
Call LogParty("NumeroPartys: " & numPartys)
Else
    Call SendData(ToIndex, UserIndex, 0, "DD13")
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": Intenta Finalizar party")
    'Call SendData(ToIndex, UserIndex, 0, "||Debes ser el lider de la party para poder finalizarla." & FONTTYPENAMES.FONTTYPE_INFO)
End If
    
Exit Sub
errhandler:
    'Call LogError("Error en quitaPArty")
    Call LogError("Error en quitaParty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " PID:" & partyid & " N: " & Err.number & " D: " & Err.Description)
End Sub
Sub addUserParty(ByVal UserIndex As Integer, PartyIndex As Integer)

    On Error GoTo errhandler
    Dim n As Integer
    
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": añadir a party nº " & PartyIndex)
    
    
    If partylist(PartyIndex).numMiembros >= MAXMIEMBROS Then
        Call SendData(ToIndex, partylist(PartyIndex).lider, 0, "||Party llena" & "´" & FontTypeNames.FONTTYPE_info)
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": no se añade por party llena.")
        Exit Sub
        Else


    End If
   'pluto:6.7-----------------
   For n = 1 To MaxUsers
   If UserList(n).flags.invitado = UserList(UserIndex).Name Then UserList(n).flags.invitado = ""
   Next
    '------------------------
    n = 0
        Do While (n < MAXMIEMBROS)
            n = n + 1
            If partylist(PartyIndex).miembros(n).ID = 0 And UserList(UserIndex).flags.party = False Then
                partylist(PartyIndex).miembros(n).ID = UserIndex
                partylist(PartyIndex).miembros(n).privi = 0
                partylist(PartyIndex).numMiembros = partylist(PartyIndex).numMiembros + 1
                UserList(UserIndex).flags.partyNum = PartyIndex
                UserList(UserIndex).flags.party = True
                UserList(UserIndex).flags.invitado = ""
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": añadido ha Party nº " & PartyIndex & " en pos " & n)
Call LogParty("Party nº " & PartyIndex & " Miembros: " & partylist(PartyIndex).numMiembros)
                
                Call sendMiembrosParty(UserIndex)
'pluto:6.3--------
Call SendData(toParty, UserIndex, 0, "DD14" & UserList(UserIndex).Name)
'Dim fp As Byte
'Dim mie As String
'For fp = 1 To MAXMIEMBROS
'mie = mie + UserList(partylist(PartyIndex).miembros(fp).ID).Char.CharIndex & ","
'Next
'Call SendData(ToIndex, UserIndex, 0, "O6" & PartyIndex & "," & mie)
'-----------------


                'Call SendData(toParty, partylist(partyindex).lider, 0, "||" & UserList(UserIndex).Name & " se ha incorporado a la party." & FONTTYPENAMES.FONTTYPE_INFO)
                'añadir user ponemos reparto proporcional
                partylist(PartyIndex).reparto = 1
                
                If partylist(PartyIndex).reparto = 1 Then
                    Call BalanceaPrivisLVL(PartyIndex)
                ElseIf partylist(PartyIndex).reparto = 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "DD15")
                    'Call SendData(ToIndex, partylist(partyindex).lider, 0, "||Modifica los privilegios para el nuevo usuario" & FONTTYPENAMES.FONTTYPE_INFO)
                ElseIf partylist(PartyIndex).reparto = 3 Then
                    Call BalanceaPrivisMiembros(PartyIndex)
                End If
                Call sendPriviParty(UserIndex)
            End If
            If partylist(PartyIndex).Solicitudes(n) = UserIndex Then
                partylist(PartyIndex).Solicitudes(n) = 0
                partylist(PartyIndex).numSolicitudes = partylist(PartyIndex).numSolicitudes - 1
          'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": Borrado de solicitudes")
Call LogParty("Nº Party: " & PartyIndex & " Solicitudes: " & partylist(PartyIndex).numSolicitudes)
            
            End If
                    
        Loop
    
Exit Sub
errhandler:
    Call LogError("Error en addUserParty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " PID:" & PartyIndex & " N: " & Err.number & " D: " & Err.Description)
    'Call LogError("Error en addUserPArty")
End Sub
Sub addSoliParty(ByVal UserIndex As Integer, PartyIndex As Integer)

    On Error GoTo errhandler
    Dim n As Integer
    Dim encontrado As Boolean
          'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": añadir solicitud a la party nº " & PartyIndex)

If UserList(UserIndex).flags.party = False Then
    encontrado = False
    For n = 1 To MAXMIEMBROS
        If partylist(PartyIndex).Solicitudes(n) = UserIndex Then
            Call SendData(ToIndex, UserIndex, 0, "DD26")
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": no añadida pq ya envío antes.")

            Exit Sub
        End If
    Next
    n = 0
    If partylist(PartyIndex).numSolicitudes >= MAXMIEMBROS Then
        Call SendData(ToIndex, UserIndex, 0, "DD16")
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": no añadida por cola llena.")
        'Call SendData(ToIndex, UserIndex, 0, "||Cola de solicitudes llena, no puedes unirte en este momento." & FONTTYPENAMES.FONTTYPE_INFO)
    Else
        Do While (n < MAXMIEMBROS And encontrado <> True)
            n = n + 1
            If partylist(PartyIndex).Solicitudes(n) = 0 Then
                encontrado = True
            End If
            
        Loop
        If encontrado = True Then
           ' UserList(UserIndex).flags.partyNum = PartyIndex
            partylist(PartyIndex).Solicitudes(n) = UserIndex
            partylist(PartyIndex).numSolicitudes = partylist(PartyIndex).numSolicitudes + 1
            Call SendData(ToIndex, partylist(PartyIndex).lider, 0, "DD17" & UserList(UserIndex).Name)
            'Call SendData(ToIndex, partylist(partyindex).lider, 0, "||" & UserList(UserIndex).Name & " solicita entrar en la party ." & FONTTYPENAMES.FONTTYPE_INFO)
            Call SendData(ToIndex, UserIndex, 0, "DD18" & UserList(partylist(PartyIndex).lider).Name)
            'Call SendData(ToIndex, UserIndex, 0, "||Solicitud enviada a la party de " + UserList(partylist(partyindex).lider).Name + " ." & FONTTYPENAMES.FONTTYPE_INFO)
        'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": solicitud añadida a la party nº " & PartyIndex & " en pos " & n)
Call LogParty("Total solicitudes: " & partylist(PartyIndex).numSolicitudes)
        End If
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "DD11")
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": no añadida pq ya pertenece a una party.")
    'Call SendData(ToIndex, UserIndex, 0, "||Ya perteneces a una party." & FONTTYPENAMES.FONTTYPE_INFO)
End If
    
Exit Sub
errhandler:
    Call LogError("Error en addSoliParty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " PID:" & PartyIndex & " N: " & Err.number & " D: " & Err.Description)
    'Call LogError("Error en addUserPArty")
End Sub
Sub quitSoliParty(ByVal UserIndex As Integer, PartyIndex As Integer)
    
    On Error GoTo errhandler
    Dim n As Integer
    Dim encontrado As Boolean
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": quitar solicitud a la party nº " & PartyIndex)
  
    
    encontrado = False
    n = 0
    Do While (n < MAXMIEMBROS And encontrado <> True)
        n = n + 1
        If partylist(PartyIndex).Solicitudes(n) = UserIndex Then
            encontrado = True
        End If
    Loop
    If encontrado = True Then
        partylist(PartyIndex).Solicitudes(n) = 0
        partylist(PartyIndex).numSolicitudes = partylist(PartyIndex).numSolicitudes - 1
        UserList(UserIndex).flags.partyNum = 0
        UserList(UserIndex).flags.party = False
          'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": solicitud quitada en pos " & n)
Call LogParty("Party: " & PartyIndex & " Solicitudes: " & partylist(PartyIndex).numSolicitudes)

        
    Else
Call LogParty(UserList(UserIndex).Name & ": error quitar user no encontrado en party: " & PartyIndex)

        GoTo errhandler
    End If

    Exit Sub
errhandler:
    Call LogError("Error en quitSoliParty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " PID:" & PartyIndex & " N: " & Err.number & " D: " & Err.Description)
    'Call LogError("Error en quitSoliPArty")
    
End Sub
Sub quitUserParty(ByVal UserIndex As Integer)
    On Error GoTo errhandler
    Dim n As Integer
    Dim encontrado As Boolean
    Dim PartyIndex As Integer
    
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": vamos a quitarlo de party")
  
If UserIndex = 0 Then Exit Sub
If UserList(UserIndex).flags.party = True Then
If esLider(UserIndex) = True And partylist(UserList(UserIndex).flags.partyNum).numMiembros > 1 Then
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": no se quita pq es lider")

Exit Sub
End If
    PartyIndex = UserList(UserIndex).flags.partyNum
    encontrado = False
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": está en la party " & PartyIndex)

    'n = 1
    'Do While (n < MAXMIEMBROS)
    For n = 1 To MAXMIEMBROS
        If partylist(PartyIndex).miembros(n).ID = UserIndex Then
        'Debug.Print UserList(UserIndex).Name
            partylist(PartyIndex).miembros(n).ID = 0
            partylist(PartyIndex).miembros(n).privi = 0
            'partylist(PartyIndex).numMiembros = partylist(PartyIndex).numMiembros - 1
                       
            Call SendData(ToIndex, UserIndex, 0, "DD19" & partylist(UserList(UserIndex).flags.partyNum).expAc)
          Call SendData(ToIndex, UserIndex, 0, "DD20")
        'Call SendData(ToIndex, UserIndex, 0, "||Has abandonado la party!" & FONTTYPENAMES.FONTTYPE_INFO)
            Call SendData(ToIndex, UserIndex, 0, "W10,")
                   'Call SendData(ToIndex, UserIndex, 0, "||Has ganado un total de " & partylist(UserList(UserIndex).flags.partyNum).expAc & " puntos de experiencia" & FONTTYPENAMES.FONTTYPE_INFO)
            'Call sendMiembrosParty(partylist(UserList(UserIndex).flags.partyNum).lider)
        'pluto:6.3---------
        Call SendData(toParty, UserIndex, 0, "O5" & UserList(UserIndex).Char.CharIndex)
   
      
      '-----------------
 'pluto:6.3 ponemos esto detras de enviar miembrosparty
 partylist(PartyIndex).numMiembros = partylist(PartyIndex).numMiembros - 1
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & ": quitado en pos " & n)
Call LogParty("Miembros Party: " & partylist(PartyIndex).numMiembros)


            UserList(UserIndex).flags.partyNum = 0
            UserList(UserIndex).flags.party = False
            UserList(UserIndex).flags.invitado = ""
            'pluto:6.7 añade reparto 2
            If partylist(PartyIndex).reparto = 1 Or partylist(PartyIndex).reparto = 2 Then
                Call BalanceaPrivisLVL(PartyIndex)
            ElseIf partylist(PartyIndex).reparto = 3 Then
                Call BalanceaPrivisMiembros(PartyIndex)
            End If
        'pluto:6.7-----------
        Call sendMiembrosParty(partylist(PartyIndex).lider)
        Call sendPriviParty(partylist(PartyIndex).lider)
        '----------------------
        
        End If
        'n = n + 1
    Next
Else
    Call SendData(ToIndex, UserIndex, 0, "DD8A")
'pluto:6.5
Call LogParty(UserList(UserIndex).Name & " no está en party")
    'Call SendData(ToIndex, UserIndex, 0, "||No estas en ninguna party" & FONTTYPENAMES.FONTTYPE_INFO)
End If
    Exit Sub
errhandler:
    Call LogError("Error en quitUserParty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " PID:" & PartyIndex & "n: " & n & " N: " & Err.number & " D: " & Err.Description)
    'Call LogError("Error en quitUserPArty")
End Sub
Sub InvitaParty(indexAnfitrion As Integer, indexInvitado As Integer)
    On Error GoTo errhandler
'pluto:6.5
Call LogParty(UserList(indexAnfitrion).Name & " invita a " & UserList(indexInvitado).Name)
        
        If UserList(indexInvitado).flags.party = True Then
            Call SendData(ToIndex, indexAnfitrion, 0, "DD21" & UserList(indexInvitado).Name)
            'Call SendData(ToIndex, indexAnfitrion, 0, "||No puedes invitar a " & UserList(indexInvitado).Name & ", ya esta en una party." & FONTTYPENAMES.FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, indexAnfitrion, 0, "DD22" & UserList(indexInvitado).Name)
            'Call SendData(ToIndex, indexAnfitrion, 0, "||Has invitado a " & UserList(indexInvitado).Name & " a la party." & FONTTYPENAMES.FONTTYPE_INFO)
            Call SendData(ToIndex, indexInvitado, 0, "DD23" & UserList(indexAnfitrion).Name)
            'Call SendData(ToIndex, indexInvitado, 0, "||" & UserList(indexAnfitrion).Name & " te ha invitado a crear una party. Escribe /unirme para unirte" & FONTTYPENAMES.FONTTYPE_INFO)
            UserList(indexInvitado).flags.invitado = UserList(indexAnfitrion).Name
            'pluto:6.7
            UserList(indexAnfitrion).flags.invitado = ""
        End If
        Exit Sub
errhandler:
    Call LogError("Error en InvitaParty Anfitrion:" & UserList(indexAnfitrion).Name & " AnfiID:" & indexAnfitrion & " Invitado:" & UserList(indexInvitado).Name & " InviID:" & indexInvitado & " N: " & Err.number & " D: " & Err.Description)
End Sub
Function totalexpParty(ByVal PartyIndex As Integer) As Long
    On Error GoTo errhandler
    Dim n As Integer
    Dim total As Double
    total = 0
    For n = 1 To MAXMIEMBROS
        If partylist(PartyIndex).miembros(n).ID <> 0 Then
            total = total + UserList(partylist(PartyIndex).miembros(n).ID).Stats.ELV
        End If
    Next
    totalexpParty = total
    Exit Function
errhandler:
    Call LogError("Error en totalexpParty PID " & PartyIndex & " N: " & Err.number & " D: " & Err.Description)
End Function
Sub PartyReparteExp(NpcIndex As npc, UserIndex As Integer)
On Error GoTo errhandler
    Dim n As Byte
    Dim expIndi As Long
    Dim b As Long
    Dim aa As Integer
    Dim oo As Integer
    
        
        partylist(UserList(UserIndex).flags.partyNum).expAc = partylist(UserList(UserIndex).flags.partyNum).expAc + NpcIndex.GiveEXP
        For n = 1 To MAXMIEMBROS
        oo = partylist(UserList(UserIndex).flags.partyNum).miembros(n).ID
        
        If oo <> 0 Then
            If UserList(oo).Pos.Map = NpcIndex.Pos.Map Then
                If UserList(oo).flags.Muerto = 0 Then
                    If UserList(oo).Bebe = 0 Then
                       
                        b = partylist(UserList(UserIndex).flags.partyNum).miembros(n).privi
                        expIndi = (NpcIndex.GiveEXP / 100) * b
                        
                            If ServerPrimario = 1 Then
                                    If UserList(oo).Remort > 0 Then
                                    expIndi = expIndi * 1
                                    Else
                                    expIndi = expIndi * 1
                                    End If
                            Else 'secundario Pluto:6.5
                                    If UserList(UserIndex).Remort > 0 Then
                                    expIndi = expIndi * 1
                                    Else
                                    expIndi = expIndi * 1
                                    End If
                            End If 'primario =1
                       
                        If UserList(oo).flags.Montura > 0 Then
                        
                        expIndi = Int(expIndi / 2)  ' ORIGINAL: expIndi / 2
                        
                        
                        aa = Int(expIndi / 1000) ' ORIGINAL: Int(expIndi / 1000)
                        Else
                        aa = 0
                        End If
    
                        
      'El user tiene montura (hay que repartir exp con ella)
        If UserList(oo).flags.Montura > 0 And UserList(oo).flags.ClaseMontura > 0 Then
'añade topelevel
            If PMascotas(UserList(oo).flags.ClaseMontura).TopeLevel > UserList(oo).Montura.Nivel(UserList(oo).flags.ClaseMontura) Then
                'Comprobamos que no este bugueada
                If UserList(oo).Montura.Elu(UserList(oo).flags.ClaseMontura) = 0 Then
                    Call SendData(ToGM, 0, 0, "|| Matanpc Mascota Bugueada: " & UserList(oo).Name & "´" & FontTypeNames.FONTTYPE_COMERCIO)
                    Call LogCasino("BUG MataNpcMASCOTAparty Serie: " & UserList(oo).Serie & " IP: " & UserList(oo).ip & " Nom: " & UserList(oo).Name)
                End If
'----------------
                'Le metemos la exp a la montura
                Call AddtoVar(UserList(oo).Montura.exp(UserList(oo).flags.ClaseMontura), Int(expIndi / 1000), MAXEXP)
                Call CheckMonturaLevel(oo)
            End If
        End If 'topelevel
                        
                        
                        
                        
                        
                        
                        
                        
                        'expIndi = (NpcIndex.GiveEXP * (partylist(UserList(UserIndex).flags.partyNum).miembros(n).privi) / 100)
                         'expIndi = NpcIndex.GiveEXP * (partylist(UserList(UserIndex).flags.partyNum).miembros(n).privi) / 100)
                        'expIndi = expIndi \ 100
                        
                        'pluto:6.3 AUMENTO EXP----------
    'If ServerPrimario = 1 Then
                'If UserList(oo).Remort > 0 Then
                'expIndi = expIndi * 2
                'GoTo Nomire
                'End If

                       ' Select Case UserList(oo).Stats.ELV
                       ' Case Is < 50
                       ' expIndi = expIndi * 10
                        'Case 50 To 60
                        'expIndi = expIndi * 5
                        'Case Is > 60
                        'expIndi = expIndi * 3
                        'End Select
     'Else 'secundario Pluto:6.5
                        
              ' If UserList(oo).Remort > 0 Then
               ' expIndi = expIndi * 1
                'GoTo Nomire
                'End If

                 '       Select Case UserList(oo).Stats.ELV
                  '      Case Is < 30
                   '     expIndi = expIndi * 10
                    '    Case 30 To 40
                     '   expIndi = expIndi * 5
                      '  Case 41 To 50
                       ' expIndi = expIndi * 3
                        'Case Is > 50
                        'expIndi = expIndi * 2
                        'End Select
                        
                        
    'End If 'primario =1
                        'If UserList(oo).Remort > 0 Then expIndi = expIndi * 2
                        'Debug.Print UserList(oo).Name
                        '-----------------------------
Nomire:
                        Call AddtoVar(UserList(oo).Stats.exp, expIndi, MAXEXP)
                        Call SendData(ToIndex, oo, 0, "V6" & expIndi & "," & aa)
                        Call CheckUserLevel(oo)
            
                    End If
                End If
            End If
        End If
        Next
Exit Sub
errhandler:
    Call LogError("Error en PartyReparteExp Nom:" & UserList(UserIndex).Name & " UI: " & UserIndex & " NPCID: " & NpcIndex.Name & " N: " & Err.number & " D: " & Err.Description)
End Sub

Function partyid(ByVal liderName As String) As Integer
    On Error GoTo errhandler
    Dim n As Integer
    Dim encontrado As Boolean
    encontrado = False
    n = 0
    Do While (n < numPartys And encontrado <> True)
        n = n + 1
        If UCase$(UserList(partylist(n).lider).Name) = UCase$(liderName) Then
            encontrado = True
        End If
    Loop
    If encontrado = True Then
        partyid = n
    End If
    Exit Function
errhandler:
        Call LogError("Error en partyid NomLider:" & liderName & " N: " & Err.number & " D: " & Err.Description)
End Function
Function esLider(ByVal UserIndex As Integer) As Boolean
On Error GoTo errhandler

    esLider = False
    If UserList(UserIndex).flags.party = True Then
        If UserIndex = partylist(UserList(UserIndex).flags.partyNum).lider Then
            esLider = True
        End If
    End If
Exit Function
errhandler:
    Call LogError("Error en esLider Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)
End Function
Sub sendExpParty(exp As Long, UserIndex As Integer)

On Error GoTo errhandler
            Dim n As Integer
            For n = 1 To partylist(UserList(UserIndex).flags.partyNum).numMiembros
                If partylist(UserList(UserIndex).flags.partyNum).miembros(n).ID <> 0 Then
                    If UserList(partylist(UserList(UserIndex).flags.partyNum).miembros(n).ID).flags.Muerto = 0 Then
                        Call SendData(ToIndex, partylist(UserList(UserIndex).flags.partyNum).miembros(n).ID, 0, "V6" & (exp * UserList(partylist(UserList(UserIndex).flags.partyNum).miembros(n).ID).Stats.ELV) / totalexpParty(UserList(UserIndex).flags.partyNum) & ",")
                    End If
                End If
            Next
Exit Sub
errhandler:
    Call LogError("Error en sendExpParty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " exp:" & exp & " N: " & Err.number & " D: " & Err.Description)
End Sub
Sub sendMiembrosParty(UserIndex As Integer)
On Error GoTo errhandler
            Dim miempar$
            Dim npar As Byte
            If UserList(UserIndex).flags.party = False Then Exit Sub
            npar = 1
            miempar$ = partylist(UserList(UserIndex).flags.partyNum).numMiembros & ", "
            Do While (npar <= MAXMIEMBROS)
                If partylist(UserList(UserIndex).flags.partyNum).miembros(npar).ID <> 0 Then
                'pluto:6.3-------
                Call SendData(ToIndex, partylist(UserList(UserIndex).flags.partyNum).miembros(npar).ID, 0, "O4" & UserList(UserIndex).flags.partyNum)
                '-----------------
                    miempar$ = miempar$ & UserList(partylist(UserList(UserIndex).flags.partyNum).miembros(npar).ID).Name & "," & UserList(partylist(UserList(UserIndex).flags.partyNum).miembros(npar).ID).Char.CharIndex & ","
                End If
                npar = npar + 1
            Loop
            npar = 1
            Do While (npar <= MAXMIEMBROS)
                If partylist(UserList(UserIndex).flags.partyNum).miembros(npar).ID <> 0 Then
                    Call SendData(ToIndex, partylist(UserList(UserIndex).flags.partyNum).miembros(npar).ID, 0, "W1" & miempar$)
               
                End If
                npar = npar + 1
            Loop
Exit Sub
errhandler:
    Call LogError("Error en sendMiembrosParty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)
End Sub
Sub sendPriviParty(UserIndex As Integer)
On Error GoTo errhandler
    Dim miempar$
    Dim npar As Byte
    If UserList(UserIndex).flags.party = False Then Exit Sub
    npar = 1
    miempar$ = partylist(UserList(UserIndex).flags.partyNum).numMiembros & ", "
    Do While (npar <= MAXMIEMBROS)
        If partylist(UserList(UserIndex).flags.partyNum).miembros(npar).ID <> 0 Then
            miempar$ = miempar$ & partylist(UserList(UserIndex).flags.partyNum).miembros(npar).privi & ", "
        End If
        npar = npar + 1
    Loop
    Call SendData(toParty, UserIndex, 0, "W3" & miempar$)
Exit Sub
errhandler:
    Call LogError("Error en sendPriviParty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)

End Sub

Sub sendSolicitudesParty(UserIndex As Integer)
On Error GoTo errhandler
    If esLider(UserIndex) = True Then
        Dim miempar2$
        Dim npar As Byte
        npar = 1
        miempar2$ = partylist(UserList(UserIndex).flags.partyNum).numSolicitudes & ", "
        Do While (npar <= MAXMIEMBROS)
            If partylist(UserList(UserIndex).flags.partyNum).Solicitudes(npar) <> 0 Then
                miempar2$ = miempar2$ & UserList(partylist(UserList(UserIndex).flags.partyNum).Solicitudes(npar)).Name & ", "
            End If
                npar = npar + 1
        Loop
        Call SendData(ToIndex, UserIndex, 0, "W2" & miempar2$)
    End If
Exit Sub
errhandler:
    Call LogError("Error en sendSolicitudesParty Nom:" & UserList(UserIndex).Name & " UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)
End Sub
Sub resetParty(PartyIndex As Integer)
'completar
On Error GoTo errhandler
'pluto:6.5
Call LogParty("Party: " & PartyIndex & " reseteada.")

        partylist(PartyIndex).expAc = 0
        partylist(PartyIndex).lider = 0
        partylist(PartyIndex).numMiembros = 0
        partylist(PartyIndex).numSolicitudes = 0
        partylist(PartyIndex).reparto = 1
        Dim lpp2 As Integer
        For lpp2 = 1 To MAXMIEMBROS
            partylist(partyidex).miembros(lpp2).ID = 0
            partylist(partyidex).miembros(lpp2).privi = 0
            partylist(partyidex).Solicitudes(lpp2) = 0
        Next
        numPartys = numPartys - 1
 'pluto:6.5
Call LogParty("Numero de Partys: " & numPartys)
Exit Sub
errhandler:
    Call LogError("Error en resetparty PID:" & PartyIndex & " N: " & Err.number & " D: " & Err.Description)
End Sub

Sub BalanceaPrivisLVL(PartyIndex As Integer)
On Error GoTo errhandler
Dim n As Integer
For n = 1 To MAXMIEMBROS 'partylist(PartyIndex).numMiembros
    If partylist(PartyIndex).miembros(n).ID <> 0 Then
        partylist(PartyIndex).miembros(n).privi = (UserList(partylist(PartyIndex).miembros(n).ID).Stats.ELV * 100) \ totalexpParty(PartyIndex)
    End If
Next
Exit Sub
errhandler:
    Call LogError("Error en BalanceaPrivisLVL PID:" & PartyIndex & " N: " & Err.number & " D: " & Err.Description)
End Sub
Sub BalanceaPrivisMiembros(PartyIndex As Integer)
On Error GoTo errhandler
Dim n As Integer
For n = 1 To MAXMIEMBROS 'partylist(PartyIndex).numMiembros
    If partylist(PartyIndex).miembros(n).ID <> 0 Then
        partylist(PartyIndex).miembros(n).privi = 100 \ partylist(PartyIndex).numMiembros
    End If
Next
Exit Sub
errhandler:
    Call LogError("Error en BalanceaPrivisMiembros PID:" & PartyIndex & " N: " & Err.number & " D: " & Err.Description)
End Sub

