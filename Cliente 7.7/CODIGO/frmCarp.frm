VERSION 5.00
Begin VB.Form frmCarp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   4320
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   2760
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2760
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Botón derecho sobre el objeto para ver materiales"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   6855
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmCarp.Picture = LoadPicture(DirGraficos & "Carp.jpg")

End Sub


Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
'[MerLiNz:6]
'If frmCarp.Caption = "Ermitaño" Then frmCarp.Caption = "Carpintero"
'[\END]

End Sub

Private Sub Form_Deactivate()
If Me.Visible Then
Me.SetFocus
Else
Unload Me
End If
End Sub

Private Sub Image1_Click()
On Error Resume Next
'pluto:6.2
    If UserCanAttack = 1 Then
    UserCanAttack = 0


'[MerLiNz:6]
If UserClase = "Ermitaño" Then
Call SendData("CER" & ObjErmitaño(lstArmas.ListIndex))
ElseIf UserClase = "Carpintero" Then
Call SendData("CNC" & ObjCarpintero(lstArmas.ListIndex))
Else
SendData ("JOP" & UserClase)
End If
    
    End If
'[\END]
'frmCarp.Caption = "Carpintero"
'Unload Me
End Sub

Private Sub Image4_Click()
Unload Me
End Sub

Private Sub lstArmas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button = vbRightButton) Then
      If lstArmas.ListIndex < 0 Then Exit Sub
   frmMain.TimerLabel.Enabled = True
 
Dim Esti As Integer

    Esti = ObjErmitaño(lstArmas.ListIndex)

    Dim aux As Integer
    Dim Lainv As String
    Dim ob As Integer
    Dim n As Byte
    
  aux = Esti

    'If aux > 0 And aux < MAX_INVENTORY_SLOTS Then _
       ' picInv.ToolTipText = UserInventory(aux).Name
        
        
       If aux = 0 Then Exit Sub
   If IndiceLabel <> aux Then
                
        'Call ReestablecerLabel
                        
        'picInv.FontUnderline = True
        'picInv.ForeColor = vbGreen


'------------------------------------------------------------
Lainv = ""
'pluto:6.0A
ob = ObjErmitaño(lstArmas.ListIndex)
If ob = 0 Then Exit Sub
If ObjData(ob).MaxHIT > 0 Then Lainv = Lainv & vbNewLine & "Máximo Golpe: " & ObjData(ob).MaxHIT
If ObjData(ob).MinHIT > 0 Then Lainv = Lainv & vbNewLine & "Mínimo Golpe: " & ObjData(ob).MinHIT
If ObjData(ob).MaxDef > 0 Then Lainv = Lainv & vbNewLine & "Máxima Defensa: " & ObjData(ob).MaxDef
If ObjData(ob).MinDef > 0 Then Lainv = Lainv & vbNewLine & "Mínima Defensa: " & ObjData(ob).MinDef

    If ObjData(ob).MaxHIT > 0 Then

If ObjData(ob).Apuñala > 0 Then
Lainv = Lainv & vbNewLine & "Apuñala: Sí"
Else
Lainv = Lainv & vbNewLine & "Apuñala: No"
End If
If ObjData(ob).Envenena > 0 Then
Lainv = Lainv & vbNewLine & "Veneno: Sí"
Else
Lainv = Lainv & vbNewLine & "Veneno: No"
End If

    End If


If ObjData(ob).Magia > 0 Then Lainv = Lainv & vbNewLine & "Mejora Magias: " & ObjData(ob).Magia & "%"
If ObjData(ob).MaxModificador > 0 Then Lainv = Lainv & vbNewLine & "Efecto Máximo: " & ObjData(ob).MaxModificador
If ObjData(ob).MinModificador > 0 Then Lainv = Lainv & vbNewLine & "Efecto Mínimo: " & ObjData(ob).MinModificador
If ObjData(ob).DuracionEfecto > 0 Then Lainv = Lainv & vbNewLine & "Duración Efecto: " & ObjData(ob).DuracionEfecto
If ObjData(ob).MinSed > 0 Then Lainv = Lainv & vbNewLine & "Recupera Sed: " & ObjData(ob).MinSed
If ObjData(ob).MinHam > 0 Then Lainv = Lainv & vbNewLine & "Recupera Hambre: " & ObjData(ob).MinHam
If ObjData(ob).MinSta > 0 Then Lainv = Lainv & vbNewLine & "Recupera Energía: " & ObjData(ob).MinSta

If ObjData(ob).MinSkill > 0 Then
Dim ala As Integer
ala = ObjData(ob).MinSkill
If UCase$(UserClase) <> "PIRATA" And UCase$(UserClase) <> "PESCADOR" Then ala = ala * 2
Lainv = Lainv & vbNewLine & "Skill Mínimo: " & ala
End If

If ObjData(ob).objetoespecial > 0 Then
Select Case ObjData(ob).objetoespecial
Case 1
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Ahorro 33% Flechas"
Case 53
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Ahorro 50% Flechas"
Case 54
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Ahorro 75% Flechas"
Case 2
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Fuerza +5"
Case 3
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Fuerza +2"
Case 4
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Fuerza +3"
Case 5
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Agilidad +5"
Case 6
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Agilidad +2"
Case 7
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Agilidad +3"
Case 8
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Mana +100"
Case 9
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Mana +200"
Case 10
Lainv = Lainv & vbNewLine & "Habilidad Mágica: Mana +300"
End Select
End If
If ObjData(ob).nocaer > 0 Then Lainv = Lainv & vbNewLine & "Habilidad Mágica: No se cae al morir."

'peso
'If Index = 1 Then
'Lainv = Lainv & vbNewLine & "Cantidad: " & UserInventory(aux).Amount
Lainv = Lainv & vbNewLine & "Peso Unidad: " & ObjData(ob).peso & " kg."
'Lainv = Lainv & vbNewLine & "Peso Total: " & ObjData(ob).peso * UserInventory(aux).Amount & " Kg."



If ObjData(ob).SkArco > 0 And ObjData(ob).proyectil > 0 Then Lainv = Lainv & vbNewLine & "Skill Mínimo: " & ObjData(ob).SkArco
If ObjData(ob).SkArma > 0 And ObjData(ob).proyectil = 0 Then Lainv = Lainv & vbNewLine & "Skill Mínimo: " & ObjData(ob).SkArma

If ObjData(ob).Vendible = 0 Then
Lainv = Lainv & vbNewLine & "Vendible: Sí"
Else
Lainv = Lainv & vbNewLine & "Vendible: No"
End If

If ObjData(ob).razaelfa > 0 Then Lainv = Lainv & vbNewLine & "Raza: Elfos."
If ObjData(ob).RazaEnana > 0 Then Lainv = Lainv & vbNewLine & "Raza: Enanos."
If ObjData(ob).razahumana > 0 Then Lainv = Lainv & vbNewLine & "Raza: Humanos."
If ObjData(ob).razaorca > 0 Then Lainv = Lainv & vbNewLine & "Raza: Orcos."
If ObjData(ob).razavampiro > 0 Then Lainv = Lainv & vbNewLine & "Raza: Vampiros."
If ObjData(ob).Real > 0 Then Lainv = Lainv & vbNewLine & "Armada: Armada Real."
If ObjData(ob).Caos > 0 Then Lainv = Lainv & vbNewLine & "Armada: Armada del Caos."
If ObjData(ob).Hombre > 0 Then Lainv = Lainv & vbNewLine & "Sexo: Hombres."
If ObjData(ob).Mujer > 0 Then Lainv = Lainv & vbNewLine & "Sexo: Mujeres."
If ObjData(ob).ObjetoClan <> "" Then Lainv = Lainv & vbNewLine & "Clan: " & ObjData(ob).ObjetoClan
If ObjData(ob).HechizoIndex > 0 Then
Dim afeti As String
Lainv = Lainv & vbNewLine & "Skill de Magia Necesario: " & Hechizos(ObjData(ob).HechizoIndex).MinSkill
Lainv = Lainv & vbNewLine & "Mana Necesario: " & Hechizos(ObjData(ob).HechizoIndex).ManaRequerido
If Hechizos(ObjData(ob).HechizoIndex).MaxHP > 0 Then Lainv = Lainv & vbNewLine & "P.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxHP
If Hechizos(ObjData(ob).HechizoIndex).MinHP > 0 Then Lainv = Lainv & vbNewLine & "P.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinHP
If Hechizos(ObjData(ob).HechizoIndex).MaxFuerza > 0 Then Lainv = Lainv & vbNewLine & "F.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxFuerza
If Hechizos(ObjData(ob).HechizoIndex).MinFuerza > 0 Then Lainv = Lainv & vbNewLine & "F.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinFuerza
If Hechizos(ObjData(ob).HechizoIndex).MaxAgilidad > 0 Then Lainv = Lainv & vbNewLine & "A.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxAgilidad
If Hechizos(ObjData(ob).HechizoIndex).MinAgilidad > 0 Then Lainv = Lainv & vbNewLine & "A.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinAgilidad
If Hechizos(ObjData(ob).HechizoIndex).MaxHam > 0 Then Lainv = Lainv & vbNewLine & "H.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxHam
If Hechizos(ObjData(ob).HechizoIndex).MinHam > 0 Then Lainv = Lainv & vbNewLine & "H.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinHam
If Hechizos(ObjData(ob).HechizoIndex).MaxSed > 0 Then Lainv = Lainv & vbNewLine & "S.Máximo: " & Hechizos(ObjData(ob).HechizoIndex).MaxSed
If Hechizos(ObjData(ob).HechizoIndex).MinSed > 0 Then Lainv = Lainv & vbNewLine & "S.Mínimo: " & Hechizos(ObjData(ob).HechizoIndex).MinSed

'target
If Hechizos(ObjData(ob).HechizoIndex).Target = 1 Then afeti = "Sólo Usuarios."
If Hechizos(ObjData(ob).HechizoIndex).Target = 2 Then afeti = "Sólo Npc´s."
If Hechizos(ObjData(ob).HechizoIndex).Target = 3 Then afeti = "Usuarios y Npc´s."
If Hechizos(ObjData(ob).HechizoIndex).Target = 4 Then afeti = "Terreno."
Lainv = Lainv & vbNewLine & "Objetivo: " & afeti

End If

Dim pit As Byte
For n = 1 To 20
If UCase$(ObjData(ob).ClaseProhibida(n)) = UCase$(UserClase) Then pit = 1
Next
If pit = 1 Then
Lainv = Lainv & vbNewLine & "El " & UserClase & " NO puede usarlo."
Else
Lainv = Lainv & vbNewLine & "El " & UserClase & " puede usarlo."
End If
pit = 0

'-----------------------------------------------------------
'fabricación---------
If ObjData(ob).LingH > 0 Or ObjData(ob).LingP > 0 Or ObjData(ob).LingO > 0 Or ObjData(ob).Madera > 0 Or ObjData(ob).Diamantes Or ObjData(ob).Gemas > 0 Then
Lainv = Lainv & vbNewLine & vbNewLine & "Se puede Fabricar Con:"
If ObjData(ob).LingO > 0 Then Lainv = Lainv & vbNewLine & "Lingotes Oro: " & ObjData(ob).LingO
If ObjData(ob).LingP > 0 Then Lainv = Lainv & vbNewLine & "Lingotes Plata: " & ObjData(ob).LingP
If ObjData(ob).LingH > 0 Then Lainv = Lainv & vbNewLine & "Lingotes Hierro: " & ObjData(ob).LingH
If ObjData(ob).Madera > 0 Then Lainv = Lainv & vbNewLine & "Madera: " & ObjData(ob).Madera
If ObjData(ob).Diamantes > 0 Then Lainv = Lainv & vbNewLine & "Diamantes: " & ObjData(ob).Diamantes
If ObjData(ob).Gemas > 0 Then Lainv = Lainv & vbNewLine & "Gemas: " & ObjData(ob).Gemas
End If
'--------------------

    
   
    Call Mostrar_ToolTip(Form2, ObjData(ob).name & vbNewLine & String(Len(ObjData(ob).name), "_") & vbNewLine & Lainv, _
    &H80000018, vbBlack, App.Path & "\Graficos\aodrag.ico")
  
        
        
                
        IndiceLabel = aux
        
    End If
    
    End If 'boton derecho
End Sub
