VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Height          =   4650
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   5400
   End
   Begin VB.ListBox lstArmaduras 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   5400
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2040
      MouseIcon       =   "frmHerrero.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":0CCA
      Top             =   840
      Width           =   1650
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   360
      MouseIcon       =   "frmHerrero.frx":51C9
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":5E93
      Top             =   840
      Width           =   1650
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Construcción de Objetos"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   720
      MouseIcon       =   "frmHerrero.frx":9F63
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":AC2D
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   3480
      MouseIcon       =   "frmHerrero.frx":B255
      MousePointer    =   99  'Custom
      Picture         =   "frmHerrero.frx":BF1F
      Top             =   6000
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Botón derecho sobre el objeto para ver los materiales necesarios."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6480
      Width           =   5535
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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

Private Sub Form_Load()
frmHerrero.Picture = LoadPicture(DirGraficos & "ventanas.jpg")
End Sub
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()



End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
On Error Resume Next
'pluto:6.2
    If UserCanAttack = 1 Then
    UserCanAttack = 0
If lstArmas.Visible Then
 Call SendData("CNS" & ArmasHerrero(lstArmas.ListIndex))
Else
 Call SendData("CNS" & ArmadurasHerrero(lstArmaduras.ListIndex))
End If
    End If
'Unload Me
End Sub

Private Sub Image3_Click()
lstArmaduras.Visible = False
lstArmas.Visible = True
End Sub

Private Sub Image4_Click()
lstArmaduras.Visible = True
lstArmas.Visible = False
End Sub

Private Sub lstArmaduras_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbRightButton) Then
      If lstArmaduras.ListIndex < 0 Then Exit Sub
   frmMain.TimerLabel.Enabled = True
    
Dim Esti As Integer

    Esti = ArmadurasHerrero(lstArmaduras.ListIndex)

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
ob = ArmadurasHerrero(lstArmaduras.ListIndex)
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

Private Sub lstArmas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbRightButton) Then
       If lstArmas.ListIndex < 0 Then Exit Sub
   frmMain.TimerLabel.Enabled = True
    
Dim Esti As Integer

    Esti = ArmasHerrero(lstArmas.ListIndex)

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
ob = ArmasHerrero(lstArmas.ListIndex)
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
