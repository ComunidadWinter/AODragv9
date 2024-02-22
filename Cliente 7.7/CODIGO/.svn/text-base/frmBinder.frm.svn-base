VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmBinder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Macros"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBinder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restaurar configuración original"
      Height          =   495
      Left            =   60
      MouseIcon       =   "frmBinder.frx":000C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6540
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Caption         =   "Macros definidos"
      Height          =   4275
      Left            =   60
      TabIndex        =   3
      Top             =   2160
      Width           =   5955
      Begin VB.CommandButton cmdAdd 
         Height          =   375
         Left            =   4320
         MaskColor       =   &H00000000&
         MouseIcon       =   "frmBinder.frx":015E
         MousePointer    =   99  'Custom
         Picture         =   "frmBinder.frx":02B0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Agregar Macro (Insert)"
         Top             =   3780
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CommandButton cmdEdit 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   4740
         MaskColor       =   &H00000000&
         MouseIcon       =   "frmBinder.frx":0884
         MousePointer    =   99  'Custom
         Picture         =   "frmBinder.frx":09D6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Editar Macro"
         Top             =   3780
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CommandButton cmdRemove 
         Height          =   375
         Left            =   5160
         MaskColor       =   &H00000000&
         MouseIcon       =   "frmBinder.frx":0BDA
         MousePointer    =   99  'Custom
         Picture         =   "frmBinder.frx":0D2C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Eliminar Macro (Del/Supr)"
         Top             =   3780
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin MSComctlLib.ListView lvwMacros 
         Height          =   3975
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MouseIcon       =   "frmBinder.frx":1300
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tecla(s)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comando"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Guardar"
      Height          =   315
      Left            =   3540
      MouseIcon       =   "frmBinder.frx":1462
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6540
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveClose 
      Caption         =   "G&uardar y Cerrar"
      Height          =   315
      Left            =   4560
      MouseIcon       =   "frmBinder.frx":15B4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6540
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   315
      Left            =   2520
      MouseIcon       =   "frmBinder.frx":1706
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6540
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   1920
      Y1              =   6480
      Y2              =   7080
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   6000
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label3 
      Caption         =   "Lee con atención:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"frmBinder.frx":1858
      Height          =   1035
      Left            =   60
      TabIndex        =   10
      Top             =   300
      Width           =   5955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Utiliza los botones que se encuentran en la parte inferior del cuadro para Agregar, Editar y Eliminar macros"
      Height          =   555
      Left            =   120
      TabIndex        =   8
      Top             =   1620
      Width           =   5835
   End
End
Attribute VB_Name = "frmBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public bChangesMade As Boolean

Private Sub cmdAdd_Click()
    Call frmBinderEdit.Show(vbModal, frmBinder)
End Sub

Private Sub cmdClose_Click()
    Dim iVal As Integer
    If bChangesMade Then
        iVal = MsgBox("Se han hecho cambios a los macros, ¿desea guardarlos?", vbApplicationModal + vbExclamation + vbYesNo, "Cerrar sin guardar")
        If iVal = vbYes Then
            Call cmdSaveClose_Click
        Else
            Call Unload(Me)
        End If
    End If
        
    Call Unload(Me)
    frmMain.SetFocus
End Sub

Private Sub cmdEdit_Click()
    Call Load(frmBinderEdit)
    Dim lX As Long, sLetra As String, sShift As String, sAction As String
    Dim sBind As String: sBind = lvwMacros.SelectedItem.Text
    
    If Len(sBind) > 1 Then
        For lX = 1 To Len(sBind)
            sLetra = Mid(sBind, lX, 1)
            If sLetra = "+" Then
                sShift = Left(sBind, lX - 1)
                Exit For
            End If
        Next lX
        sAction = Right(sBind, 1)
    Else
        sAction = sBind
    End If
    
    With frmBinderEdit
        .lShift = GetIndexFromComboStr(.cmbShift, sShift)
        .lKey = GetIndexFromComboStr(.cmbKey, sAction)
        If Left(lvwMacros.SelectedItem.SubItems(1), 1) = "/" Then
            .sCommand = lvwMacros.SelectedItem.SubItems(1)
            .optAction(0).value = True
        ElseIf Left(lvwMacros.SelectedItem.SubItems(1), 1) = "-" Then
            .sCommComp = lvwMacros.SelectedItem.SubItems(1)
            .optAction(2).value = True
        Else
            .lFunc = GetIndexFromComboStr(.cmbFunc, lvwMacros.SelectedItem.SubItems(1), True)
            .optAction(1).value = True
        End If
    End With

    frmBinderEdit.InitValues (True)
    Call frmBinderEdit.Show(vbModal, frmBinder)
End Sub

Private Sub cmdRemove_Click()
    Dim iX As Integer: iX = 1
    
    Do Until iX = lvwMacros.ListItems.Count + 1
        If lvwMacros.ListItems(iX).Selected = True Then
            Call lvwMacros.ListItems.Remove(iX)
        Else
            iX = iX + 1
        End If
    Loop
    
    bChangesMade = True
End Sub

Private Sub cmdRestore_Click()
    lvwMacros.ListItems.Clear
    
    ' funciones predefinidas
    lvwMacros.ListItems.Add , "Ca", "C": lvwMacros.ListItems("Ca").SubItems(1) = "+modocombate"
    lvwMacros.ListItems.Add , "Sa", "S": lvwMacros.ListItems("Sa").SubItems(1) = "+seguro"
    lvwMacros.ListItems.Add , "Ma", "M": lvwMacros.ListItems("Ma").SubItems(1) = "+musica"
    lvwMacros.ListItems.Add , "Oa", "O": lvwMacros.ListItems("Oa").SubItems(1) = "+ocultarse"
    lvwMacros.ListItems.Add , "Ra", "R": lvwMacros.ListItems("Ra").SubItems(1) = "+robar"
    lvwMacros.ListItems.Add , "Aa", "A": lvwMacros.ListItems("Aa").SubItems(1) = "+agarrar"
    lvwMacros.ListItems.Add , "Ea", "E": lvwMacros.ListItems("Ea").SubItems(1) = "+equipar"
    lvwMacros.ListItems.Add , "Na", "N": lvwMacros.ListItems("Na").SubItems(1) = "+nombres"
    lvwMacros.ListItems.Add , "Da", "D": lvwMacros.ListItems("Da").SubItems(1) = "+domar"
    lvwMacros.ListItems.Add , "Ta", "T": lvwMacros.ListItems("Ta").SubItems(1) = "+tirar"
    lvwMacros.ListItems.Add , "Ua", "U": lvwMacros.ListItems("Ua").SubItems(1) = "+usar"
    
    ' comandos
    lvwMacros.ListItems.Add , "CTRL+Ca", "CTRL+C": lvwMacros.ListItems("CTRL+Ca").SubItems(1) = "/comerciar"
    lvwMacros.ListItems.Add , "CTRL+Xa", "CTRL+X": lvwMacros.ListItems("CTRL+Xa").SubItems(1) = "/salir"
    lvwMacros.ListItems.Add , "CTRL+Ma", "CTRL+M": lvwMacros.ListItems("CTRL+Ma").SubItems(1) = "/meditar"
    lvwMacros.ListItems.Add , "CTRL+Da", "CTRL+D": lvwMacros.ListItems("CTRL+Da").SubItems(1) = "/descansar"
    lvwMacros.ListItems.Add , "CTRL+Ea", "CTRL+E": lvwMacros.ListItems("CTRL+Ea").SubItems(1) = "/enlistar"
    lvwMacros.ListItems.Add , "CTRL+Ia", "CTRL+I": lvwMacros.ListItems("CTRL+Ia").SubItems(1) = "/informacion"
    lvwMacros.ListItems.Add , "CTRL+Ra", "CTRL+R": lvwMacros.ListItems("CTRL+Ra").SubItems(1) = "/recompensa"
    lvwMacros.ListItems.Add , "CTRL+Ga", "CTRL+G": lvwMacros.ListItems("CTRL+Ga").SubItems(1) = "/gm ayuda"
    
    ' comandos compuestos
    lvwMacros.ListItems.Add , "SHIFT+Ba", "SHIFT+B": lvwMacros.ListItems("SHIFT+Ba").SubItems(1) = "-bug "
    lvwMacros.ListItems.Add , "SHIFT+Da", "SHIFT+D": lvwMacros.ListItems("SHIFT+Da").SubItems(1) = "-desc "
    lvwMacros.ListItems.Add , "SHIFT+Pa", "SHIFT+P": lvwMacros.ListItems("SHIFT+Pa").SubItems(1) = "-passwd "
    lvwMacros.ListItems.Add , "SHIFT+Ra", "SHIFT+R": lvwMacros.ListItems("SHIFT+Ra").SubItems(1) = "-retirar "
    lvwMacros.ListItems.Add , "SHIFT+Ea", "SHIFT+E": lvwMacros.ListItems("SHIFT+Ea").SubItems(1) = "-depositar "
    
End Sub

Private Sub cmdSave_Click()
    Call SaveBinds
End Sub

Private Sub cmdSaveClose_Click()
    Call SaveBinds
    Call Unload(Me)
    frmMain.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyInsert
            Call cmdAdd_Click
        Case vbKeyDelete
            Call cmdRemove_Click
    End Select
End Sub

Private Sub Form_Load()
    Dim iOpenPos As Integer: iOpenPos = FreeFile()
    Dim sLine As String
    
    'On Error GoTo onerror
    
'[CODE 005]:MatuX
    Open App.Path & "\init\default.bnd" For Input Access Read Lock Write As #iOpenPos
'[END]
        On Error Resume Next
        Call LookFor("[BINDS]", iOpenPos)
        Do Until EOF(iOpenPos)
            Line Input #iOpenPos, sLine
            Call ParseBind(sLine)
        Loop
        On Error GoTo 0
    Close #iOpenPos

onerror:
    bChangesMade = False
    'Close #iOpenPos
    'On Error GoTo 0
    'Err.Clear
End Sub

Private Sub SaveBinds()
    Dim iOpenPos As Integer: iOpenPos = FreeFile()
    Dim lX As Long
    
'[CODE 005]:MatuX
    Open App.Path & "\init\default.bnd" For Output Access Write Lock Write As #iOpenPos
'[END]
        'Call LookFor("[BINDS]", iOpenPos)
        Print #iOpenPos, "//"
        Print #iOpenPos, "// No modificar este archivo a menos que"
        Print #iOpenPos, "// se esté seguro de lo que se hace!"
        Print #iOpenPos, "//"
        Print #iOpenPos,
        Print #iOpenPos, "[BINDS]"
        
        For lX = 1 To lvwMacros.ListItems.Count
            Print #iOpenPos, lvwMacros.ListItems(lX).Text & "=" & lvwMacros.ListItems(lX).SubItems(1)
        Next lX
    Close #iOpenPos
End Sub

Private Function ParseBind(sBind As String) As Long
    'cogemos la letra y su combinación, tío
    Dim sLetra As String, lX As Long
    Dim sShift As String, sAction As String
    
    On Error GoTo errHnd:
    
    For lX = 1 To Len(sBind)
        sLetra = Mid(sBind, lX, 1)
        If sLetra = "=" Then
            sShift = Left(sBind, lX - 1)
            Exit For
        End If
    Next lX
    sAction = Right(sBind, Len(sBind) - (lX))
    
    If sShift = "" Or sAction = "" Then _
        Call Err.Raise(1)

    lvwMacros.ListItems.Add , sShift & "a", sShift
    lvwMacros.ListItems(sShift & "a").SubItems(1) = sAction

    ParseBind = 0
    Exit Function
        
errHnd:
    ParseBind = -1

End Function

Public Function GetIndexFromComboStr(cmbCombo As ComboBox, sString As String, Optional UseInstr As Boolean = False) As Long
    Dim iX As Long
    
    If Not UseInstr Then
        Do Until LCase(cmbCombo.List(iX)) = LCase(sString)
            iX = iX + 1
        Loop
    Else
        Do Until InStr(LCase(cmbCombo.List(iX)), LCase(sString))
            iX = iX + 1
        Loop
    End If
    
    GetIndexFromComboStr = iX
End Function
