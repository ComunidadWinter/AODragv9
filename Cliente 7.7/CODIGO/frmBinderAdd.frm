VERSION 5.00
Begin VB.Form frmBinderEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Macro"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmComm 
      Caption         =   "Escribe el comando:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   3495
      Begin VB.TextBox txtCommand 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "/"
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame frmCommComp 
      Caption         =   "Escribe el comando compuesto:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtCommComp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Text            =   "-"
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.OptionButton optAction 
      Caption         =   "Comando Compuesto (ej: /desc Gran Guerrero, amigo de... o /depositar 1000)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   240
      MouseIcon       =   "frmBinderAdd.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   2220
      Width           =   3195
   End
   Begin VB.Frame frmFunc 
      Caption         =   "Selecciona la función:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ComboBox cmbFunc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmBinderAdd.frx":0152
         Left            =   120
         List            =   "frmBinderAdd.frx":0177
         MouseIcon       =   "frmBinderAdd.frx":02BD
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.OptionButton optAction 
      Caption         =   "Función Predefinida (ej: activar modo combate, apagar sonidos)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      MouseIcon       =   "frmBinderAdd.frx":040F
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   2880
      Width           =   3195
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   1620
      MouseIcon       =   "frmBinderAdd.frx":0561
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4860
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   315
      Left            =   2640
      MouseIcon       =   "frmBinderAdd.frx":06B3
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   4860
      Width           =   975
   End
   Begin VB.OptionButton optAction 
      Caption         =   "Comando (ej: /comerciar o /balance)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   240
      MouseIcon       =   "frmBinderAdd.frx":0805
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1800
      Value           =   -1  'True
      Width           =   3195
   End
   Begin VB.ComboBox cmbShift 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmBinderAdd.frx":0957
      Left            =   240
      List            =   "frmBinderAdd.frx":0964
      MouseIcon       =   "frmBinderAdd.frx":0977
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cmbKey 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmBinderAdd.frx":0AC9
      Left            =   1500
      List            =   "frmBinderAdd.frx":0B39
      MouseIcon       =   "frmBinderAdd.frx":0BA9
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Line Line3 
      X1              =   60
      X2              =   3660
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   $"frmBinderAdd.frx":0CFB
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   5
      Top             =   3420
      Width           =   3315
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   3600
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Selecciona la acción que desarrollará este macro:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   3
      Top             =   1260
      Width           =   3315
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   3600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Seleccioná la letra o combinación que utilizaras para este macro:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3315
   End
End
Attribute VB_Name = "frmBinderEdit"
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

Public lShift As Long, lKey As Long, lAction As Long, _
       lFunc As Long, sCommand As String, sCommComp As String, _
       bEditMode As Boolean

Private Sub cmdAdd_Click()
    'On Error Resume Next
        Dim sMacro As String, iVal As Integer
        sMacro = IIf(cmbShift.Text = "", cmbKey.Text, cmbShift.Text & "+" & cmbKey.Text)
        
        If LVWCompare(sMacro, frmBinder.lvwMacros) Then
            Call UnSelectAll(frmBinder.lvwMacros)
            frmBinder.lvwMacros.ListItems(sMacro & "a").Selected = True
            frmBinder.lvwMacros.ListItems(sMacro & "a").EnsureVisible
            iVal = MsgBox("El macro ya existe, ¿desea sobre-escribirlo?", vbApplicationModal + vbQuestion + vbYesNo, "El macro ya fue definido")
            If iVal = vbNo Then
                Exit Sub
            Else
                bEditMode = True
            End If
        End If
        
        If bEditMode Then
            Call frmBinder.lvwMacros.ListItems.Remove(frmBinder.lvwMacros.SelectedItem.Key)
        '    frmBinder.lvwMacros.SelectedItem.Key = sMacro
        '    frmBinder.lvwMacros.ListItems(sMacro).Text = sMacro
        End If
        'Else
            frmBinder.lvwMacros.ListItems.Add , sMacro & "a", sMacro
        'End If
        
        If cmbFunc.Visible = True Then
            Dim sString As String, lPos As Long
            lPos = InStr(cmbFunc.Text, "(") + 1
            sString = Mid(cmbFunc.Text, lPos, Len(cmbFunc.Text) - lPos)
            frmBinder.lvwMacros.ListItems(sMacro & "a").SubItems(1) = sString
        ElseIf frmCommComp.Visible = True Then
            frmBinder.lvwMacros.ListItems(sMacro & "a").SubItems(1) = txtCommComp.Text
        Else
            frmBinder.lvwMacros.ListItems(sMacro & "a").SubItems(1) = txtCommand.Text
        End If
        frmBinder.lvwMacros.ListItems(sMacro & "a").EnsureVisible
    On Error GoTo 0
    
    frmBinder.bChangesMade = True
    
    'ClearValues
    Call Unload(Me)
End Sub

Private Sub cmdCancel_Click()
    'ClearValues
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    Call InitValues
End Sub

Private Sub optAction_Click(Index As Integer)
    Select Case Index
        Case 0
            frmFunc.Visible = False
            frmComm.Visible = True
            frmCommComp.Visible = False
        Case 1
            frmFunc.Visible = True
            frmComm.Visible = False
            frmCommComp.Visible = False
        Case 2
            frmCommComp.Visible = True
            frmComm.Visible = False
            frmFunc.Visible = False
    End Select
End Sub

Public Sub InitValues(Optional EditMode As Boolean = False)
    cmbShift.ListIndex = lShift
    cmbKey.ListIndex = lKey
    cmbFunc.ListIndex = lFunc
    'optAction(IIf(lAction = 0, 0, 1)).value = True
    txtCommand.Text = IIf(sCommand = "", "/", sCommand)
    txtCommComp.Text = IIf(sCommComp = "", "-", sCommComp)
    
    bEditMode = EditMode
    If bEditMode = True Then
        cmdAdd.Caption = "Modificar"
    Else
        cmdAdd.Caption = "Agregar"
    End If
End Sub

Public Sub ClearValues()
    cmbShift.ListIndex = 0
    cmbKey.ListIndex = 0
    cmbFunc.ListIndex = 0
    txtCommand.Text = ""
End Sub
