VERSION 5.00
Begin VB.Form frmBancoObj 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   486
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRetiDepo 
      Caption         =   "Depositar"
      Height          =   360
      Index           =   1
      Left            =   4320
      TabIndex        =   19
      Top             =   7800
      Width           =   2610
   End
   Begin VB.CommandButton cmdRetiDepo 
      Caption         =   "Retirar"
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   7800
      Width           =   2610
   End
   Begin VB.CommandButton cCerrar 
      Caption         =   "X"
      Height          =   240
      Left            =   7080
      TabIndex        =   17
      Top             =   0
      Width           =   210
   End
   Begin VB.CommandButton cmdDespositar 
      Caption         =   "Depositar"
      Height          =   360
      Left            =   4440
      TabIndex        =   16
      Top             =   960
      Width           =   1335
   End
   Begin VB.PictureBox PicBancoInv 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   120
      ScaleHeight     =   284
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   227
      TabIndex        =   14
      Top             =   3240
      Width           =   3465
   End
   Begin VB.PictureBox PicInv 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   3720
      ScaleHeight     =   284
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   227
      TabIndex        =   13
      Top             =   3240
      Width           =   3465
   End
   Begin VB.Frame Frame1 
      Caption         =   "Deposito de oro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton cmdRetirar 
         Caption         =   "Retirar"
         Height          =   360
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxTCantidadGLD 
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Text            =   "1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblBilletera 
         Alignment       =   2  'Center
         Caption         =   "Tienes 00000000000 monedas de oro en tu billetera."
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   5535
      End
      Begin VB.Label lblDeposito 
         Alignment       =   2  'Center
         Caption         =   "Actuamente tienes un deposito de 000000000 monedas de oro."
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   5340
      End
   End
   Begin VB.TextBox cantidad 
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
      Left            =   3120
      TabIndex        =   2
      Text            =   "1"
      Top             =   7920
      Width           =   960
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label lblCantidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3240
      TabIndex        =   3
      Top             =   7680
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3990
      TabIndex        =   1
      Top             =   3015
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3990
      TabIndex        =   0
      Top             =   2670
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public NoPuedeMover As Boolean

Private Sub cantidad_Change()

    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub CantidadOro_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cCerrar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call WriteBankEnd
    NoPuedeMover = False
End Sub

Private Sub cmdDespositar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call WriteBankDepositGold(TxTCantidadGLD.Text)
End Sub

Private Sub CMDRetirar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call WriteBankExtractGold(TxTCantidadGLD.Text)
End Sub

Private Sub Form_Activate()
'on error resume next
    InvBanco(0).DrawInv
    InvBanco(1).DrawInv
End Sub

Private Sub Form_GotFocus()
'on error resume next
    InvBanco(0).DrawInv
    InvBanco(1).DrawInv
End Sub

Private Sub cmdRetiDepo_Click(Index As Integer)
    
    Call Sound.Sound_Play(SND_CLICK)
    
    If InvBanco(Index).SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    
    Select Case Index
        Case 0
            LastIndex1 = InvBanco(0).SelectedItem
            LasActionBuy = True
            Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
            
       Case 1
            LastIndex2 = InvBanco(1).SelectedItem
            LasActionBuy = False
            Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
    End Select

End Sub

Private Sub PicBancoInv_Click()
    Call Sound.Sound_Play(SND_CLICK)

    InvBanco(1).DrawInv
    If InvBanco(0).SelectedItem <> 0 Then
        With UserBancoInventory(InvBanco(0).SelectedItem)
            lblNombre.Caption = .name
            
            Select Case .OBJType
                Case 2, 32
                    lblInfo.Caption = "Máx Golpe:" & .MaxHit & "/Mín Golpe:" & .MinHit
                    lblInfo.Visible = True
                    
                Case 3, 16, 17
                    lblInfo.Caption = "Defensa:" & .MaxDef & "/Mín Defensa:" & .MinDef
                    lblInfo.Visible = True

                    
                Case Else
                    lblInfo.Visible = False
                    
            End Select
            
        End With
        
    Else
        lblNombre.Caption = vbNullString
        lblInfo.Visible = False
    End If

End Sub

Private Sub PicInv_Click()
    Call Sound.Sound_Play(SND_CLICK)

    InvBanco(1).DrawInv
    If InvBanco(1).SelectedItem <> 0 Then
        With Inventario
            lblNombre.Caption = .ItemName(InvBanco(1).SelectedItem)
            
            Select Case .OBJType(InvBanco(1).SelectedItem)
                Case eObjType.otWeapon, eObjType.otFlechas
                    lblInfo.Caption = "Máx Golpe:" & .MaxHit(InvBanco(1).SelectedItem) & "/Mín Golpe:" & .MinHit(InvBanco(1).SelectedItem)
                    lblInfo.Visible = True
                    
                Case eObjType.otcasco, eObjType.otArmadura, eObjType.otescudo ' 3, 16, 17
                    lblInfo.Caption = "Máx Defensa:" & .MaxDef(InvBanco(1).SelectedItem) & "/Mín Defensa:" & .MinDef(InvBanco(1).SelectedItem)
                    lblInfo.Visible = True
                    
                Case Else
                    lblInfo.Visible = False
                    
            End Select
            
        End With
    Else
        lblNombre.Caption = vbNullString
        lblInfo.Visible = False
    End If
End Sub


