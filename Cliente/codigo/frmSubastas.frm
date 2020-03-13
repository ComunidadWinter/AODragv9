VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSubastas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Subastas Goliath"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   406
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraSubastar 
      Caption         =   "Subastar"
      Height          =   5895
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdSubastar 
         Caption         =   "Nueva Subasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   16
         Top             =   5400
         Width           =   1530
      End
      Begin VB.ComboBox ComHoras 
         Height          =   315
         ItemData        =   "frmSubastas.frx":0000
         Left            =   240
         List            =   "frmSubastas.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox txtPrecioInicial 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Text            =   "1"
         Top             =   3960
         Width           =   2295
      End
      Begin VB.ListBox lstItems 
         Height          =   2790
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Text            =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.PictureBox ImgItemSubastar 
         BackColor       =   &H00000000&
         Height          =   480
         Left            =   120
         ScaleHeight     =   420
         ScaleWidth      =   420
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblComisión0 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comisión: 0 moneda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   17
         Top             =   5040
         Width           =   1995
      End
      Begin VB.Label lblDuración 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Duración"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   4320
         Width           =   630
      End
      Begin VB.Label lblPrecioInicial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de compra"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   3720
         Width           =   1230
      End
      Begin VB.Label lblCantidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame FraBuscarEn 
      Caption         =   "Buscar en Subasta"
      Height          =   5895
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.TextBox txtMax 
         Height          =   285
         Left            =   4320
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtMin 
         Height          =   285
         Left            =   3480
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9720
         TabIndex        =   6
         Top             =   5400
         Width           =   930
      End
      Begin VB.CommandButton cmdComprar 
         Caption         =   "Comprar"
         Height          =   360
         Left            =   8640
         TabIndex        =   5
         Top             =   5400
         Width           =   930
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5160
         TabIndex        =   2
         Top             =   360
         Width           =   1290
      End
      Begin VB.PictureBox picItem 
         BackColor       =   &H00000000&
         Height          =   480
         Left            =   240
         ScaleHeight     =   420
         ScaleWidth      =   420
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   480
      End
      Begin MSFlexGridLib.MSFlexGrid Items 
         Height          =   4335
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   7
         ScrollBars      =   2
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   23
         Top             =   480
         Width           =   90
      End
      Begin VB.Label lblRango 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rango de precios"
         Height          =   195
         Left            =   3480
         TabIndex        =   20
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Left            =   840
         TabIndex        =   18
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lbloro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tienes 0 monedas de oro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   5400
         Width           =   2700
      End
   End
End
Attribute VB_Name = "frmSubastas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    Call WriteSalirSubasta
End Sub

Private Sub cmdSubastar_Click()
    If txtPrecioInicial.Text < 0 And txtPrecioInicial.Text > 999999999 Then
        MsgBox "Precio invalido"
        Exit Sub
    End If
    
    If lstItems.ListIndex > MAX_INVENTORY_SLOTS Then
        MsgBox "Item invalido"
        Exit Sub
    End If
    
    If ComHoras.ListIndex < 0 Or ComHoras.ListIndex > 3 Then
        MsgBox "Duracion invalida"
        Exit Sub
    End If
    
    If txtCantidad < 0 Or txtCantidad > 999 Then
        MsgBox "Cantidad invalida"
        Exit Sub
    End If
    
    Call WriteNuevaSubasta(lstItems.ListIndex, txtPrecioInicial.Text, ComHoras.ListIndex, txtCantidad.Text)
    
End Sub

Private Sub Form_Load()
    ComHoras.ListIndex = 1
End Sub
