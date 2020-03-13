VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   11505
   ClientLeft      =   0
   ClientTop       =   480
   ClientWidth     =   15360
   ClipControls    =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   Icon            =   "frmMainbackup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainbackup.frx":1CCA
   ScaleHeight     =   767
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton lvButtons_H1 
      BackColor       =   &H0000FFFF&
      Caption         =   "REPORTAR BUG"
      Height          =   240
      Left            =   3360
      MaskColor       =   &H0000FFFF&
      TabIndex        =   39
      Top             =   11280
      Width           =   2850
   End
   Begin VB.Frame FramHechi 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   13560
      TabIndex        =   36
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
      Begin AODrag_Cliente.uAOButton UAOInformación 
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   37
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         TX              =   ""
         ENAB            =   -1  'True
         FCOL            =   16777215
         OCOL            =   16777215
         PICE            =   "frmMainbackup.frx":25D01
         PICF            =   "frmMainbackup.frx":25D5F
         PICH            =   "frmMainbackup.frx":25DBD
         PICV            =   "frmMainbackup.frx":25E1B
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AODrag_Cliente.uAOButton UAOOlvidar 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   38
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         TX              =   ""
         ENAB            =   -1  'True
         FCOL            =   16777215
         OCOL            =   16777215
         PICE            =   "frmMainbackup.frx":25E79
         PICF            =   "frmMainbackup.frx":25ED7
         PICH            =   "frmMainbackup.frx":25F35
         PICV            =   "frmMainbackup.frx":25F93
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      Left            =   10935
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.PictureBox renderer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7860
      Left            =   135
      ScaleHeight     =   524
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   692
      TabIndex        =   22
      Top             =   2730
      Width           =   10380
      Begin VB.Frame fMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   4335
         Left            =   8880
         TabIndex        =   23
         Top             =   3720
         Visible         =   0   'False
         Width           =   1575
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Monturas"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":25FF1
            PICF            =   "frmMainbackup.frx":2604F
            PICH            =   "frmMainbackup.frx":260AD
            PICV            =   "frmMainbackup.frx":2610B
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   3360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Desconectar"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":26169
            PICF            =   "frmMainbackup.frx":261C7
            PICH            =   "frmMainbackup.frx":26225
            PICV            =   "frmMainbackup.frx":26283
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Estadisticas"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":262E1
            PICF            =   "frmMainbackup.frx":2633F
            PICH            =   "frmMainbackup.frx":2639D
            PICV            =   "frmMainbackup.frx":263FB
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Clan"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":26459
            PICF            =   "frmMainbackup.frx":264B7
            PICH            =   "frmMainbackup.frx":26515
            PICV            =   "frmMainbackup.frx":26573
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   2280
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Opciones"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":265D1
            PICF            =   "frmMainbackup.frx":2662F
            PICH            =   "frmMainbackup.frx":2668D
            PICV            =   "frmMainbackup.frx":266EB
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Tutorial"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":26749
            PICF            =   "frmMainbackup.frx":267A7
            PICH            =   "frmMainbackup.frx":26805
            PICV            =   "frmMainbackup.frx":26863
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   30
            Top             =   1920
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   ""
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":268C1
            PICF            =   "frmMainbackup.frx":2691F
            PICH            =   "frmMainbackup.frx":2697D
            PICV            =   "frmMainbackup.frx":269DB
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   31
            Top             =   2640
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Party"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":26A39
            PICF            =   "frmMainbackup.frx":26A97
            PICH            =   "frmMainbackup.frx":26AF5
            PICV            =   "frmMainbackup.frx":26B53
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   32
            Top             =   3800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Cerrar menu"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":26BB1
            PICF            =   "frmMainbackup.frx":26C0F
            PICH            =   "frmMainbackup.frx":26C6D
            PICV            =   "frmMainbackup.frx":26CCB
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   33
            Top             =   1560
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Ayuda"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":26D29
            PICF            =   "frmMainbackup.frx":26D87
            PICH            =   "frmMainbackup.frx":26DE5
            PICV            =   "frmMainbackup.frx":26E43
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AODrag_Cliente.uAOButton uAOMenu 
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   34
            Top             =   3000
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   "Salir"
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMainbackup.frx":26EA1
            PICF            =   "frmMainbackup.frx":26EFF
            PICH            =   "frmMainbackup.frx":26F5D
            PICV            =   "frmMainbackup.frx":26FBB
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Timer TimerMin 
      Interval        =   10000
      Left            =   14280
      Top             =   600
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   14760
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
   End
   Begin VB.PictureBox picSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   10920
      ScaleHeight     =   1980
      ScaleWidth      =   3870
      TabIndex        =   16
      Top             =   5520
      Width           =   3870
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   11160
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   235
      TabIndex        =   13
      Top             =   2400
      Width           =   3525
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1620
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2355
      Visible         =   0   'False
      Width           =   6750
   End
   Begin VB.PictureBox Minimap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   8790
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   420
      Width           =   1530
      Begin VB.Shape GroupMember 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Index           =   4
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Shape GroupMember 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Index           =   3
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Shape GroupMember 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Shape GroupMember 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Index           =   1
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Shape GroupMember 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Index           =   0
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Shape UserM 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   90
         Left            =   750
         Shape           =   3  'Circle
         Top             =   750
         Width           =   90
      End
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   2130
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   120
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   3757
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMainbackup.frx":27019
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   14280
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14760
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5685
      TabIndex        =   43
      Top             =   10950
      Width           =   615
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4020
      TabIndex        =   42
      Top             =   10950
      Width           =   615
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2505
      TabIndex        =   41
      Top             =   10950
      Width           =   615
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   780
      TabIndex        =   40
      Top             =   10950
      Width           =   615
   End
   Begin VB.Image cierre 
      Height          =   480
      Left            =   14700
      Top             =   120
      Width           =   480
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   135
      Index           =   1
      Left            =   14535
      Top             =   7755
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   135
      Index           =   0
      Left            =   14535
      Top             =   8040
      Width           =   225
   End
   Begin VB.Image cmdInfo 
      Height          =   495
      Left            =   13680
      Top             =   7680
      Width           =   750
   End
   Begin VB.Image cmdLanzar 
      Height          =   495
      Left            =   11070
      Top             =   7680
      Width           =   2400
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   1
      Left            =   13560
      Top             =   10800
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   0
      Left            =   12360
      Top             =   10800
      Width           =   375
   End
   Begin VB.Image cmdCastillo 
      Height          =   495
      Left            =   14355
      Top             =   10740
      Width           =   615
   End
   Begin VB.Image cmdPestaña 
      Height          =   390
      Index           =   1
      Left            =   12960
      Top             =   1920
      Width           =   1980
   End
   Begin VB.Label lblAgilidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   14835
      TabIndex        =   21
      Top             =   10125
      Width           =   135
   End
   Begin VB.Label lblFuerza 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   13830
      TabIndex        =   20
      Top             =   10125
      Width           =   135
   End
   Begin VB.Image ImgResu 
      Height          =   360
      Left            =   11025
      Top             =   10830
      Width           =   360
   End
   Begin VB.Label lblMapCoord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000, 00, 00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8730
      TabIndex        =   19
      Top             =   2490
      Width           =   1815
   End
   Begin VB.Label LbLChat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1. Normal"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2370
      Width           =   1335
   End
   Begin VB.Image imgCanjes 
      Height          =   255
      Left            =   11340
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label DCLbL 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11670
      TabIndex        =   17
      Top             =   4920
      Width           =   735
   End
   Begin VB.Image imgMenu 
      Height          =   405
      Left            =   8715
      Top             =   10845
      Width           =   1350
   End
   Begin VB.Label lblItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(None)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10920
      TabIndex        =   15
      Top             =   5310
      Width           =   4050
   End
   Begin VB.Image DuelosSet 
      Height          =   375
      Left            =   12960
      Top             =   10800
      Width           =   375
   End
   Begin VB.Label FriendsCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12600
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Norte 
      Height          =   225
      Left            =   9420
      MouseIcon       =   "frmMainbackup.frx":27096
      Picture         =   "frmMainbackup.frx":27D60
      ToolTipText     =   "Castillo Norte atacado."
      Top             =   210
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Sur 
      Height          =   225
      Left            =   9420
      MouseIcon       =   "frmMainbackup.frx":27FFA
      Picture         =   "frmMainbackup.frx":28CC4
      ToolTipText     =   "Castillo Sur atacado."
      Top             =   1920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Oeste 
      Height          =   225
      Left            =   8610
      MouseIcon       =   "frmMainbackup.frx":28F5E
      Picture         =   "frmMainbackup.frx":29C28
      ToolTipText     =   "Castillo Oeste atacado."
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Este 
      Height          =   225
      Left            =   10320
      MouseIcon       =   "frmMainbackup.frx":29EC2
      Picture         =   "frmMainbackup.frx":2AB8C
      ToolTipText     =   "Castillo Este atacado."
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image cmdPestaña 
      Height          =   390
      Index           =   0
      Left            =   10980
      Top             =   1920
      Width           =   1980
   End
   Begin VB.Image BarMove 
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   15375
   End
   Begin VB.Image cmdMensaje 
      Height          =   450
      Left            =   8400
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000/000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11640
      TabIndex        =   12
      Top             =   10170
      Width           =   735
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000/000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   13980
      TabIndex        =   11
      Top             =   9285
      Width           =   735
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000/000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   13980
      TabIndex        =   10
      Top             =   8835
      Width           =   735
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000/000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11640
      TabIndex        =   9
      Top             =   9465
      Width           =   735
   End
   Begin VB.Image MANShp 
      Height          =   90
      Left            =   10920
      Top             =   9525
      Width           =   2100
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000/000"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11640
      TabIndex        =   8
      Top             =   8790
      Width           =   735
   End
   Begin VB.Image PicSeg 
      Height          =   375
      Left            =   11460
      Stretch         =   -1  'True
      Top             =   10830
      Width           =   375
   End
   Begin VB.Image cmdGold 
      Height          =   195
      Index           =   0
      Left            =   13140
      Top             =   4920
      Width           =   240
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   13440
      TabIndex        =   7
      Top             =   4920
      Width           =   105
   End
   Begin VB.Image Hpshp 
      Height          =   90
      Left            =   10920
      Top             =   8850
      Width           =   2100
   End
   Begin VB.Image STAShp 
      Height          =   90
      Left            =   10920
      Top             =   10230
      Width           =   2100
   End
   Begin VB.Image COMIDAsp 
      Height          =   75
      Left            =   13635
      Top             =   8880
      Width           =   1365
   End
   Begin VB.Image AGUAsp 
      Height          =   75
      Left            =   13635
      Top             =   9315
      Width           =   1365
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   11280
      TabIndex        =   6
      Top             =   735
      Width           =   315
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11880
      TabIndex        =   5
      Top             =   795
      Width           =   2265
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa Desconocido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8730
      TabIndex        =   4
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12720
      TabIndex        =   3
      Top             =   1170
      Width           =   300
   End
   Begin VB.Image ExpShp 
      Height          =   135
      Left            =   11835
      Top             =   1200
      Width           =   2115
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private last_i As Long
Public UsandoDrag As Boolean
Public UsabaDrag As Boolean

Public WithEvents dragInventory As clsGraphicalInventory
Attribute dragInventory.VB_VarHelpID = -1
Public WithEvents dragSpells As clsGraphicalSpells
Attribute dragSpells.VB_VarHelpID = -1

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long
Private Minuto As Byte

Public IsPlaying As Byte

Dim PuedeMacrear As Boolean

Private m_Jpeg As clsJpeg
Private m_FileName As String

Private Sub BarMove_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = vbLeftButton) Then
        Call Auto_Drag(Me.hwnd)
        Inventario.DrawInv
        Engine.DrawSpells
    End If
End Sub

Private Sub cierre_Click()
    Call CloseClient
End Sub

Private Sub cmdGold_Click(Index As Integer)
    Inventario.SelectGold
    If UserGLD > 0 Then
        frmCantidad.Show , frmMain
    End If
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub

Private Sub cmdLanzar_Click()
    If Opciones.HechizosClasicos = False Then Exit Sub
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.magia)
        End If
    End If
End Sub

Private Sub cmdMensaje_Click()
    frmMensaje.PopupMenuMensaje
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
    
        Select Case Index
            Case 1 'subir
                If hlst.ListIndex = 0 Then Exit Sub
            Case 0 'bajar
                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select
    
        Call WriteMoveSpell(Index = 1, hlst.ListIndex + 1)
        
        Select Case Index
            Case 1 'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1
            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
        End Select
    End If
End Sub

Private Sub cmdpestaña_Click(Index As Integer)
Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
        Case 0
            picInv.Visible = True
            GldLbl.Visible = True
            cmdGold(0).Visible = True
            DCLbL.Visible = True
            imgCanjes.Visible = True
            Inventario.DrawInv
        Case 1
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Lo sentimos, este sistema no esta operativo en estos momentos.", .red, .green, .blue, .bold, .italic)
            End With
    End Select

End Sub

Private Sub DuelosSet_Click()

    LlegoRank = False
    Call WriteSolicitarRank
    Call FlushBuffer
            
    Do While Not LlegoRank
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmBatalla.Iniciar_Labels
    frmBatalla.Show , frmMain
    LlegoRank = False
            
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyFisico = False Then Exit Sub

    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If Not SendTxt.Visible Then
                If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
                  (Not frmBancoObj.Visible) And _
                  (Not frmMSG.Visible) And _
                  (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                    Call CompletarEnvioMensajes
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                End If
            Else
                Call Enviar_SendTxt
            End If
    End Select

    If SendTxt.Visible Then Exit Sub
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                If Opciones.sMusica = CONST_MP3 Then
                    Sound.Music_Stop
                    Opciones.sMusica = CONST_DESHABILITADA
                Else
                    Opciones.sMusica = CONST_MP3
                End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Opciones.NamePlayers = Not Opciones.NamePlayers
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
                    FPSFLAG = Not FPSFLAG
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
                    Call frmMain.Client_Screenshot(frmMain.hDC, 1024, 768)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                    
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
                    If Shift <> 0 Then Exit Sub
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                    If Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        If UserDescansar Then
                            Call ShowConsoleMsg("¡Estás descansando!", .red, .green, .blue, .bold, .italic)
                            Exit Sub
                        ElseIf UserMeditar Then
                            Call ShowConsoleMsg("¡Estás meditando!", .red, .green, .blue, .bold, .italic)
                            Exit Sub
                        End If
                    End With
                    
                    Call WriteAttack
                    
                 Case CustomKeys.BindedKey(mKeyMapa)
                    'frmMapa.Show
            End Select
        End If
    
    
    Select Case KeyCode
        
        Case vbKey1
            SendingType = 1
            If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
            LbLChat.Caption = "1.Normal"
                
        Case vbKey2
            SendingType = 2
            If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
            LbLChat.Caption = "2.Gritar"
                
        Case vbKey3
            sndPrivateTo = InputBox("Nombre del destinatario:", vbNullString)
    
            If sndPrivateTo <> vbNullString Then
                SendingType = 3
                If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
            Else
                MsgBox "¡Escribe un nombre."
            End If
            LbLChat.Caption = "3.Privado"
                
        Case vbKey4
            SendingType = 4
            If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
            LbLChat.Caption = "4.Clan"
                
        Case vbKey5
            SendingType = 5
            If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
            LbLChat.Caption = "5.Party"
                
        Case vbKey6
            SendingType = 6
            If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
            LbLChat.Caption = "6.Global"
                
        Case vbKeyF1:
            Call CustomKeys.DoAccionTecla("F1")
                
        Case vbKeyF2:
            Call CustomKeys.DoAccionTecla("F2")

        Case vbKeyF3:
            Call CustomKeys.DoAccionTecla("F3")

        Case vbKeyF4:
            Call CustomKeys.DoAccionTecla("F4")
                
        Case vbKeyF5:
            Call CustomKeys.DoAccionTecla("F5")

        Case vbKeyF6:
            Call CustomKeys.DoAccionTecla("F6")
                
        Case vbKeyF7:
            Call CustomKeys.DoAccionTecla("F7")
                
        Case vbKeyF8:
            Call CustomKeys.DoAccionTecla("F8")
                
        Case vbKeyF9:
            Call CustomKeys.DoAccionTecla("F9")
                
        Case vbKeyF10:
            Call CustomKeys.DoAccionTecla("F10")
                
        Case vbKeyF11:
            Call CustomKeys.DoAccionTecla("F11")
                
        Case vbKeyF12:
            Call CustomKeys.DoAccionTecla("F12")

        Case vbKeyEscape
            fMenu.Visible = Not fMenu.Visible
    
        Case vbKeyZ
            RecTxt.Text = vbNullString
    End Select
     KeyFisico = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    clicX = x
    clicY = y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ChangeCursorMain(cur_Normal, frmMain)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub cmdCastillo_Click()
    Call WriteMotdCastillos
End Sub

Private Sub Image3_Click(Index As Integer)
    Call AddtoRichTextBox(frmMain.RecTxt, "Esta función no esta disponible en estos momentos.", 204, 193, 155, 0, 1)
End Sub

Private Sub imgCanjes_Click()
    Call WritePremios(0)
End Sub

Private Sub imgMenu_Click()
fMenu.Visible = Not fMenu.Visible
End Sub

Private Sub ImgResu_Click()
    Call WriteResuscitationToggle
End Sub

Private Sub lblPorcLvl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     MouseX = x
     MouseY = y
     
    lblexpactivo = True
        
    Call LabelExperiencia

End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub lvButtons_H1_Click()
    Call ParseUserCommand("/BUG")
End Sub

Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then Call ParseUserCommand("/TELEP YO " & UserMap & " " & CByte(x) & " " & CByte(y))
End Sub

Private Sub PicSeg_Click()
    Call WriteSafeToggle
End Sub

Private Sub picSpell_DblClick()
    UsandoDrag = False
    Engine.DrawSpells
End Sub

Private Sub picSpell_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    UsaMacro = False
    CnTd = 0

    If Not UsandoDrag Then
        If Button = vbRightButton Then
          
            If Spells.SelectedItem = 0 Then Exit Sub
            
            If Spells.GrhIndex(Spells.SelectedItem) > 0 Then
            
                last_i = Spells.SelectedItem
                If last_i > 0 And last_i <= 30 Then
                
                    Dim i As Integer
                    Dim Data() As Byte
                    Dim Handle As Integer
                    Dim bmpData As StdPicture
                    Dim poss As Integer
                    
                    poss = BuscarI(Spells.GrhIndex(Spells.SelectedItem))
                    
                    If poss = 0 Then
                        i = GrhData(Spells.GrhIndex(Spells.SelectedItem)).FileNum
                        If Extract_File_Memory(Graphics, App.Path & "\Recursos\", CStr(GrhData(Spells.GrhIndex(Spells.SelectedItem)).FileNum & ".bmp"), Data()) Then
                            Set bmpData = ArrayToPicture(Data(), 0, UBound(Data) + 1) ' GSZAO
                            frmMain.ImageList1.ListImages.Add , CStr("g" & Spells.GrhIndex(Spells.SelectedItem)), Picture:=bmpData
                            poss = frmMain.ImageList1.ListImages.count
                            Set bmpData = Nothing
                        End If
                    End If
                    
                    UsandoDrag = True
                    If frmMain.ImageList1.ListImages.count <> 0 Then
                        Set picSpell.MouseIcon = frmMain.ImageList1.ListImages(poss).ExtractIcon
                    End If
                    frmMain.picSpell.MousePointer = vbCustom
                    Exit Sub
                    
                End If
            End If
        Else
            If CurrentCursor <> cur_Action Then
                Call ChangeCursorMain(cur_Normal, frmMain)
            End If
    End If
    End If
End Sub

Private Sub picSpell_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSpell.MousePointer = vbDefault
    
End Sub

Private Sub renderer_Click()
    Form_Click
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
    
    If Button = 2 Then
        If Not frmComerciar.Visible And Not frmBancoObj.Visible Then
            Call WriteDoubleClick(tX, tY)
        End If
    End If
End Sub

Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
    
    Dim selInvSlot      As Byte
    'Get new target positions
    ConvertCPtoTP x, y, tX, tY
        
    With MapData(tX, tY)
        If UsabaDrag = False Then
            If CurrentCursor <> cur_Action Then
                If .CharIndex <> 0 Then
                    If charlist(.CharIndex).invisible = False Then
                        If charlist(.CharIndex).bType = 0 Then ' NPC friendly
                            Call ChangeCursorMain(cur_Npc, frmMain)
                        ElseIf charlist(.CharIndex).bType = 1 Then ' NPC hostile
                            Call ChangeCursorMain(cur_Npc_Hostile, frmMain)
                        ElseIf charlist(.CharIndex).bType = 2 Then ' User
                            If MapDat.battle_mode = False Then
                                Call ChangeCursorMain(cur_User, frmMain)
                            Else
                                Call ChangeCursorMain(cur_User_Danger, frmMain)
                            End If
                        End If
                    Else
                        Call ChangeCursorMain(cur_Normal, frmMain)
                    End If
                ElseIf .ObjGrh.GrhIndex <> 0 Then
                    Call ChangeCursorMain(cur_Obj, frmMain)
                Else
                    Call ChangeCursorMain(cur_Normal, frmMain)
                End If
            End If
        Else ' Utiliza Drag
            'Drag de items a posiciones. [maTih.-]
            
            'Get the selected slot of the inventory.
            selInvSlot = Inventario.SelectedItem
            
            'Not selected item?
            If Not selInvSlot <> 0 Then Exit Sub
            
            'There is invalid position?.
            If .blocked <> 0 Then
               Call ShowConsoleMsg("Posición inválida")
               Call StopDragInv
               Exit Sub
            End If
            
            ' Not Drop on ilegal position; Standelf
            Dim IS_VALID_POS As Boolean
            
            IS_VALID_POS = MoveToLegalPos(tX + 1, tY) = False And _
                            MoveToLegalPos(tX - 1, tY) = False And _
                            MoveToLegalPos(tX, tY - 1) = False And _
                            MoveToLegalPos(tX, tY + 1) = False
                
            If IS_VALID_POS Then
                Call ShowConsoleMsg("La posición donde desea tirar el ítem es ilegal.")
                Call StopDragInv
                Exit Sub
            End If
            
            'There is already an object in that position?.
            If Not .CharIndex <> 0 Then
                If .ObjGrh.GrhIndex <> 0 Then
                    Call ShowConsoleMsg("Hay un objeto en esa posición!")
                    Call StopDragInv
                    Exit Sub
                End If
            End If
            
            'Send the package.
            Call WriteDropObj(selInvSlot, tX, tY, 1)
            
            'Reset the flag.
            Call StopDragInv
        End If
    End With
End Sub

Private Sub StopDragInv()
' GSZAO
    UsabaDrag = False
    UsandoDrag = False
    If CurrentCursor <> cur_Action Then
        Call ChangeCursorMain(cur_Normal, frmMain)
        frmMain.picInv.MousePointer = vbNormal
    End If
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else
                If Inventario.amount(Inventario.SelectedItem) > 1 Then
                frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()
    If pausa Then Exit Sub
    
    If Inventario.ItemName(Inventario.SelectedItem) = "Amuleto Ankh" And UserEstado = 1 Then _
        If MsgBox("¿Quieres regresar al cementerio más cercano?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
    End If
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub Form_Click()

    If Cartel Then Cartel = False
    fMenu.Visible = False

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        If MouseBoton <> vbRightButton Then
        
            '¿Quiere solicitar party?
            If SolicitudParty = True Then
                Call ChangeCursorMain(cur_Normal, frmMain)
                Call WriteInvitacionAParty(tX, tY)
                SolicitudParty = False
                Exit Sub
            End If
        
            'Lorwik> AntiMacros !!
            If UsaMacro Then
                CnTd = CnTd + 1
                If CnTd = 5 Then
                    Call WriteUseSpellMacro
                    CnTd = 0
                End If
                UsaMacro = False
            End If
            '[/ybarra]
            If UsingSkill = 0 Then
                Call WriteLeftClick(tX, tY)
            Else
                
                If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                    Call ChangeCursorMain(cur_Normal, frmMain)
                    UsingSkill = 0
                    With FontTypes(FontTypeNames.FONTTYPE_TALK)
                        Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                'Splitted because VB isn't lazy!
                If UsingSkill = Proyectiles Then
                    If Not MainTimer.Check(TimersIndex.Arrows) Then
                        Call ChangeCursorMain(cur_Normal, frmMain)
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                End If
                
                'Splitted because VB isn't lazy!
                If UsingSkill = magia Then
                    If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                        If Not MainTimer.Check(TimersIndex.CastAttack) Then
                            If Opciones.BloqCruceta = 0 Then
                                Call ChangeCursorMain(cur_Normal, frmMain)
                                UsingSkill = 0
                            End If
                            Exit Sub
                        End If
                    Else
                        If Not MainTimer.Check(TimersIndex.CastSpell) Then
                            If Opciones.BloqCruceta = 0 Then
                                Call ChangeCursorMain(cur_Normal, frmMain)
                                UsingSkill = 0
                            End If
                            Exit Sub
                        End If
                    End If
                End If
                
                'Splitted because VB isn't lazy!
                If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                    If Not MainTimer.Check(TimersIndex.Work) Then
                        Call ChangeCursorMain(cur_Normal, frmMain)
                        UsingSkill = 0
                        Exit Sub
                    End If
                End If
                
                If CurrentCursor <> cur_Action Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                
                Call ChangeCursorMain(cur_Normal, frmMain)
                Call WriteWorkLeftClick(tX, tY, UsingSkill)
                UsingSkill = 0
            End If
        End If
    End If
    
End Sub

Private Sub dragSpells_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call WriteMoveItem(originalSlot, newSlot, eMoveType.SpellsI)
End Sub

Private Sub Form_Load()
  
    Me.Caption = Form_Caption
    
    'Lo desactivo en desarrollo por que el VB da error al pausar el proyecto y no quiero que se joda todo.
    #If Desarrollo = 2 Then
    If Opciones.URLCON = 1 Then Call StartURLDetect(RecTxt.hwnd, Me.hwnd)
    #End If
    
    Set dragInventory = Inventario
    Set dragSpells = Spells
    
    '****************************[MAIN]**********************************
    Me.Picture = General_Load_Picture_From_Resource("10.gif")
    'FondoCentro.Picture = LoadPicture(App.Path & "\Interfaces\inventario.jpg")
    
    '**********************[Barras de Stats]*****************************
    Hpshp.Picture = General_Load_Picture_From_Resource("12.gif")
    MANShp.Picture = General_Load_Picture_From_Resource("13.gif")
    STAShp.Picture = General_Load_Picture_From_Resource("14.gif")
    COMIDAsp.Picture = General_Load_Picture_From_Resource("27.gif")
    AGUAsp.Picture = General_Load_Picture_From_Resource("28.gif")
    ExpShp.Picture = General_Load_Picture_From_Resource("26.gif")
    ImgResu.Picture = General_Load_Picture_From_Resource("39.gif")
    '********************************************************************
    
    Call CargarCursoresMain
        
    If Opciones.HechizosClasicos Then
        hlst.Visible = True
        cmdLanzar.Visible = True
        cmdInfo.Visible = True
        picSpell.Visible = False
        cmdMoverHechi(1).Visible = True
        cmdMoverHechi(0).Visible = True
    Else
        hlst.Visible = False
        cmdLanzar.Visible = False
        cmdInfo.Visible = False
        picSpell.Visible = True
        cmdMoverHechi(1).Visible = False
        cmdMoverHechi(0).Visible = False
    End If
    
    Me.Left = 0
    Me.Top = 0
    SendingType = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopURLDetect
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
    
    lblexpactivo = False
    Call LabelExperiencia
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    Call UsarItem
    Call EquiparItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Sound.Sound_Play(SND_CLICK)
    Call ChangeCursorMain(cur_Normal, frmMain)
End Sub

Private Sub PicInv_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Inventario.DrawInv
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If Not UsandoDrag Then
        If Button = vbRightButton Then
          
            If Inventario.SelectedItem = 0 Then Exit Sub
            
            If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then
            
                last_i = Inventario.SelectedItem
                If last_i > 0 And last_i <= MAX_INVENTORY_SLOTS Then
                
                    Dim i As Integer
                    Dim Data() As Byte
                    Dim Handle As Integer
                    Dim bmpData As StdPicture
                    Dim poss As Integer
                    
                    poss = BuscarI(Inventario.GrhIndex(Inventario.SelectedItem))
                    
                    If poss = 0 Then
                        i = GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum
                        If Extract_File_Memory(Graphics, App.Path & "\Recursos\", CStr(GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum & ".bmp"), Data()) Then
                            Set bmpData = ArrayToPicture(Data(), 0, UBound(Data) + 1)
                            frmMain.ImageList1.ListImages.Add , CStr("g" & Inventario.GrhIndex(Inventario.SelectedItem)), Picture:=bmpData
                            poss = frmMain.ImageList1.ListImages.count
                            Set bmpData = Nothing
                            Erase Data
                        End If
                    End If
                    
                    UsandoDrag = True
                    If frmMain.ImageList1.ListImages.count <> 0 Then
                        Set picInv.MouseIcon = frmMain.ImageList1.ListImages(poss).ExtractIcon
                    End If
                    frmMain.picInv.MousePointer = vbCustom
                    Exit Sub
                    
                End If
            End If
        Else
            If CurrentCursor <> cur_Action Then
                Call ChangeCursorMain(cur_Normal, frmMain)
            End If
        End If
    End If
    
End Sub

Public Sub picSpell_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call WriteMoveItem(originalSlot, newSlot, eMoveType.SpellsI)
    frmMain.picSpell.MousePointer = vbNormal
End Sub

Public Sub picInv_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
    frmMain.picInv.MousePointer = vbNormal
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not Mod_General.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    Else
      If (Not frmComerciar.Visible) And (Not frmMSG.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) And (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    End If
        picSpell.SetFocus
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Or (KeyAscii = 126) Or (KeyAscii = 176) Then _
        KeyAscii = 0
End Sub

Private Sub CompletarEnvioMensajes()

Select Case SendingType
    Case 1
        SendTxt.Text = vbNullString
    Case 2
        SendTxt.Text = "-"
    Case 3
        SendTxt.Text = ("\" & sndPrivateTo & " ")
    Case 4
        SendTxt.Text = "/c "
    Case 5
        SendTxt.Text = "/p "
    Case 6
        SendTxt.Text = ";"
End Select

stxtbuffer = SendTxt.Text
SendTxt.SelStart = Len(SendTxt.Text)

End Sub

Private Sub Enviar_SendTxt()
    
    Dim str1 As String
    Dim str2 As String
    
    If Len(stxtbuffer) > 255 Then stxtbuffer = mid$(stxtbuffer, 1, 255)
    
    'Send text
    If Left$(stxtbuffer, 1) = "/" Then
        Call ParseUserCommand(stxtbuffer)

    'Shout
    ElseIf Left$(stxtbuffer, 1) = "-" Then
        If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> vbNullString Then Call ParseUserCommand(stxtbuffer)
        SendingType = 2
        
    'Global
    ElseIf Left$(stxtbuffer, 1) = ";" Then
        If LenB(Right$(stxtbuffer, Len(stxtbuffer) - 1)) > 0 And InStr(stxtbuffer, ">") = 0 Then Call ParseUserCommand(stxtbuffer)
        SendingType = 6

    'Privado
    ElseIf Left$(stxtbuffer, 1) = "\" Then
        str1 = Right$(stxtbuffer, Len(stxtbuffer) - 1)
        str2 = ReadField(1, str1, 32)
        If LenB(str1) > 0 And InStr(str1, ">") = 0 Then Call ParseUserCommand("\" & str1)
        sndPrivateTo = str2
        SendingType = 3
                
    'Say
    Else
        If LenB(stxtbuffer) > 0 Then Call ParseUserCommand(stxtbuffer)
        SendingType = 1
    End If

    stxtbuffer = vbNullString
    SendTxt.Text = vbNullString
    SendTxt.Visible = False
    
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        Call WriteLeftClick(tX, tY)
        
    Case 1 'Comerciar
        Call WriteLeftClick(tX, tY)
        Call WriteCommerceStart
    End Select
End Select
End Sub


Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
End Sub

Private Function BuscarI(gh As Integer) As Integer
Dim i As Integer
For i = 1 To frmMain.ImageList1.ListImages.count
    If frmMain.ImageList1.ListImages(i).Key = "g" & CStr(gh) Then
        BuscarI = i
        Exit For
    End If
Next i
End Function

Private Sub Timer1_Timer()
    Call FlushBuffer
End Sub

Private Sub TimerMin_Timer()
    Dim Segundo As Byte
    'Este timer cuenta 10 seg por tanto 10 x 6 = 1 minuto.
    
    Segundo = Segundo + 1
    
    'Si los segundos llego a 60, sumamos 1 minuto y restablecemos los segundos.
    If Segundo = 60 Then
        Minuto = Minuto + 1
        Segundo = 0
    End If
    
    #If Desarrollo = 0 Then
        Call ModSeguridad.BuscarEngine
        Call BuscarCheats
        
        '*******Anti Speed Hack*********
        If AntiSh() Then
            Call AntiShOn
            End
        End If
    #End If
    
End Sub

Private Sub UAOInformación_Click(Index As Integer)
    Call WriteSpellInfo(Spells.SelectedItem)
    FramHechi.Visible = False
    UsandoDrag = False
    Engine.DrawSpells
End Sub

Private Sub uAOMenu_Click(Index As Integer)
    Select Case Index
        Case 0
            Call WriteSolInfMonturaClient(0)
        Case 1
            If MsgBox("¿Estas seguro de que quieres salir del juego?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
            Call CloseClient
        Case 2
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            Call WriteRequestAtributes
            Call WriteRequestSkills
            Call WriteRequestMiniStats
            Call WriteRequestFame
            Call FlushBuffer
            
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            
       Case 3
            If frmGuildLeader.Visible Then Unload frmGuildLeader
            
            Call WriteRequestGuildLeaderInfo
            
       Case 4
            Call frmOpciones.Init
            
       Case 5
            frmTutorial.Show
            frmTutorial.DesdElMain = True
       Case 7
            If UserPartyId = 0 Then
                Call ChangeCursorMain(cur_Action, frmMain)
                Call AddtoRichTextBox(frmMain.RecTxt, "Haz click en el usuario con el que quieres hacer Party", 204, 193, 155, 0, 1)
                SolicitudParty = True
            Else
                Call ParseUserCommand("/PARTY")
            End If
            fMenu.Visible = Not fMenu.Visible
       Case 8
            fMenu.Visible = Not fMenu.Visible
       Case 9
            Call ParseUserCommand("/GM")
       Case 10
            Call ParseUserCommand("/SALIR")
       End Select
End Sub

'[Seguridad LwK - AntiMacros de palo]
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'********************************************u
'Lorwik> El codigo no es completamente mio :(
'********************************************

If Not GetAsyncKeyState(KeyCode) < 0 Then Exit Sub
KeyFisico = True
End Sub
'[/Seguridad LwK - AntiMacros de palo]

Public Sub Client_Screenshot(ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long)

On Error GoTo ErrorHandler

Dim i As Long
Dim Index As Long
i = 1

Set m_Jpeg = New clsJpeg

'80 Quality
m_Jpeg.Quality = 100

'Sample the cImage by hDC
m_Jpeg.SampleHDC hDC, Width, Height

m_FileName = App.Path & "\Fotos\AODrag_Foto"

If Dir(App.Path & "\Fotos", vbDirectory) = vbNullString Then
    MkDir (App.Path & "\Fotos")
End If

Do While Dir(m_FileName & Trim(str(i)) & ".jpg") <> vbNullString
    i = i + 1
    DoEvents
Loop

Index = i

m_Jpeg.Comment = "Character: " & UserName & " - " & Format(Date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm AM/PM")

'Save the JPG file
m_Jpeg.SaveFile m_FileName & Trim(str(Index)) & ".jpg"

Call AddtoRichTextBox(frmMain.RecTxt, "¡Captura realizada con exito! Se guardo en " & m_FileName & Trim(str(Index)) & ".jpg", 204, 193, 155, 0, 1)

Set m_Jpeg = Nothing

Exit Sub

ErrorHandler:
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Error en la captura!", 204, 193, 155, 0, 1)

End Sub

Private Sub UAOOlvidar_Click(Index As Integer)
    If MsgBox("¿Seguro que quierers olvidar " & UserHechizos(Spells.SpellSelectedItem).name & "?", vbExclamation + vbYesNo) = vbYes Then _
    Call WriteOlvidaSpell(Spells.SelectedItem)
    FramHechi.Visible = False
    UsandoDrag = False
    Engine.DrawSpells
End Sub

Private Sub CargarCursoresMain()
    Dim i As Byte
    For i = 0 To 1
        cmdMoverHechi(i).MousePointer = vbCustom
        cmdMoverHechi(i).MouseIcon = LoadPicture(App.Path & MouseHand)
        cmdPestaña(i).MousePointer = vbCustom
        cmdPestaña(i).MouseIcon = LoadPicture(App.Path & MouseHand)
        Image3(i).MousePointer = vbCustom
        Image3(i).MouseIcon = LoadPicture(App.Path & MouseHand)
    Next i
    'Cursores
    cmdLanzar.MousePointer = vbCustom
    cmdLanzar.MouseIcon = LoadPicture(App.Path & MouseHand)
    cmdInfo.MousePointer = vbCustom
    cmdInfo.MouseIcon = LoadPicture(App.Path & MouseHand)
    imgMenu.MousePointer = vbCustom
    imgMenu.MouseIcon = LoadPicture(App.Path & MouseHand)
    LbLChat.MousePointer = vbCustom
    LbLChat.MouseIcon = LoadPicture(App.Path & MouseHand)
    imgCanjes.MousePointer = vbCustom
    imgCanjes.MouseIcon = LoadPicture(App.Path & MouseHand)
    cmdGold(0).MousePointer = vbCustom
    cmdGold(0).MouseIcon = LoadPicture(App.Path & MouseHand)
    ImgResu.MousePointer = vbCustom
    ImgResu.MouseIcon = LoadPicture(App.Path & MouseHand)
    DuelosSet.MousePointer = vbCustom
    DuelosSet.MouseIcon = LoadPicture(App.Path & MouseHand)
    cmdCastillo.MousePointer = vbCustom
    cmdCastillo.MouseIcon = LoadPicture(App.Path & MouseHand)
    cmdMensaje.MousePointer = vbCustom
    cmdMensaje.MouseIcon = LoadPicture(App.Path & MouseHand)
End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"

    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If Not frmCrearPersonaje.Visible Then
        Call MostrarConnect
        frmConnect.Show vbModeless, frmRenderConnect
    End If
    
    Do While i < Forms.count - 1
        i = i + 1
        If Forms(i).name <> Me.name And Forms(i).name <> frmConnect.name And Forms(i).name <> frmCrearPersonaje.name And Forms(i).name <> frmMain.name And Forms(i).name <> frmRenderConnect.name And Forms(i).name <> frmCuenta.name Then
            Unload Forms(i)
        End If
    Loop
    On Local Error GoTo 0
    
    frmMain.Visible = False
    
    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserEmail = ""
    frmMain.picSpell.Picture = LoadPicture("")
    Call CerrarCuenta
    
    bTechoAB = 255
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    Alocados = 0

    frmMensaje.msg.Caption = "Se ha perdido la conexion con el servidor."
    frmMensaje.Show
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connect"
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)

    Select Case EstadoLogin
        Case E_MODO.Dados
            Unload frmConnect
            frmCrearPersonaje.Show vbModeless, frmRenderConnect
            
        Case Else
            Call Login
    End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim Data() As Byte
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD
    
    Data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
        
        If frmCuenta.Visible Then frmCuenta.Visible = False
        If frmMain.Visible Then
            frmMain.Visible = False
            frmMain.picSpell.Picture = LoadPicture("")
        End If
        
        Call CerrarCuenta
        
        frmRenderConnect.btnConsejo.Visible = False
        
        Call MostrarConnect
        
        frmConnect.Show vbModeless, frmRenderConnect
End Sub
