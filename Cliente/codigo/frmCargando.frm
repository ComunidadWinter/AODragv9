VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "AODrag"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
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
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando..."
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
      Left            =   1920
      TabIndex        =   0
      Top             =   8250
      Width           =   8055
   End
   Begin VB.Image imgProgress 
      Height          =   390
      Left            =   1635
      Picture         =   "frmCargando.frx":1CCA
      Top             =   8160
      Width           =   8700
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private porcentajeActual As Integer
 
Private Const PROGRESS_DELAY = 10
Private Const PROGRESS_DELAY_BACKWARDS = 4
Private Const DEFAULT_PROGRESS_WIDTH = 608
Private Const DEFAULT_STEP_FORWARD = 1
Private Const DEFAULT_STEP_BACKWARDS = -3
 
Public Sub progresoConDelay(ByVal porcentaje As Integer)
 
If porcentaje = porcentajeActual Then Exit Sub
 
Dim step As Integer, stepInterval As Integer, Timer As Long, tickCount As Long
 
If (porcentaje > porcentajeActual) Then
    step = DEFAULT_STEP_FORWARD
    stepInterval = PROGRESS_DELAY
Else
    step = DEFAULT_STEP_BACKWARDS
    stepInterval = PROGRESS_DELAY_BACKWARDS
End If
 
Do Until compararPorcentaje(porcentaje, porcentajeActual, step)
    Do Until (Timer + stepInterval) <= GetTickCount()
        DoEvents
    Loop
    Timer = GetTickCount()
    porcentajeActual = porcentajeActual + step
    Call establecerProgreso(porcentajeActual)
Loop
 
End Sub
 
 
Public Sub establecerProgreso(ByVal nuevoPorcentaje As Integer)
 
If nuevoPorcentaje >= 0 And nuevoPorcentaje <= 100 Then
    imgProgress.Width = DEFAULT_PROGRESS_WIDTH * CLng(nuevoPorcentaje) / 100
ElseIf nuevoPorcentaje > 100 Then
    imgProgress.Width = DEFAULT_PROGRESS_WIDTH
Else
    imgProgress.Width = 0
End If
porcentajeActual = nuevoPorcentaje
 
End Sub
 
Private Function compararPorcentaje(ByVal porcentajeTarget As Integer, ByVal porcentajeAct As Integer, ByVal step As Integer) As Boolean
 
If step = DEFAULT_STEP_FORWARD Then
    compararPorcentaje = (porcentajeAct >= porcentajeTarget)
Else
    compararPorcentaje = (porcentajeAct <= porcentajeTarget)
End If
 
End Function
Private Sub Form_Load()
Me.Caption = Form_Caption
Me.Picture = General_Load_Picture_From_Resource("11.gif")
imgProgress.Picture = General_Load_Picture_From_Resource("38.gif")
End Sub
