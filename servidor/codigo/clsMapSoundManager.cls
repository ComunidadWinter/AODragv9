VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundMapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type p_tSoundMapInfo
    Cantidad As Integer
    SoundIndex() As Integer
    flags() As Long
    Probabilidad() As Single
End Type

Private p_Mapas() As p_tSoundMapInfo


'sonidos conocidos, pasados a enum para intelisense
Public Enum e_SoundIndex
    MUERTE_HOMBRE = 11
    MUERTE_MUJER = 109
    FLECHA_IMPACTO = 65
    CONVERSION_BARCO = 55
    MORFAR_MANZANA = 77
    SOUND_COMIDA = 7
    SOUND_DROP = 132
    SOUND_APUÑALA = 150
    DROPPROB = 132
End Enum

Private Sub Class_Initialize()
'armar el array
    ReDim p_Mapas(1 To NumMaps) As p_tSoundMapInfo
End Sub

Public Sub ReproducirSonido(ByVal Destino As SendTarget, ByVal Index As Integer, ByVal SoundIndex As Integer)
    Call SendData(Destino, Index, PrepareMessagePlayWave(SoundIndex, UserList(Index).Pos.X, UserList(Index).Pos.Y))
End Sub
