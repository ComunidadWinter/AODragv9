VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoundEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsSoundEngine - ImperiumAO - v1.4.5 R5
'
'Sound frmMain.Engine.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private Const MapExt As Byte = 0

Private dX As DirectX8

'**** Direct Sound ********
Private DS As DirectSound8
Private DS_Enum As DirectSoundEnum8

'**** Direct Music ********
Private dmLoader As DirectMusicLoader8
Private dmPerf As DirectMusicPerformance8
Private dmSeg As DirectMusicSegment8
Private dmSegState As DirectMusicSegmentState8
Private dmPath As DirectMusicAudioPath8

'**** MP3 ********
Dim IMC   As IMediaControl
Dim IBA   As IBasicAudio
Dim IME   As IMediaEvent
Dim IMPos As IMediaPosition

'**** AMBIENT ********
Dim ambient_buffer As DirectSoundSecondaryBuffer8
Dim ambient_rain_buffer As DirectSoundSecondaryBuffer8
Dim ambient_rain_int_buffer As DirectSoundSecondaryBuffer8
Dim ambient_fire_buffer As DirectSoundSecondaryBuffer8

Private mAmbienteActual As Integer
Private mMusActual As String
Private WavPath As String
Private MidiPath As String
Private MP3Path As String

'¿Está el engine iniciado?
Private Engine_Initialized As Boolean

'Manejador de buffers
Private Buffers As clsBufferMan
Private Buffers_Sec As clsBufferMan
Private Buffers_Ter As clsBufferMan

'El volúmen actual real (sonidos)
Private mEffectsVolume As Long

'El volúmen actual real (ambientes)
Private mEffectsVolumeAmbient As Long

'El volúmen actual real (música)
Private mMusicVolume As Long

'El volúmen máximo en la configuración (sonidos)
Private mEffectsVolumeMax As Long

'El volúmen máximo en la configuración (ambientes)
Private mAmbientVolumeMax As Long

'El volúmen máximo en la configuración (música)
Private mMusicVolumeMax As Long

'¿Estamos haciendo efecto fade?
Private mFadingStatus As Byte
Private mFadingMusicMod As Long

'¿Próximo MP3 o MIDI?
Private mNextMusic As String
Private mLastMusic As String

'Playing flags
Private play_fire As Boolean

'Midi de mapa
Private mLastMapMusic As Integer

Private INTERIOR_VOLUME As Long
Private EXTERIOR_VOLUME As Long

Private TimeElapsed As Long
Private mInvert_Sound As Boolean

Private Const VOLUME_DISTANCE_MOD As Long = 24
Private Const PAN_MOD As Long = 600

Private lastMid As Boolean

Private Property Let SetWavPath(ByVal inString As String)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

WavPath = inString

End Property

Private Property Let SetMP3Path(ByVal inString As String)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

MP3Path = inString
End Property

Private Property Let SetMidiPath(ByVal inString As String)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

MidiPath = inString
End Property

Public Property Get Engine_Running() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

Engine_Running = Engine_Initialized

End Property

Public Function Engine_DeInitialize()
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

'Destroy all buffers
If Not Buffers Is Nothing Then
    Call Buffers.BorrarTodo
    Set Buffers = Nothing
End If

If Not Buffers_Sec Is Nothing Then
    Call Buffers_Sec.BorrarTodo
    Set Buffers_Sec = Nothing
End If

If Not Buffers_Ter Is Nothing Then
    Call Buffers_Ter.BorrarTodo
    Set Buffers_Ter = Nothing
End If

Call Music_Empty

Set dmSegState = Nothing

If Not dmPath Is Nothing Then
    Call dmPath.Activate(False)
    Set dmPath = Nothing
End If

If Not dmPerf Is Nothing Then
    dmPerf.CloseDown
    Set dmPerf = Nothing
End If

Set dmLoader = Nothing
Set dmSeg = Nothing

Set DS_Enum = Nothing
Set DS = Nothing
Set dX = Nothing

End Function

Public Function Initialize_Engine(ByVal main_hwnd As String, ByVal Set_Wav_Path As String, ByVal Set_MP3_Path As String, ByVal Set_Midi_Path As String, Optional ByVal Play_only_on_focus As Boolean = True, Optional ByVal Initialize_Sounds As Boolean = True, Optional ByVal Initialize_Music As Boolean = True, Optional ByVal sound_volume As Long, Optional ByVal music_volume As Long, Optional ByVal invert_snd As Boolean) As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

On Error GoTo ErrHandler

Dim dmA As DMUS_AUDIOPARAMS

If Initialize_Sounds Or Initialize_Music Then
    Set dX = New DirectX8

    If Initialize_Sounds Then
        SetWavPath = Set_Wav_Path
        Set DS_Enum = dX.GetDSEnum
        Set DS = dX.DirectSoundCreate(DS_Enum.GetGuid(1))
        
        Set Buffers = New clsBufferMan
        Set Buffers_Sec = New clsBufferMan
        Set Buffers_Ter = New clsBufferMan
        
        Buffers.Path = Set_Wav_Path
        Buffers_Sec.Path = Set_Wav_Path
        Buffers_Ter.Path = Set_Wav_Path
        
        Call Buffers.Init(DS, 400)
        Call Buffers_Sec.Init(DS, 200)
        Call Buffers_Ter.Init(DS, 100)
        
        If Play_only_on_focus Then
            DS.SetCooperativeLevel main_hwnd, DSSCL_NORMAL
        Else
            DS.SetCooperativeLevel main_hwnd, DSSCL_PRIORITY
        End If
    
        VolumenActual = sound_volume
        mInvert_Sound = invert_snd
        
    End If
    
    Sound.VolumenActualAmbient_set Opciones.AmbientVol
    
    If Initialize_Music Then
        SetMidiPath = Set_Midi_Path
        SetMP3Path = Set_MP3_Path
        VolumenActualMusicMax = music_volume
        Set dmLoader = dX.DirectMusicLoaderCreate
        Set dmPerf = dX.DirectMusicPerformanceCreate
        dmPerf.InitAudio main_hwnd, DMUS_AUDIOF_ENVIRON + DMUS_AUDIOF_BUFFERS + DMUS_AUDIOF_3D + DMUS_AUDIOF_EAX + DMUS_AUDIOF_DMOS, dmA
        dmPerf.SetMasterVolume (music_volume)
        Set dmPath = dmPerf.CreateStandardAudioPath(DMUS_APATH_DYNAMIC_3D, 64, True)
    End If
    
    Engine_Initialized = True
    Initialize_Engine = True
Else
    Engine_Initialized = False
    Initialize_Engine = True
End If

Call Ambient_General_Load(mEffectsVolumeAmbient)

Exit Function

ErrHandler:
    If Len(Trim$(Err.Description)) Then
        Engine_Initialized = False
        Initialize_Engine = False
    End If
    
End Function

Public Property Get AmbienteActual() As Integer
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

AmbienteActual = mAmbienteActual
End Property

Public Property Get MusicActual() As String
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

MusicActual = mMusActual
End Property

Public Property Let MusicActual(ByVal vNewValue As String)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

mMusActual = vNewValue
End Property

Public Property Let Fading(ByVal vNewValue As Long)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

mFadingMusicMod = vNewValue
TimeElapsed = GetTickCount
End Property

Public Property Get Fading() As Long
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

Fading = mFadingMusicMod
End Property

Public Property Get VolumenActual() As Long
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

VolumenActual = mEffectsVolumeMax
End Property

Public Property Get VolumenActualAmbient() As Long
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

VolumenActualAmbient = mEffectsVolumeAmbient
End Property

Public Sub VolumenActualAmbient_set(ByVal vNewValue As Long)
    mEffectsVolumeAmbient = vNewValue
End Sub

Public Property Let VolumenActual(ByVal vNewValue As Long)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

Dim sglVolume As Single

If vNewValue > 0 Then vNewValue = 0
If vNewValue < -4000 Then vNewValue = -4000

mEffectsVolumeMax = vNewValue
Buffers.ModificarVolumenTodo (vNewValue)
Buffers_Sec.ModificarVolumenTodo (vNewValue)
Buffers_Ter.ModificarVolumenTodo (vNewValue)

sglVolume = ((vNewValue - 1) * 100) / -4000

End Property


Public Property Get InvertirSonido() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

InvertirSonido = mInvert_Sound
End Property

Public Property Let InvertirSonido(ByVal vNewValue As Boolean)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

mInvert_Sound = vNewValue
End Property

Public Property Let VolumenActualMusicMax(ByVal vNewValue As Long)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

If vNewValue > 0 Then vNewValue = 0
If vNewValue < -4000 Then vNewValue = -4000

mMusicVolumeMax = vNewValue

End Property

Public Property Get VolumenActualMusicMax() As Long
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

VolumenActualMusicMax = mMusicVolumeMax
End Property

Public Property Get VolumenActualMusic() As Long
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

VolumenActualMusic = mMusicVolume
End Property

Public Property Get LastMapMusic() As Integer
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

LastMapMusic = mLastMapMusic
End Property

Public Property Let LastMapMusic(ByVal vNewValue As Integer)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

mLastMapMusic = vNewValue
End Property

Public Property Let VolumenActualMusic(ByVal vNewValue As Long)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

mMusicVolume = vNewValue
End Property

Public Property Let NextMusic(ByVal vNewValue As String)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

If mFadingStatus <> 2 And LenB(vNewValue) > 0 Then 'Si no estamos en fade up ya...
    mNextMusic = vNewValue
Else
    mNextMusic = vNewValue
    mFadingStatus = 0
End If

End Property

Public Function Sound_Stop_All() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

On Error GoTo Error_Handl

Buffers.DetenerTodo
Buffers_Sec.DetenerTodo
Buffers_Ter.DetenerTodo

'Ambient_Stop

Sound_Stop_All = True

Exit Function
    
Error_Handl:
    Sound_Stop_All = False

End Function

Public Function Sound_Play(ByVal Sound_File As Integer, Optional ByVal Do_Loop As Boolean = False, Optional ByVal Volume As Long = 0, Optional ByVal pan As Long = 0) As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

On Error GoTo Error_Handl
    
    If (Sound_File <= 0) Or (Opciones.Audio = 0) Or (Engine_Initialized = False) Then Exit Function
    
    If Volume = 0 Then Volume = mEffectsVolumeMax
    
    If Not Buffers.Reproducir(Sound_File, IIf(Do_Loop, DSBPLAY_LOOPING, DSBPLAY_DEFAULT), Volume, pan) Then _
        If Not Buffers_Sec.Reproducir(Sound_File, IIf(Do_Loop, DSBPLAY_LOOPING, DSBPLAY_DEFAULT), Volume, pan) Then _
            Call Buffers_Ter.Reproducir(Sound_File, IIf(Do_Loop, DSBPLAY_LOOPING, DSBPLAY_DEFAULT), Volume, pan)
    
    Sound_Play = True
    
    Exit Function
    
Error_Handl:
    Sound_Play = False
End Function

Public Function Sound_Stop(ByVal Sound_File As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

On Error GoTo Error_Handl
    
    If (Sound_File <= 0) Or (Opciones.Audio = 0) Or (Engine_Initialized = False) Then Exit Function
    
    Call Buffers.Detener(Sound_File)
    Call Buffers_Sec.Detener(Sound_File)
    Call Buffers_Ter.Detener(Sound_File)
    
    Sound_Stop = True
    
    Exit Function

Error_Handl:
    Sound_Stop = False

End Function

Public Function Music_Empty() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************
    
    On Error GoTo Error_Handl
    
    Dim ret As Long
    
    If (Opciones.sMusica = CONST_DESHABILITADA) Or (Engine_Initialized = False) Then Exit Function
    
    If lastMid = False Then
        If ObjPtr(IMC) > 0 Then
            IMC.Stop
        End If
        
        Set IBA = Nothing
        Set IME = Nothing
        Set IMPos = Nothing
        Set IMC = Nothing
    Else
        If Not dmSeg Is Nothing Then
            dmPerf.StopEx dmSeg, 0, 0
            Set dmSeg = Nothing
        End If
    End If
    
    Music_Empty = True
    Exit Function

Error_Handl:
    Music_Empty = False
End Function

Public Function Music_Load(ByVal file_str As String, Optional ByVal Volume As Long = 0, Optional ByVal balance As Long = 0) As Boolean '**** Loads a MP3 *****
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************
    
    On Error GoTo Error_Handl
    
    Dim ret As Long
    
    If (LenB(file_str) <= 0) Or (Opciones.sMusica = CONST_DESHABILITADA) Or (Engine_Initialized = False) Then Exit Function

    If Opciones.sMusica = CONST_MP3 Then
        If Extract_File(Musica, MP3Path, file_str & ".mp3", Windows_Temp_Dir, False) Then
            If Not Music_Empty() = True Then GoTo Error_Handl
            Set IMC = New FilgraphManager
            IMC.RenderFile Windows_Temp_Dir & file_str & ".mp3"
            Kill Windows_Temp_Dir & file_str & ".mp3"
            lastMid = False
        'Lorwik> Deshabilito el midi.
        'ElseIf Extract_File(Midi, MidiPath, file_str & ".mid", Windows_Temp_Dir, False) Then
        '    If Not Music_Empty() = True Then GoTo Error_Handl
        '    Set dmSeg = dmLoader.LoadSegment(Windows_Temp_Dir & file_str & ".mid")
        '    dmSeg.SetStandardMidiFile
        '    dmSeg.Download dmPath
        '    lastMid = True
        Else
            GoTo Error_Handl
        End If
    'Lorwik> Deshabilito el midi.
    'ElseIf Opciones.sMusica = CONST_MIDI Then
    '    If Extract_File(Midi, MidiPath, file_str & ".mid", Windows_Temp_Dir, False) Then
    '        If Not Music_Empty() = True Then GoTo Error_Handl
    '        Set dmSeg = dmLoader.LoadSegment(Windows_Temp_Dir & file_str & ".mid")
    '        dmSeg.SetStandardMidiFile
    '        dmSeg.Download dmPath
    '        lastMid = True
    '    End If
    Else
        GoTo Error_Handl
    End If
    
    If Volume < -4000 Then Volume = -4000
    If balance > 5000 Then balance = 5000
    If balance < -5000 Then balance = -5000
    
    If lastMid = False Then
        Set IBA = IMC
        IBA.Volume = Volume
        IBA.balance = balance
        
        Set IME = IMC
        
        Set IMPos = IMC
        If ObjPtr(IMPos) Then IMPos.Rate = 1#
        IMPos.CurrentPosition = 0
    Else
        Call dmSeg.SetRepeats(-1)
    End If
    
    mMusActual = file_str
    Music_Load = True
    Exit Function

Error_Handl:
    Music_Load = False
End Function

Public Function Music_Play() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************
    
    On Error GoTo Error_Handl
    
    Dim ret As Long
    
    If (Opciones.sMusica = CONST_DESHABILITADA) Or (Engine_Initialized = False) Then Exit Function
    
    If lastMid = False Then
        IMC.Run
    Else
        Set dmSegState = dmPerf.PlaySegmentEx(dmSeg, 0, 0, Nothing, dmPath)
    End If
    
    Music_Play = True
    Exit Function

Error_Handl:
    Music_Play = False
End Function

Public Function Music_GetLoop() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************
    
    On Error GoTo Error_Handl
    
    If (Opciones.sMusica = CONST_DESHABILITADA) Or (Engine_Initialized = False) Then Exit Function
    
    If lastMid = False Then
        If IMPos Is Nothing Then Exit Function
        If IMPos.StopTime = IMPos.CurrentPosition Then
            Music_Stop
            Music_Play
            Music_GetLoop = True
        End If
    End If
    
    Exit Function
    
Error_Handl:
    Music_GetLoop = False
End Function

Public Function Music_Stop() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

    On Error GoTo Error_Handl
    
    Dim ret As Long
    
    If (Opciones.sMusica = CONST_DESHABILITADA) Or (Engine_Initialized = False) Then Exit Function
    
    If lastMid = False Then
        IMC.Stop
        IMPos.CurrentPosition = 0
    Else
        dmPerf.StopEx dmSeg, 0, 0
    End If
        
    Music_Stop = True
    Exit Function

Error_Handl:
    Music_Stop = False
End Function

Public Function Music_Pause() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

    On Error GoTo Error_Handl
    
    If (Opciones.sMusica = CONST_DESHABILITADA) Or (Engine_Initialized = False) Then Exit Function
    
    IMC.Pause
    
    Music_Pause = True
    Exit Function

Error_Handl:
    Music_Pause = False
End Function

Public Function Music_Volume_Set(ByVal sound_volume As Long) As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

    On Error GoTo Error_Handl
    
    If (Opciones.sMusica = CONST_DESHABILITADA) Or (Engine_Initialized = False) Then Exit Function
    If IBA Is Nothing Then Exit Function
    
    If sound_volume > 0 Then sound_volume = 0
    If sound_volume < -4000 Then sound_volume = -4000
    
    If lastMid = False Then
        IBA.Volume = sound_volume
    Else
        dmPerf.SetMasterVolume (sound_volume)
    End If

    Music_Volume_Set = True
    Exit Function

Error_Handl:
    Music_Volume_Set = False

End Function

Public Function Ambient_Load(ByVal file_name As Integer, Optional ByVal Volume As Long = 0)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************
    
    On Error GoTo Error_Handl

    If (file_name <= 0) Or (Opciones.Ambient = 0) Or (Engine_Initialized = False) Then Exit Function
    
    Call Buffers.GetBuffer(file_name, ambient_buffer, Ambient)
    
    If Volume < -4000 Then Volume = -4000
    ambient_buffer.SetVolume Volume
    
    Ambient_Load = True
    Exit Function

Error_Handl:
    Ambient_Load = False

End Function

Public Function Ambient_General_Load(ByVal Volume As Long)
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************
    
    On Error GoTo Error_Handl

    If (Opciones.Ambient = 0) Or (Engine_Initialized = False) Then Exit Function
    
    'Lorwik> No hay lluvia de momento.
    'Call Buffers.GetBuffer(SND_LLUVIAIN, ambient_rain_int_buffer)
    'Call Buffers.GetBuffer(SND_LLUVIAOUT, ambient_rain_buffer)
    Call Buffers.GetBuffer(SND_FUEGO, ambient_fire_buffer, Ambient)
    
    If Volume < -4000 Then Volume = -4000
    'ambient_rain_buffer.SetVolume Volume
    'ambient_rain_int_buffer.SetVolume Volume
    
    Ambient_General_Load = True
    Exit Function

Error_Handl:
    Ambient_General_Load = False

End Function

Public Function Ambient_Play() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************
    
    On Error GoTo Error_Handl
    
    If (Opciones.Ambient = 0) Or (Engine_Initialized = False) Then Exit Function
    ambient_buffer.Play DSBPLAY_LOOPING
    
    Ambient_Play = True
    Exit Function

Error_Handl:
    Ambient_Play = False
End Function

Public Function Ambient_Rain_Play() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************
    
    On Error GoTo Error_Handl
    
    If (Opciones.Ambient = 0) Or (Engine_Initialized = False) Then Exit Function
    ambient_rain_buffer.SetVolume mEffectsVolumeAmbient
    ambient_rain_buffer.Play DSBPLAY_LOOPING
    
    Ambient_Rain_Play = True
    Exit Function

Error_Handl:
    Ambient_Rain_Play = False
End Function

Public Function Ambient_Rain_Int_Play() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************
    
    On Error GoTo Error_Handl
    
    If (Opciones.Ambient = 0) Or (Engine_Initialized = False) Then Exit Function
    ambient_rain_int_buffer.SetVolume mEffectsVolumeAmbient
    ambient_rain_int_buffer.Play DSBPLAY_LOOPING
    
    Ambient_Rain_Int_Play = True
    Exit Function

Error_Handl:
    Ambient_Rain_Int_Play = False
End Function

Public Function Ambient_Stop() As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

    On Error GoTo Error_Handl
    
    If (Opciones.Ambient = 0) Or (Engine_Initialized = False) Then Exit Function
    
    If Not ambient_buffer Is Nothing Then ambient_buffer.Stop
    'ambient_rain_int_buffer.Stop
    'ambient_rain_buffer.Stop
    'ambient_fire_buffer.Stop
    
    play_fire = False
    mAmbienteActual = 0
    
    Ambient_Stop = True
    Exit Function

Error_Handl:
    Ambient_Stop = False
End Function

Public Function Ambient_Volume_Set(ByVal sound_volume As Long) As Boolean
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

    On Error GoTo Error_Handl
    
    If (Opciones.Ambient = 0) Or (Engine_Initialized = False) Then Exit Function
    
    If sound_volume > 0 Then sound_volume = 0
    If sound_volume < -4000 Then sound_volume = -4000
    
    If Not ambient_buffer Is Nothing Then ambient_buffer.SetVolume sound_volume
    'ambient_rain_int_buffer.SetVolume sound_volume
    'ambient_rain_buffer.SetVolume sound_volume
    
    Ambient_Volume_Set = True
    Exit Function

Error_Handl:
    Ambient_Volume_Set = False
End Function

Public Function Sound_Render()
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 17/08/2005
'**************************************************************

On Error Resume Next

Dim TiempoActual As Long, X As Integer, Y As Integer, mAmb As Integer

Static under_roof As Boolean
Static Sound_Buffer As DirectSoundSecondaryBuffer8

If Engine_Initialized Then

    TiempoActual = GetTickCount
    
    EXTERIOR_VOLUME = Opciones.AmbientVol - 300
    INTERIOR_VOLUME = Opciones.AmbientVol - 1000

    If UserMap > 0 Then
        under_roof = bTecho
        If MapExt Then
            mAmb = Ambient_Calculate
                
            If mAmbienteActual <> mAmb Then
                Call Ambient_Stop
                mAmbienteActual = mAmb
                Call Ambient_Load(mAmbienteActual, mEffectsVolumeAmbient)
                Call Ambient_Play
            End If
                
            If under_roof Then 'En los interiores el sonido es más suave :)
                If mEffectsVolumeAmbient <> INTERIOR_VOLUME Then
                    Call Ambient_Volume_Set(INTERIOR_VOLUME)
                    mEffectsVolumeAmbient = INTERIOR_VOLUME
                End If
            Else
                If mEffectsVolumeAmbient <> EXTERIOR_VOLUME Then
                    Call Ambient_Volume_Set(EXTERIOR_VOLUME)
                    mEffectsVolumeAmbient = EXTERIOR_VOLUME
                End If
            End If
                
        Else 'CurrentUser.MapExt = 0
            
            If mEffectsVolumeAmbient <> INTERIOR_VOLUME Then
                Call Ambient_Volume_Set(EXTERIOR_VOLUME)
                mEffectsVolumeAmbient = EXTERIOR_VOLUME
            End If
            
            mAmb = Ambient_Calculate
            
            If mAmbienteActual <> mAmb Then
                Call Ambient_Stop
                mAmbienteActual = mAmb
                Call Ambient_Load(mAmbienteActual, mEffectsVolumeAmbient)
                Call Ambient_Play
            End If
            
        End If
    
        If Map_Item_Grh_In_Current_Area(GRH_FOGATA, X, Y) Then
            If Not play_fire Then
                ambient_fire_buffer.SetVolume Calculate_Volume(X, Y)
                ambient_fire_buffer.SetPan Calculate_Pan(X, Y)
                ambient_fire_buffer.Play DSBPLAY_LOOPING
                play_fire = True
            ElseIf charlist(UserCharIndex).moved Then
                ambient_fire_buffer.SetVolume Calculate_Volume(X, Y)
                ambient_fire_buffer.SetPan Calculate_Pan(X, Y)
            End If
        ElseIf play_fire Then
            play_fire = False
            ambient_fire_buffer.Stop
        End If
                
    End If
            
    If mFadingMusicMod <> 0 And Opciones.sMusica <> CONST_DESHABILITADA Then
        
        '¿Fade up or fade down? Primero fade down, después fade up
        If TimeElapsed + 60 < TiempoActual Then
            
            TimeElapsed = TiempoActual
            
            'Fade down
            If mMusicVolume > mMusicVolumeMax Then
                
                mMusicVolume = mMusicVolumeMax
                
                If mFadingStatus = 0 Then
                    mFadingStatus = 1
                    mFadingMusicMod = mFadingMusicMod * (-1)
                Else
                    'TERMINAMOS!!!
                    mFadingMusicMod = 0
                    mFadingStatus = 0
                End If
            'Fade up
            ElseIf mMusicVolume < -4000 Then
                mFadingStatus = 2
                mMusicVolume = -4000
                mFadingMusicMod = mFadingMusicMod * (-1)
            End If
            
            mMusicVolume = mMusicVolume + mFadingMusicMod
            
            If Val(mNextMusic) > 0 And mFadingStatus = 2 Then
                If Music_Load(mNextMusic, mMusicVolume) Then
                    'Music_Stop
                    Music_Play
                End If
                mNextMusic = 0
            Else
                Music_Volume_Set mMusicVolume
            End If
        End If
    Else
        mMusicVolume = mMusicVolumeMax
    End If

    If Opciones.sMusica <> CONST_DESHABILITADA Then Call Music_GetLoop
    If UserCharIndex > 0 Then charlist(UserCharIndex).moved = False
    
End If

End Function

Public Function Sound_Load(ByVal file_num As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/16/2005
'**************************************************************
    
    If (Opciones.Audio = 0) Or (Engine_Initialized = False) Then Exit Function
    
    If Buffers.CargarBuffer(file_num, Wav) Then
        Sound_Load = True
    End If

End Function

Public Function Calculate_Pan(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/16/2005
'**************************************************************
    
    Dim total_distance As Integer, position_sgn As Integer, curr_x As Integer, curr_y As Integer
    
    If Char_Pos_Get(UserCharIndex, curr_x, curr_y) Then
        total_distance = General_Distance_Get(map_x, map_y, curr_x, curr_y)
        
        If mInvert_Sound = False Then
            If map_x < curr_x Then
                position_sgn = -1
            Else
                position_sgn = 1
            End If
        Else
            If map_x > curr_x Then
                position_sgn = -1
            Else
                position_sgn = 1
            End If
        End If
        
        If (total_distance = 0) Or (map_x = curr_x) Then
            Calculate_Pan = 0
        ElseIf total_distance < 9 Then
            Calculate_Pan = position_sgn * (total_distance * PAN_MOD)
        Else
            Calculate_Pan = position_sgn * (PAN_MOD * 9)
        End If
    End If

End Function

Public Function Calculate_Volume(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/16/2005
'**************************************************************
    
    Dim total_distance As Integer, curr_x As Integer, curr_y As Integer
    
    If (Opciones.Audio = 0) Or (Engine_Initialized = False) Then Exit Function
    
    If Char_Pos_Get(UserCharIndex, curr_x, curr_y) Then
        total_distance = General_Distance_Get(map_x, map_y, curr_x, curr_y)
        
        If (total_distance = 0) Then
            Calculate_Volume = mEffectsVolumeMax
        ElseIf total_distance < 9 Then
            Calculate_Volume = mEffectsVolumeMax - (total_distance * VOLUME_DISTANCE_MOD)
        Else
            Calculate_Volume = mEffectsVolumeMax - (VOLUME_DISTANCE_MOD * 9)
        End If
    End If

    If Calculate_Volume < -4000 Then Calculate_Volume = -4000

End Function

Public Sub BorraTimer()
'**************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 2/08/2006
'**************************************************************

If Buffers Is Nothing Then Exit Sub
Buffers.BorraTimerProc

If Buffers_Sec Is Nothing Then Exit Sub
Buffers_Sec.BorraTimerProc

If Buffers_Ter Is Nothing Then Exit Sub
Buffers_Ter.BorraTimerProc

End Sub

Private Function Ambient_Calculate() As Integer

Dim intSound As Integer

If MapInfo.Ambient = "" Then MapInfo.Ambient = 0

intSound = MapInfo.Ambient

'If MapExt > 0 And intSound <> 73 Then
'    If Val(MapDat.battle_mode) = 0 Then
'        Ambient_Calculate = SND_AMBIENTE_NOCHE
'    Else
'        Ambient_Calculate = SND_AMBIENTE_NOCHE_CIU
'    End If
'Else
 '   If intSound = 73 Then intSound = 75
    Ambient_Calculate = intSound
'End If

End Function


