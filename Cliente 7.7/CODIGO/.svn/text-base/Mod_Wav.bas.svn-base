Attribute VB_Name = "Mod_Wav"
Option Explicit

Public Const SND_SYNC = &H0 ' SINCRONO
Public Const SND_ASYNC = &H1 ' ASINCRONO
Public Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Public Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10 ' don't stop any currently playing sound

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?WAVS¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public Const SND_CLICK = "click.Wav"

Public Const SND_PASOS1 = "23.Wav"
Public Const SND_PASOS2 = "24.Wav"
Public Const SND_NAVEGANDO = "50.wav"
Public Const SND_OVER = "click2.Wav"
Public Const SND_DICE = "cupdice.Wav"

Function LoadWavetoDSBuffer(DS As DirectSound, DSB As DirectSoundBuffer, sFile As String) As Boolean
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX

    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC

    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set DSB = DS.CreateSoundBufferFromFile(DirSound & sFile, bufferDesc, waveFormat)

    If Err.Number <> 0 Then
        Exit Function
    End If

    LoadWavetoDSBuffer = True
End Function

Sub audio.playwave(file As String)
 On Error GoTo fin
    If Fx = 1 Then Exit Sub
    If Not FileExist(DirSound & file, vbNormal) Then Exit Sub
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If

    If LoadWavetoDSBuffer(DirectSound, DSBuffers(LastSoundBufferUsed), file) Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If
fin:
End Sub
