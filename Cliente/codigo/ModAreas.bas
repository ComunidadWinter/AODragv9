Attribute VB_Name = "ModAreas"
Option Explicit

'Lorwik> Este sistema tendria que volar, me da asco.

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer

Private Const TamañoAreas As Byte = 11

Public Sub CambioDeArea(ByVal X As Byte, ByVal Y As Byte)
    Dim loopX As Long, loopY As Long
    
    MinLimiteX = (X \ TamañoAreas - 1) * TamañoAreas
    MaxLimiteX = MinLimiteX + ((TamañoAreas * 3) - 1)
    
    MinLimiteY = (Y \ TamañoAreas - 1) * TamañoAreas
    MaxLimiteY = MinLimiteY + ((TamañoAreas * 3) - 1)
    
    For loopX = 1 To XMaxMapSize
        For loopY = 1 To XMaxMapSize
            
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                'Erase NPCs
                
                If MapData(loopX, loopY).CharIndex > 0 Then
                    If MapData(loopX, loopY).CharIndex <> UserCharIndex Then
                        Call EraseChar(MapData(loopX, loopY).CharIndex)
                    End If
                End If
                
                'Erase OBJs
                MapData(loopX, loopY).ObjGrh.GrhIndex = 0
            End If
        Next loopY
    Next loopX
    
    Call RefreshAllChars
End Sub

