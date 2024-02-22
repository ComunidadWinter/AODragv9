Attribute VB_Name = "Mod_Blacklist"
Public DirLista As String
Public NNegro As String
Public nx As Integer
Public nxstring As String

Function AddNombre(NickNegro As String)
If FileExist(App.Path & "\ListaNegra\" & NickNegro & ".feo", vbArchive) Then
Call SendData(ToGM, 0, 0, "||Ya estaba agregado." & "´" & FontTypeNames.FONTTYPE_info)
Else
Call WriteVar(App.Path & "\ListaNegra\" & NickNegro & ".feo", "", "", NickNegro)
Call SendData(ToGM, 0, 0, "||Agregado a la Lista Negra " & NickNegro & "´" & FontTypeNames.FONTTYPE_info)
End If
End Function

Function QuitNombre(NickNegro As String)
If FileExist(App.Path & "\ListaNegra\" & NickNegro & ".feo", vbArchive) Then
Kill (App.Path & "\ListaNegra\" & NickNegro & ".feo")
Call SendData(ToGM, 0, 0, "||Borrado de la Lista Negra " & NickNegro & "´" & FontTypeNames.FONTTYPE_info)
Else
Call SendData(ToGM, 0, 0, "||No existe en la Lista Negra " & NickNegro & "´" & FontTypeNames.FONTTYPE_info)
End If

End Function

Function ComprobarLista(NickNegro As String)
If FileExist(App.Path & "\ListaNegra\" & NickNegro & ".feo", vbArchive) Then
Call SendData(ToGM, 0, 0, "||Ha entrado el Sospechoso " & NickNegro & "´" & FontTypeNames.FONTTYPE_TALK)
End If
End Function



