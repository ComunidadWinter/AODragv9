Attribute VB_Name = "Drag_Clanes"
'27/02/2016 Irongete: Mensaje a todos los miembros del clan
Public Sub BroadcastClan(ByVal clanid As Integer, ByVal Mensaje As String)
    Set RS = New ADODB.Recordset
    If clanid > 0 Then
        '27/02/2016 Irongete: Envío el mensaje a todos los miembros del clan que están logeados
        Set RS = SQL.Execute("SELECT userindex FROM personaje WHERE id_clan = '" & clanid & "' AND logged = '1'")
        While Not RS.EOF
            Call WriteConsoleMsg(RS!UserIndex, Mensaje, FontTypeNames.FONTTYPE_GUILDMSG)
            RS.MoveNext
        Wend
    End If
    RS.Close
End Sub
