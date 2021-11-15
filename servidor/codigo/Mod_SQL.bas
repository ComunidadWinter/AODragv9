Attribute VB_Name = "Mod_SQL"
    Option Explicit
    Public Con As ADODB.Connection
    Public RS As ADODB.Recordset
    
    Public Type DBSQL
        Driver As String
        Server As String
        Database As String
        name As String
        Pass As String
        Modo As String
    End Type
    
    Public SQL As DBSQL
     
    Public Sub CargarDB()
    On Error GoTo ErrHandler
     
    Set Con = New ADODB.Connection
     
    Con.CursorLocation = adUseClient
    Con.ConnectionString = "DRIVER=" & SQL.Driver & ";" & "SERVER=" & SQL.Server & ";" & " DATABASE=" & SQL.Database & ";" & "UID=" & SQL.name & ";PWD=" & SQL.Pass & "; OPTION=" & SQL.Modo
    Con.Open
    Exit Sub
     
ErrHandler:
       MsgBox "Error en CargarDB: " & Err.Description & " String: " & Con.ConnectionString
       End
    End Sub
     
    Public Sub CerrarDB()
    On Error GoTo ErrHandle
     
    Con.Close
    Set Con = Nothing
     
    Exit Sub
     
ErrHandle:
        MsgBox "Ha surgido un error al cerrar la base de datos MySQL"
        End
       
    End Sub
     
    Public Function AsignarIndexPJ(ByVal UserIndex As Integer)
        Dim mUser As User
        mUser = UserList(UserIndex)
     
        If Len(mUser.name) = 0 Then Exit Function
         
        Set RS = New ADODB.Recordset
         
    Set RS = Con.Execute("SELECT * FROM `Amigos` WHERE Nombre='" & mUser.name & "'")
    If RS.BOF Or RS.EOF Then
        Con.Execute ("INSERT INTO `Amigos` (NOMBRE) VALUES ('" & mUser.name & "')")
        Set RS = Nothing
        Set RS = Con.Execute("SELECT * FROM `Amigos` WHERE Nombre='" & mUser.name & "'")
        UserList(UserIndex).IndexPJ = RS!IndexPJ
    Else
        Set RS = Con.Execute("SELECT * FROM `Amigos` WHERE IndexPJ=" & mUser.IndexPJ)
        UserList(UserIndex).IndexPJ = RS!IndexPJ
    End If
         
        Set RS = Nothing
    End Function
     
    Public Function IndexPJ(ByVal name As String) As Integer
     
    Set RS = Con.Execute("SELECT * FROM `Flags` WHERE Nombre='" & UCase$(name) & "'")
     
    If RS.BOF Or RS.EOF Then Exit Function
     
    IndexPJ = RS!IndexPJ
     
    Set RS = Nothing
     
    End Function

Public Function SQLConsulta(ByVal index As Long, ByVal Tabla As String, ByVal dato As String) As String
    Set RS = Con.Execute("SELECT * FROM `" & Tabla & "` WHERE IndexPJ=" & index)

    If RS.BOF Or RS.EOF Then Exit Function
     
    SQLConsulta = RS.Fields(dato)
     
    Set RS = Nothing
End Function

Public Function SQLSave(ByVal index As Long, ByVal Tabla As String, ByVal indice As String, ByVal dato As String)
    Dim str As String
    Set RS = Con.Execute("SELECT * FROM `" & Tabla & "` WHERE IndexPJ=" & index)
    If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `" & Tabla & "` (IndexPJ) VALUES (" & index & ")")
    Set RS = Nothing
     
    str = "UPDATE `" & Tabla & "` SET"
    str = str & " IndexPJ=" & index
    str = str & "," & indice & "=" & dato
    str = str & " WHERE IndexPJ=" & index
     
    Call Con.Execute(str)
End Function

Public Function NameToIndex(ByVal name As String) As String
    Set RS = Con.Execute("SELECT * FROM `amigos` WHERE Nombre='" & UCase$(name) & "'")

    If RS.BOF Or RS.EOF Then Exit Function
     
    NameToIndex = RS!IndexPJ
     
    Set RS = Nothing
End Function


Public Function SQLBORRAR(ByVal index As Long, ByVal Tabla As String) As String
    Set RS = Con.Execute("DELETE FROM `" & Tabla & "` WHERE IndexPJ=" & index)
     
    Set RS = Nothing
End Function
