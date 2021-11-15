Attribute VB_Name = "Drag_MySQL"
Option Explicit

Public SQL As ADODB.Connection
Public RS As ADODB.Recordset
Public RS2 As ADODB.Recordset


Public Sub MySQL_Connect()
On Error GoTo errHandler
  Set SQL = New ADODB.Connection
  #If Desarrollo = 0 Then
    SQL.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=aodrag.es;PORT=3306;DATABASE=aodrag9;UID=aodrag9;PWD=aodrag9;OPTION=3"
  #Else
    SQL.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=aodrag.es;PORT=3306;DATABASE=aodrag9;UID=aodrag9;PWD=aodrag9;OPTION=3"
  #End If
  SQL.CursorLocation = adUseClient
  SQL.Open
  
  Exit Sub
errHandler:
    Debug.Print "error en MySQL_Connect" & Err.Number & Err.Description
End Sub


Public Sub CheckSQL()
  If SQL Is Nothing Then MySQL_Connect
  If SQL.State <> 1 Then MySQL_Connect
End Sub





'02/11/2015 Irongete: Devuelve True o False dependiendo de si una cuenta existe en la base de datos
Function ExisteCuenta(ByRef nombre As String)
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id FROM cuenta WHERE nombre = '" & nombre & "'")
    If RS!id > 0 Then
        ExisteCuenta = True
    Else
        ExisteCuenta = False
    End If
End Function

'17/11/2015 Irongete: Devuelve hora en formato YYYY-MM-DD HH:MM:SS muy útil para introducirlo en sql como datetime
Function Ahora()
    Ahora = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & " " & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
End Function

'17/11/2015 Irongete: Devuelve el userindex del personaje segun su id
Function GetPersonajeIndex(ByRef PersonajeId As Integer) As Integer
    Set RS = SQL.Execute("SELECT userindex FROM personaje WHERE id = '" & PersonajeId & "' AND logged = '1'")
    If RS.RecordCount = 1 Then
        GetPersonajeIndex = RS!UserIndex
    Else
        GetPersonajeIndex = 0
    End If
End Function

'20/11/2015 Irongete: Devuelve la id del personaje segun su userindex
Function GetPersonajeId(ByRef PersonajeIndex As Integer) As Integer
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT id FROM personaje WHERE userindex = '" & PersonajeIndex & "' AND logged = '1'")
    If RS.RecordCount = 1 Then
        GetPersonajeId = RS!id
    Else
        GetPersonajeId = 0
    End If
End Function

Public Function GetIdPorNombre(ByVal id As Integer) As String
    Set RS = New ADODB.Recordset
    Set RS = SQL.Execute("SELECT nombre FROM personaje WHERE id = '" & id & "'")
    GetIdPorNombre = CStr(RS!nombre)
End Function
