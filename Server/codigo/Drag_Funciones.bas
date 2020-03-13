Attribute VB_Name = "Drag_Funciones"
Public Function in_collection(ByRef col As Collection, ByVal Key) As Boolean
    On Error GoTo KeyError
    If Not col(Key) Is Nothing Then
        in_collection = True
    Else
        in_collection = False
    End If

    Exit Function
KeyError:
    Err.Clear
    in_collection = False
End Function

Public Sub remove_from_collection(col As Collection, value As Variant)
  Dim i As Integer
  Dim count As Integer
  count = col.count
  
  
  For i = 1 To count
    If col(i) = value Then
      col.Remove i
    End If
  Next i
End Sub
