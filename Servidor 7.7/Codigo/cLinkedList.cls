VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLinkedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Coded by Gulfas Morgolock
'morgolock@speedy.com.ar

Option Explicit

Private Const MAX = 300

Private Type tNodo
        Data As Object
        Next As Integer
End Type

Private m_linkedList(1 To MAX) As tNodo
Private m_size As Integer
Private m_head As Integer
Private m_tail As Integer

Public Sub ClearList()

Dim k As Integer

For k = 1 To MAX
    Set m_linkedList(k).Data = Nothing
    m_linkedList(k).Next = 0
Next k

m_size = 0
m_head = 0
m_tail = 0

End Sub

Public Function Size() As Integer
Size = m_size
End Function

Private Function GetFreeIndex() As Integer

Dim k As Integer

For k = 1 To MAX
    If m_linkedList(k).Data Is Nothing Then
                GetFreeIndex = k
                Exit Function
    End If
Next k

End Function

Public Sub Insert(ByRef obj As Object)

If Me.Size = 0 Then
    Set m_linkedList(1).Data = obj
    m_linkedList(1).Next = 0
    m_head = 1
    m_tail = 1
Else
    If m_size < MAX Then
        Dim j As Integer
        j = GetFreeIndex
        Set m_linkedList(j).Data = obj
        m_linkedList(j).Next = 0
        m_linkedList(m_tail).Next = j
        m_tail = j
    End If
End If

m_size = m_size + 1

End Sub

Public Function Includes(ByVal obj As Object) As Boolean

If m_size > 0 Then
    Dim k As Integer
    k = m_head
    Do
        If m_linkedList(k).Data.Equal(obj) Then
            Includes = True
            Exit Function
        End If
        
        k = m_linkedList(k).Next
        
    Loop While k <> 0
Else
    Includes = False
End If

End Function

Public Function GetObjByIndex(ByVal Index As Integer) As Object

If m_size > 0 And Index > 0 Then
    If Index = 1 Then
        Set GetObjByIndex = m_linkedList(m_head).Data
        Exit Function
    Else
        Dim k As Integer, ant As Integer, cont As Integer
        k = m_head
        cont = 0
        Do
            ant = k
            k = m_linkedList(k).Next
            cont = cont + 1
        Loop While k <> 0 And cont < Index
        Set GetObjByIndex = m_linkedList(ant).Data
    End If
End If



End Function

Public Function GetButonAt(ByVal X As Single, ByVal Y As Single) As Object

Dim j As Integer, MyObj As Object
For j = 1 To m_size
    Set MyObj = Me.GetObjByIndex(j)
    If MyObj.Clicked(X, Y) Then
            Set GetButonAt = MyObj
            Exit Function
    End If
Next j

Set GetButonAt = Nothing
    
End Function


Public Sub Remove(ByRef obj As Object)

If m_size > 0 Then
    Dim k As Integer, ant As Integer
    k = m_head
    ant = k
    Do
        If m_linkedList(k).Data.Equal(obj) Then
            If m_head = k Then m_head = m_linkedList(k).Next
            If m_tail = k Then m_tail = ant
            m_linkedList(ant).Next = m_linkedList(k).Next
            Set m_linkedList(k).Data = Nothing
            m_linkedList(k).Next = 0
            m_size = m_size - 1
            Exit Sub
        End If
        ant = k
        k = m_linkedList(k).Next
    Loop While k <> 0
End If

End Sub
