Attribute VB_Name = "ToolTip"
' Para colocar la ventana en primer plano
Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
    
' Constantes para el api SetWindowPos
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

'Recupera las corrdenadas del mouse
Private Declare Function GetCursorPos Lib "user32" ( _
    lpPoint As POINTAPI) As Long

'Estructura POINTAPI necesaria para el api GetCursorPos
Private Type POINTAPI
        x As Long
        y As Long
End Type

' Función que crea un timer
Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

' Función que detiene el timer iniciado
Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long

'Variables
Dim m_Form As Form ' para el form que hace de tooltiptext
Dim elLabel As Label ' para crear un label en tiempo de ejecución
Dim Control_Image As Image


' Crea y Muestra el ToolTipText
'*****************************************************
 Sub Mostrar_ToolTip(El_Form As Form, _
                     texto As String, _
                     Color_De_Fondo As Long, _
                     color_Texto As Long, _
                     Optional Path_Imagen As String)
    
    ' Elimina el toolTip por si estaba cargado
    Call Eliminar_ToolTip
      
    ' Referencia el formulario que actúa como _
      ToolTip a una variable local de tipo Form
    Set m_Form = El_Form
    
    ' Crea un label en tiempo de ejecución ( para el texto del tooltip )
    Set elLabel = m_Form.Controls.Add("vb.Label", "lb1")
    
    ' Propiedades para el label
    
    ' si se pasó el parámetro de la imagen ...
    If Len(Path_Imagen) Then
        ' Crea un image en tiempod eejecución
        Set Control_Image = m_Form.Controls.Add("vb.image", "Img1")
        'Carga la imagen
        Control_Image.Picture = LoadPicture(Path_Imagen)
        Control_Image.Move 15, 15
        Control_Image.Visible = True
    End If
        
        
    ' propiedades del label
    With elLabel
        .Caption = texto
        .BackStyle = 0
        '.FontBold = True
        .AutoSize = True
        .ForeColor = color_Texto
        If Len(Path_Imagen) Then
            .Left = 100 + Control_Image.Width
        Else
            .Left = 100
        End If
        .Top = 100
        .Visible = True
    End With
    
    ' Propiedades para el formulario
    With m_Form
        
        .BackColor = Color_De_Fondo
        
        If Len(Path_Imagen) Then
            ' Ancho y alto
            .Width = elLabel.Width + 250 + Control_Image.Width
            .Height = elLabel.Height - 250 + Control_Image.Height
        Else

            .Width = elLabel.Width + 250
            .Height = elLabel.Height - 250
        End If
         ' para que al repintar mantenga el borde
        .AutoRedraw = True
    End With
    
    ' Dibuja el borde al formulario
    
    ' Linea blanca superior
    m_Form.Line (0, 0)-(m_Form.ScaleWidth, 0), &H80000009, B
    m_Form.Line (0, 0)-(0, m_Form.ScaleHeight), &H80000009, B
    
    ' Linea negra inferior
    m_Form.Line (0, m_Form.ScaleHeight - 10)- _
                (m_Form.ScaleWidth, m_Form.ScaleHeight - 10), vbBlack, B
    
    m_Form.Line (m_Form.ScaleWidth - 10, 0)- _
                (m_Form.ScaleWidth - 10, m_Form.ScaleHeight), vbBlack, B
    
    
    
    ' Activa el timer en un milisegundo
    SetTimer m_Form.hwnd, 0, 1, AddressOf TimerProc
    
    
    
End Sub


' Procedimiento del Timer
Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, _
              ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    
    Dim mouse As POINTAPI
    
    ' REcupera la posición del mouse
    GetCursorPos mouse
    
    ' Posiciona el formulario en las coordenadas del cursor
    m_Form.Left = (mouse.x * Screen.TwipsPerPixelX) - 1800
    m_Form.Top = (mouse.y * Screen.TwipsPerPixelY) + 100
    
    
    If Not m_Form.Visible Then
        ' Coloca el formulario Alwaysontop
        SetWindowPos m_Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
                     SWP_NOMOVE + SWP_NOSIZE
                             
        m_Form.Visible = True
                     
    End If

End Sub


 ' Remueve el ToolTip
 Sub Eliminar_ToolTip()
     
     ' Por si se intenta ejecutar la Sub, y el form no está referenciado
     If m_Form Is Nothing Then
        Exit Sub
     End If
      
     ' Finaliza el timer
     KillTimer m_Form.hwnd, 0
     
     ' DEscarga el formulario
     Unload m_Form
     Set m_Form = Nothing
 
 End Sub

