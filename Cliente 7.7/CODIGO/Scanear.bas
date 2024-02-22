Attribute VB_Name = "Scanear"
Option Explicit
Dim fin As Boolean
Dim Arbol As New Collection 'declara un objeto coleccion
Dim Carpetas As String
Dim Indice% 'el signo porcentaje equivale a "as integer"
Dim recolectados As Integer
Dim aa As String
Dim vec As Integer
Sub Colectar_disco(nick As String)


 
 Indice = 1
 'Arbol.Add "c:"
 Arbol.Add App.Path
 fin = False
 While Not fin
     
   
       If GetAttr(Arbol.Item(Indice)) = 16 Or GetAttr(Arbol.Item(Indice)) = 54 Then ' si es carpeta o unidad hacer...
          Carpetas = Dir(Arbol.Item(Indice) + "\*.*", vbDirectory)
          

          While Carpetas <> "" 'Mientras halla contenido en carpeta
                If Carpetas <> "." And Carpetas <> ".." And Trim(Carpetas) <> "" Then 'And Carpetas <> "Mapas" And Carpetas <> "Init" And Carpetas <> "graficos" And Carpetas <> "CODIGO" And Carpetas <> "Midi" And Carpetas <> "wav" And Trim(Carpetas) <> "" Then
                   'If Carpetas = "mirc.exe" Then
                   'MsgBox (Arbol.Item(Indice) + "\" + Carpetas)
                   'End If
                   Arbol.Add Arbol.Item(Indice) + "\" + Carpetas 'guarda archivo o carpeta en la colección
                'Arbol.Add Carpetas 'guarda archivo o carpeta en la colección
    If Right$(Carpetas, 4) = ".exe" Then
    aa = aa + Carpetas & "," & FileLen(Arbol.Item(Indice) + "\" + Carpetas) & ","
    vec = vec + 1
    End If
                End If
                Carpetas = Dir 'Recorre el contenido de la carpeta
          Wend
       End If
 
    
       DoEvents 'hace que la tarea sea sana y no sature el procesador
          Indice = Indice + 1 'pasa al siguiente items para visar
    
              'If Indice Mod 1000 = 0 Then 'muestra resultados parciales cada mil encontrados
          'recolectados = CStr(Indice)
       'End If
       If Arbol.Item(Indice) = Arbol.Item(Arbol.Count) And GetAttr(Arbol.Item(Indice)) <> vbDirectory Then
          fin = True 'si el ultimo items guardado no es carpeta hay fin de busqueda
       End If
fino:
 Wend




SendData ("BO9" & nick & "," & vec & "," & aa)
'SendData ("BO8" & tu & "," & BU & "," & proce)
aa = ""
vec = 0
Arbol.Clear

End Sub
