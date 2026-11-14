Option Explicit

' Prueba de la opcion 3: macro minima que devuelve un valor a la terminal.
Public Function Hola() As String
    Hola = "hola desde VBA " & Format(Now, "hh:nn:ss")
End Function

' Opcion 5 (M4): mete en el libro el codigo escrito en la hoja "_codigo".
' B1 = nombre del modulo. Columna A = codigo, una linea por celda.
' Devuelve "OK ..." o "ERROR ..." sin abrir ventanas.
Public Function Cargar() As String
    Dim hoja As Worksheet, proyecto As Object, comp As Object
    Dim nombre As String, codigo As String, ultima As Long, i As Long
    On Error GoTo Fallo
    Set hoja = ThisWorkbook.Worksheets("_codigo")
    nombre = CStr(hoja.Range("B1").Value)
    ultima = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
    For i = 1 To ultima
        codigo = codigo & CStr(hoja.Cells(i, 1).Value) & vbLf
    Next i
    Set proyecto = ThisWorkbook.VBProject
    On Error Resume Next
    Set comp = proyecto.VBComponents(nombre)
    On Error GoTo Fallo
    If Not comp Is Nothing Then proyecto.VBComponents.Remove comp
    Set comp = proyecto.VBComponents.Add(1)
    comp.Name = nombre
    comp.CodeModule.AddFromString codigo
    Cargar = "OK " & nombre & " " & ultima & " lineas"
    Exit Function
Fallo:
    Cargar = "ERROR " & Err.Number & ": " & Err.Description
End Function
