Option Explicit

Rem Motor: puente entre la terminal y el VBA de cualquier libro abierto.
Rem Vive en el libro lanzador (prueba.xlsm). La terminal escribe el codigo en la hoja _codigo
Rem (columna A, cada linea con un caracter de marca delante para que Excel no se coma
Rem apostrofes ni signos =) y llama a estas funciones con run VB macro.
Rem Nada abre ventanas: todo devuelve texto "OK ..." o "ERROR n: ...".

Private Function Libro(nombre As String) As Workbook
    If Len(nombre) = 0 Then
        Set Libro = ThisWorkbook
    Else
        Set Libro = Workbooks(nombre)
    End If
End Function

Public Function CargarEn(nombreLibro As String, nombreModulo As String) As String
    Dim hoja As Worksheet, comp As Object, codigo As String, ultima As Long, i As Long
    On Error GoTo Fallo
    Set hoja = ThisWorkbook.Worksheets("_codigo")
    ultima = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
    For i = 1 To ultima
        codigo = codigo & Mid(CStr(hoja.Cells(i, 1).Value), 2) & vbLf
    Next i
    On Error Resume Next
    Set comp = Libro(nombreLibro).VBProject.VBComponents(nombreModulo)
    On Error GoTo Fallo
    If comp Is Nothing Then
        Set comp = Libro(nombreLibro).VBProject.VBComponents.Add(1)
        comp.Name = nombreModulo
    ElseIf comp.CodeModule.CountOfLines > 0 Then
        comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
    End If
    comp.CodeModule.AddFromString codigo
    CargarEn = "OK " & nombreModulo & " " & comp.CodeModule.CountOfLines & " lineas en " & Libro(nombreLibro).Name
    Exit Function
Fallo:
    CargarEn = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function Exportar(nombreLibro As String, nombreModulo As String) As String
    On Error GoTo Fallo
    With Libro(nombreLibro).VBProject.VBComponents(nombreModulo).CodeModule
        If .CountOfLines = 0 Then
            Exportar = ""
        Else
            Exportar = .Lines(1, .CountOfLines)
        End If
    End With
    Exit Function
Fallo:
    Exportar = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function Modulos(nombreLibro As String) As String
    Dim c As Object, s As String
    On Error GoTo Fallo
    For Each c In Libro(nombreLibro).VBProject.VBComponents
        s = s & c.Name & vbTab & c.Type & vbTab & c.CodeModule.CountOfLines & vbLf
    Next c
    Modulos = s
    Exit Function
Fallo:
    Modulos = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function Borrar(nombreLibro As String, nombreModulo As String) As String
    On Error GoTo Fallo
    With Libro(nombreLibro).VBProject
        .VBComponents.Remove .VBComponents(nombreModulo)
    End With
    Borrar = "OK borrado " & nombreModulo
    Exit Function
Fallo:
    Borrar = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function Ejecutar(nombreLibro As String, nombreMacro As String) As String
    Dim r As Variant
    On Error GoTo Fallo
    r = Application.Run("'" & Libro(nombreLibro).Name & "'!" & nombreMacro)
    Ejecutar = CStr(r)
    Exit Function
Fallo:
    Ejecutar = "ERROR " & Err.Number & ": " & Err.Description
End Function
