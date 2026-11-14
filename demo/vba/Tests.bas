Option Explicit

' Modulo Tests: compara el VL calculado en VBA con el VL esperado.
' La hoja "Esperado" la rellena la referencia en Python (referencia.py):
'   columna A fecha, columna B VL esperado. Las filas cuya columna A no es
'   una fecha (por ejemplo la cabecera) se saltan.
' Escribe el detalle en la hoja "Tests" y devuelve un resumen en texto.

Private Const TOLERANCIA As Double = 0.000001

Public Function EjecutarTests() As String
    Dim hEsp As Worksheet, hT As Worksheet
    Dim ultima As Long, r As Long, filaDatos As Long, salida As Long
    Dim f As Date, esperado As Double, obtenido As Double, dif As Double
    Dim n As Long, k As Long, fallos As String
    On Error GoTo Fallo
    Set hEsp = ThisWorkbook.Worksheets("Esperado")
    Set hT = HojaSalida("Tests")
    hT.Range("A1:E1").Value = Array("Fecha", "Esperado", "Obtenido", "Diferencia", "Resultado")
    hT.Range("A1:E1").Font.Bold = True
    salida = 1
    ultima = hEsp.Cells(hEsp.Rows.Count, 1).End(xlUp).Row
    For r = 1 To ultima
        If EsFecha(hEsp.Cells(r, 1).Value) Then
            f = ComoFecha(hEsp.Cells(r, 1).Value)
            esperado = ComoNumero(hEsp.Cells(r, 2).Value)
            filaDatos = FilaDeFecha(f)
            n = n + 1
            salida = salida + 1
            hT.Cells(salida, 1).Value = f
            hT.Cells(salida, 2).Value = esperado
            If filaDatos = 0 Then
                hT.Cells(salida, 3).Value = "sin datos"
                hT.Cells(salida, 5).Value = "FALLO"
                k = k + 1
                fallos = fallos & IIf(Len(fallos) > 0, ", ", "") & FechaTxt(f)
            Else
                obtenido = VL(filaDatos)
                dif = obtenido - esperado
                hT.Cells(salida, 3).Value = obtenido
                hT.Cells(salida, 4).Value = dif
                If Abs(dif) <= TOLERANCIA Then
                    hT.Cells(salida, 5).Value = "OK"
                Else
                    hT.Cells(salida, 5).Value = "FALLO"
                    k = k + 1
                    fallos = fallos & IIf(Len(fallos) > 0, ", ", "") & FechaTxt(f)
                End If
            End If
        End If
    Next r
    If n = 0 Then
        EjecutarTests = "ERROR 1006: la hoja Esperado no tiene filas con fecha"
        Exit Function
    End If
    hT.Range("A2:A" & salida).NumberFormat = "yyyy-mm-dd"
    hT.Range("B2:C" & salida).NumberFormat = "0.000000"
    hT.Range("D2:D" & salida).NumberFormat = "0.00E+00"
    hT.Columns("A:E").AutoFit
    If k = 0 Then
        EjecutarTests = "OK " & n & "/" & n
    Else
        EjecutarTests = "FALLO " & k & " de " & n & ": " & fallos
    End If
    Exit Function
Fallo:
    EjecutarTests = "ERROR " & Err.Number & ": " & Err.Description
End Function
