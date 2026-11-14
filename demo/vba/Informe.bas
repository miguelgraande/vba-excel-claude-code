Option Explicit

' Modulo Informe: escribe la hoja "Informe" con el detalle del VL de cada dia.
' La crea si no existe y la vacia si ya existe.

Public Function GenerarInforme() As String
    Dim h As Worksheet, r As Long, salida As Long, ultima As Long
    On Error GoTo Fallo
    ultima = UltimaFila()
    Set h = HojaSalida("Informe")
    h.Range("A1:H1").Value = Array("Fecha", "Patrimonio", "Participaciones", _
        "Comision gestion %", "Dias devengo", "Comision devengada", _
        "Patrimonio neto", "Valor liquidativo")
    h.Range("A1:H1").Font.Bold = True
    salida = 1
    For r = PRIMERA_FILA To ultima
        salida = salida + 1
        h.Cells(salida, 1).Value = FechaDia(r)
        h.Cells(salida, 2).Value = PatrimonioDia(r)
        h.Cells(salida, 3).Value = ParticipacionesDia(r)
        h.Cells(salida, 4).Value = ComisionPctDia(r)
        h.Cells(salida, 5).Value = DiasDevengo(r)
        h.Cells(salida, 6).Value = ComisionDevengada(r)
        h.Cells(salida, 7).Value = PatrimonioNeto(r)
        h.Cells(salida, 8).Value = VL(r)
    Next r
    If salida > 1 Then
        h.Range("A2:A" & salida).NumberFormat = "yyyy-mm-dd"
        h.Range("B2:C" & salida).NumberFormat = "#,##0.00"
        h.Range("D2:D" & salida).NumberFormat = "0.00"
        h.Range("E2:E" & salida).NumberFormat = "0"
        h.Range("F2:G" & salida).NumberFormat = "#,##0.00"
        h.Range("H2:H" & salida).NumberFormat = "0.000000"
    End If
    h.Columns("A:H").AutoFit
    GenerarInforme = "OK Informe " & (salida - 1) & " dias"
    Exit Function
Fallo:
    GenerarInforme = "ERROR " & Err.Number & ": " & Err.Description
End Function
