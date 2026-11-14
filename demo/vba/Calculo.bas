Option Explicit

' Modulo Calculo: valor liquidativo (VL) diario del fondo.
' VL = (patrimonio - comision de gestion devengada) / participaciones
' "fila" es el numero de fila de la hoja Datos (2 = primer dia).

' Tasa diaria equivalente a la comision de gestion anual del dia.
Public Function TasaDiaria(fila As Long) As Double
    TasaDiaria = ComisionPctDia(fila) / 100 / 360
End Function

' Comision de gestion devengada desde la valoracion anterior (EUR).
Public Function ComisionDevengada(fila As Long) As Double
    ComisionDevengada = PatrimonioDia(fila) * TasaDiaria(fila) * DiasDevengo(fila)
End Function

Public Function PatrimonioNeto(fila As Long) As Double
    PatrimonioNeto = PatrimonioDia(fila) - ComisionDevengada(fila)
End Function

Public Function VL(fila As Long) As Double
    VL = PatrimonioNeto(fila) / ParticipacionesDia(fila)
End Function

' Para la terminal: VL de un dia, por fecha en texto aaaa-mm-dd.
Public Function VLDeFecha(fechaTexto As String) As String
    Dim r As Long
    On Error GoTo Fallo
    r = FilaDeFecha(ComoFecha(fechaTexto))
    If r = 0 Then
        VLDeFecha = "ERROR 1005: no hay datos para " & fechaTexto
    Else
        VLDeFecha = FechaTxt(FechaDia(r)) & " " & NumTxt(VL(r), 6)
    End If
    Exit Function
Fallo:
    VLDeFecha = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Para la terminal: VL de todos los dias, una linea por dia.
Public Function VLTodos() As String
    Dim r As Long, s As String
    On Error GoTo Fallo
    For r = PRIMERA_FILA To UltimaFila()
        s = s & FechaTxt(FechaDia(r)) & " " & NumTxt(VL(r), 6) & vbLf
    Next r
    VLTodos = s
    Exit Function
Fallo:
    VLTodos = "ERROR " & Err.Number & ": " & Err.Description
End Function
