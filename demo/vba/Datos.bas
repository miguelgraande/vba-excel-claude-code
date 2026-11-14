Option Explicit

' Modulo Datos: lectura de la hoja "Datos" de este libro.
' Columnas: A fecha, B patrimonio (EUR, antes de la comision del dia),
'           C participaciones en circulacion, D comision de gestion (% anual).
' Fila 1 = cabecera. Cada fila siguiente es un dia de valoracion.
' "fila" es siempre el numero de fila de la hoja (2 = primer dia).

Private Const HOJA_DATOS As String = "Datos"
Public Const PRIMERA_FILA As Long = 2

Private Const COL_FECHA As Long = 1
Private Const COL_PATRIMONIO As Long = 2
Private Const COL_PARTICIPACIONES As Long = 3
Private Const COL_COMISION As Long = 4

Public Function HojaDatos() As Worksheet
    Set HojaDatos = ThisWorkbook.Worksheets(HOJA_DATOS)
End Function

Public Function UltimaFila() As Long
    Dim h As Worksheet
    Set h = HojaDatos()
    UltimaFila = h.Cells(h.Rows.Count, COL_FECHA).End(xlUp).Row
End Function

Public Function FechaDia(fila As Long) As Date
    FechaDia = ComoFecha(HojaDatos().Cells(fila, COL_FECHA).Value)
End Function

Public Function PatrimonioDia(fila As Long) As Double
    PatrimonioDia = ComoNumero(HojaDatos().Cells(fila, COL_PATRIMONIO).Value)
End Function

Public Function ParticipacionesDia(fila As Long) As Double
    ParticipacionesDia = ComoNumero(HojaDatos().Cells(fila, COL_PARTICIPACIONES).Value)
    If ParticipacionesDia <= 0 Then
        Err.Raise 1002, "Datos", "Participaciones no validas en la fila " & fila
    End If
End Function

Public Function ComisionPctDia(fila As Long) As Double
    ComisionPctDia = ComoNumero(HojaDatos().Cells(fila, COL_COMISION).Value)
End Function

' Dias naturales desde la valoracion anterior. El primer dia cuenta 1.
Public Function DiasDevengo(fila As Long) As Long
    If fila <= PRIMERA_FILA Then
        DiasDevengo = 1
    Else
        DiasDevengo = CLng(Int(CDbl(FechaDia(fila))) - Int(CDbl(FechaDia(fila - 1))))
        If DiasDevengo <= 0 Then
            Err.Raise 1003, "Datos", "Fechas no ordenadas en la fila " & fila
        End If
    End If
End Function

' Fila de la hoja Datos con esa fecha, o 0 si no existe.
Public Function FilaDeFecha(f As Date) As Long
    Dim r As Long, ultima As Long, buscada As Long
    buscada = CLng(Int(CDbl(f)))
    ultima = UltimaFila()
    For r = PRIMERA_FILA To ultima
        If CLng(Int(CDbl(FechaDia(r)))) = buscada Then
            FilaDeFecha = r
            Exit Function
        End If
    Next r
    FilaDeFecha = 0
End Function

' Convierte una celda en fecha: admite fecha de Excel, numero de serie,
' texto aaaa-mm-dd o texto dd/mm/aaaa.
Public Function ComoFecha(v As Variant) As Date
    Dim s As String
    If IsEmpty(v) Then
        Err.Raise 1001, "Datos", "Fecha vacia"
    ElseIf VarType(v) = vbDate Then
        ComoFecha = CDate(v)
    ElseIf VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        ComoFecha = CDate(CDbl(v))
    Else
        s = Trim(CStr(v))
        If Len(s) = 10 And Mid(s, 5, 1) = "-" And Mid(s, 8, 1) = "-" Then
            ComoFecha = DateSerial(CInt(Left(s, 4)), CInt(Mid(s, 6, 2)), CInt(Right(s, 2)))
        ElseIf Len(s) = 10 And Mid(s, 3, 1) = "/" And Mid(s, 6, 1) = "/" Then
            ComoFecha = DateSerial(CInt(Right(s, 4)), CInt(Mid(s, 4, 2)), CInt(Left(s, 2)))
        Else
            Err.Raise 1001, "Datos", "Fecha no valida: " & s
        End If
    End If
End Function

Public Function EsFecha(v As Variant) As Boolean
    Dim d As Date
    On Error GoTo NoEs
    d = ComoFecha(v)
    EsFecha = True
    Exit Function
NoEs:
    EsFecha = False
End Function

' Convierte una celda en numero. Si viene como texto, acepta punto decimal.
Public Function ComoNumero(v As Variant) As Double
    Dim s As String
    If VarType(v) = vbString Then
        s = Trim(v)
        If Len(s) = 0 Then Err.Raise 1004, "Datos", "Numero vacio"
        ComoNumero = Val(s)
    ElseIf IsEmpty(v) Then
        Err.Raise 1004, "Datos", "Numero vacio"
    Else
        ComoNumero = CDbl(v)
    End If
End Function

' Numero como texto con punto decimal, sin depender del idioma del sistema.
Public Function NumTxt(x As Double, Optional decimales As Long = 6) As String
    NumTxt = Replace(Format(x, "0." & String(decimales, "0")), ",", ".")
End Function

Public Function FechaTxt(f As Date) As String
    FechaTxt = Format(f, "yyyy-mm-dd")
End Function

' Devuelve la hoja con ese nombre, vacia. La crea al final si no existe.
Public Function HojaSalida(nombre As String) As Worksheet
    Dim h As Worksheet
    On Error Resume Next
    Set h = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If h Is Nothing Then
        Set h = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        h.Name = nombre
    Else
        h.Cells.Clear
    End If
    Set HojaSalida = h
End Function

' Para la terminal: cuantos dias hay en la hoja Datos.
Public Function ContarDias() As String
    On Error GoTo Fallo
    ContarDias = "OK " & (UltimaFila() - PRIMERA_FILA + 1) & " dias, de " & _
        FechaTxt(FechaDia(PRIMERA_FILA)) & " a " & FechaTxt(FechaDia(UltimaFila()))
    Exit Function
Fallo:
    ContarDias = "ERROR " & Err.Number & ": " & Err.Description
End Function
