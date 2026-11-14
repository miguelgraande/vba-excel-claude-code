Option Explicit
' Comentario con apostrofe: comprueba que la marca protege las lineas.
' = Linea que empieza por igual dentro de un comentario.

Public Function ConControl() As String
    Dim x As Long
    On Error GoTo Fallo
    x = 1 / 0
    ConControl = "no deberia llegar"
    Exit Function
Fallo:
    ConControl = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function SinControl() As String
    Dim x As Long
    x = 1 / 0
    SinControl = "no deberia llegar"
End Function

Public Function ConMsgBox() As String
    MsgBox "ventana de prueba"
    ConMsgBox = "despues del MsgBox"
End Function
