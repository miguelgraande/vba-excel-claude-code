Option Explicit
' SinVentanas: sustituye MsgBox e InputBox dentro del proyecto para que el codigo no abra ventanas.
' VBA busca primero los nombres en el propio proyecto, asi que estas versiones se usan en lugar
' de las de VBA. Cada llamada queda apuntada en la hoja _avisos (fecha, tipo, texto).
' Limite: no cubre Application.InputBox ni los cuadros de dialogo de Excel.

Public Function MsgBox(Prompt As Variant, Optional Buttons As Variant, Optional Title As Variant, _
                       Optional HelpFile As Variant, Optional Context As Variant) As VbMsgBoxResult
    Apuntar "MsgBox", CStr(Prompt)
    MsgBox = vbOK
End Function

Public Function InputBox(Prompt As Variant, Optional Title As Variant, Optional Default As Variant, _
                         Optional XPos As Variant, Optional YPos As Variant, _
                         Optional HelpFile As Variant, Optional Context As Variant) As String
    Apuntar "InputBox", CStr(Prompt)
    If IsMissing(Default) Then InputBox = "" Else InputBox = CStr(Default)
End Function

Private Sub Apuntar(tipo As String, texto As String)
    Dim h As Worksheet, f As Long
    On Error Resume Next
    Set h = ThisWorkbook.Worksheets("_avisos")
    On Error GoTo 0
    If h Is Nothing Then
        Set h = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        h.Name = "_avisos"
    End If
    f = h.Cells(h.Rows.Count, 1).End(xlUp).Row + 1
    h.Cells(f, 1).Value = Now
    h.Cells(f, 2).Value = tipo
    h.Cells(f, 3).Value = texto
End Sub
