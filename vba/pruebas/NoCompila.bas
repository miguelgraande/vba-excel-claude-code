Option Explicit

Public Function Bien() As String
    Bien = "compila"
End Function

Public Function Mal() As String
    Dim x As Long
    x = noDeclarada + 1
    Mal = "no deberia llegar"
End Function
