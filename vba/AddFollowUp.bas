Attribute VB_Name = "Module1"
Public usuario As String
Public salir1 As Integer


Sub Macro1()
'
' Macro1 Macro
'

'
    Dim fecha As Date
    fecha = Now()
    Dim fila As Integer
    Dim comentario As String
    salir1 = 0
    
    Load UserForm1
    UserForm1.Show
    
    If salir1 = 10 Then GoTo FIN1
    
    fila = ActiveCell.Row
    Range("I" & fila).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    comentario = ActiveCell.Value
    ActiveCell.Value = (Format(Day(fecha), "00") & "." & Format(Month(fecha), "00") & "." & Format(Year(fecha), "0000") & " [" & usuario & "] " & Chr(10) & comentario)
    Application.SendKeys "{F2}"
    Application.SendKeys "^{HOME}"
    Application.SendKeys "{END}"
FIN1:
End Sub
