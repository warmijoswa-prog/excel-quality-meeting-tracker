Attribute VB_Name = "Module4"
Public usuarioc As String
Public salir3 As Integer


Sub Macro4()
'
' Macro4 Macro
'

'
    Dim fecha As Date
    fecha = Now()
    Dim fila As Integer
    Dim comentario As String
    salir3 = 0
    
    Load UserForm3
    UserForm3.Show
    
    If salir3 = 10 Then GoTo FIN3
    
    fila = ActiveCell.Row
    Range("F" & fila).Select
    ActiveCell.Value = Format(Date, "dd/mm/yyyy")
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = ("Cerrado")
    ActiveCell.Offset(0, 1).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    
    comentario = ActiveCell.Value
    ActiveCell.Value = (Format(Day(fecha), "00") & "." & Format(Month(fecha), "00") & "." & Format(Year(fecha), "0000") & " [" & usuarioc & "] " & Chr(10) & comentario)
    Application.SendKeys "{F2}"
    Application.SendKeys "^{HOME}"
    Application.SendKeys "{END}"
FIN3:
End Sub
