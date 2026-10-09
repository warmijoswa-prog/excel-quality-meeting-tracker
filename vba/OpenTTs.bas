Attribute VB_Name = "Module5"

Sub Macro9()
Attribute Macro9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro9 Macro
'

'
    ActiveSheet.Range("$A$2:$J$10000").AutoFilter Field:=8, Criteria1:="Abierto"
    ActiveSheet.Range("$A$2:$J$10000").AutoFilter Field:=4, Criteria1:= _
    "<>*TAREA*", Operator:=xlAnd
    Range("A2").Select
    Range("D2").Select
    ActiveCell.Offset(1, 0).Select
    'BAJAR 1
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
End Sub
Sub Macro14()
Attribute Macro14.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro14 Macro
'

'
    Cells.Select
    Selection.ColumnWidth = 13
    Selection.RowHeight = 16
    Range("A1").Select
    Application.SendKeys "+{F10}"
    Application.SendKeys "{END}"
    Application.SendKeys "{UP}"
    Application.SendKeys "{UP}"
    Application.SendKeys "{UP}"
    Application.SendKeys "{UP}"
    Application.SendKeys "{RIGHT}"
    Application.SendKeys "M"
End Sub
