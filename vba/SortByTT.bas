Attribute VB_Name = "Module7"
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
Dim fila As Integer


    Range("H2").Select
    Selection.End(xlDown).Select
    fila = ActiveCell.Row
    Range("A2:M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("bitácora").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("bitácora").Sort.SortFields.Add Key:=Range( _
        "D3:D" & fila), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("bitácora").Sort
        .SetRange Range("A2:I" & fila)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G3").Select
    Selection.AutoFill Destination:=Range("G3:G" & fila)
    Range("D2").Select
End Sub
