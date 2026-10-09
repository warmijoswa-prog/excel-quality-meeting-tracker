Attribute VB_Name = "Module6"
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Range("A1").Select
    Range("Table_Query_from_AR_System_ODBC_Data_Source4[[#Headers],[Description]]") _
    .Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("K2:K10000").Select
    Selection.ClearContents
    Cells.Select
    Selection.ColumnWidth = 8
    Columns("A:A").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").Select
    Columns("B:B").EntireColumn.AutoFit
    Columns("E:E").Select
    Columns("E:E").EntireColumn.AutoFit
    Columns("K:M").Select
    Selection.ColumnWidth = 16
    Range("I:I").Select
    Selection.ColumnWidth = 0.1
    Range("L2:N2").Select
    Selection.AutoFill Destination:=Range("L2:N178")
    Range("L2:N178").Select
    Range("K2").Select
End Sub
