Attribute VB_Name = "Module2"
Public eebb As String
Public servicios As String
Public ciudad As String
Public noc As String
Public escalaanoc As String
Public salir2 As Integer


Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    fecha = Now()
    noc = False
    Rows("2:2").Select
    Selection.AutoFilter
    Range("A1048576").Select
    Selection.End(xlUp).Select
    fila = ActiveCell.Row
    Range("A" & fila + 1).Select
    salir2 = 0

    Load UserForm2
    UserForm2.Show
    
    If salir2 = 10 Then GoTo FIN2:
    ActiveCell.Value = ("EEBB " & eebb & Chr(10) & "Degradación " & servicios)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = ciudad
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = ("NOC")
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = Format(Date, "dd/mm/yyyy")
    ActiveCell.Offset(0, 2).Select
    ActiveCell.Value = "=IF(" & ActiveCell.Offset(0, -1).Address & "=" & Chr(34) & Chr(34) & ",Now()-" & ActiveCell.Offset(0, -2).Address & ",(" & ActiveCell.Offset(0, -1).Address & "-" & ActiveCell.Offset(0, -2).Address & "))"
    Columns("G:G").Select
    Selection.Replace What:="$", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Range("G1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = ("Abierto")
    ActiveCell.Offset(0, 1).Select
    
    If noc = True Then GoTo ESCALADONOC Else GoTo ESCALADOOTRO
ESCALADONOC:
    ActiveCell.Value = (Format(Day(fecha), "00") & "." & Format(Month(fecha), "00") & "." & Format(Year(fecha), "0000") & " [" & escalaanoc & "] Se escala degradación para revisión del NOC. EEBB presenta alarmas.")
GoTo FIN2
ESCALADOOTRO:
    ActiveCell.Value = (Format(Day(fecha), "00") & "." & Format(Month(fecha), "00") & "." & Format(Year(fecha), "0000") & " [] ")
    Application.SendKeys "{F2}"
    Application.SendKeys "{LEFT}"
    Application.SendKeys "{LEFT}"
FIN2:
End Sub



