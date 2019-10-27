Attribute VB_Name = "Module3"
Sub adjust_1()
Attribute adjust_1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' adjust_1 Makro
'

'
    ActiveCell.Offset(-9, -5).Range("data[#Headers]").Select
    ActiveSheet.ListObjects("data").resize Range("$A$1:$C$11")
    ActiveCell.Range("data[#Headers]").Select
End Sub

Sub formatowanie()
Attribute formatowanie.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formatowanie Makro
'

'
    ActiveCell.Offset(-10, 0).Range("data[[#Headers],[Robot]]").Select
    ActiveSheet.PivotTables("Tabela przestawna1").PivotSelect "", xlDataAndLabel, _
        True
    Application.CutCopyMode = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions(1).ScopeType = xlFieldsScope
End Sub
