Attribute VB_Name = "Module2"
Sub ConnectToOracle()

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim mtxData As Variant

Worksheets(1).Activate
ActiveSheet.ListObjects("data").DataBodyRange.Value = ""

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

cn.Open ( _
"User ID=HR" & _
";Password=oracle_test" & _
";Data Source=xe" & _
";Provider=OraOLEDB.Oracle")

rs.CursorType = adOpenForwardOnly
rs.Open ("select * from hr.datasets"), cn

mtxData = Application.Transpose(rs.GetRows)

Worksheets(1).Activate
'ActiveSheet.Range("a1:a2") = mtxData
ActiveSheet.ListObjects("data").DataBodyRange.Resize(UBound(mtxData, 1) - LBound(mtxData, 1) + 1, UBound(mtxData, 2) - LBound(mtxData, 2) + 1) = mtxData


'Cleanup in the end
Set rs = Nothing
Set cn = Nothing

End Sub
Sub pivot()
Attribute pivot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' pivot Makro
'

'
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "data", Version:=6).CreatePivotTable TableDestination:="Arkusz1!R1C6", _
        TableName:="Tabela przestawna1", DefaultVersion:=6
    Sheets("Arkusz1").Select
    Cells(1, 6).Select
    With ActiveSheet.PivotTables("Tabela przestawna1").PivotFields("Robot")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabela przestawna1").PivotFields("Data")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabela przestawna1").AddDataField ActiveSheet. _
        PivotTables("Tabela przestawna1").PivotFields("Dataset"), "Liczba z Dataset", _
        xlCount
End Sub
Sub usuniecie_pivota()
Attribute usuniecie_pivota.VB_ProcData.VB_Invoke_Func = " \n14"
'
' usuniecie_pivota Makro
'

'
    ActiveSheet.PivotTables("Tabela przestawna1").PivotSelect "", xlDataAndLabel, _
        True
    Selection.ClearContents
    ActiveCell.Select
End Sub
Sub refresh()
Attribute refresh.VB_ProcData.VB_Invoke_Func = " \n14"
'
' refresh Makro
'

'
    ActiveSheet.PivotTables("Tabela przestawna1").PivotCache.refresh
End Sub
