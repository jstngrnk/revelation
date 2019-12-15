Attribute VB_Name = "Connect_to_DB"
Sub ConnectToOracle()

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim mtxData As Variant

Worksheets(1).Activate
'Uzupe³nij dane tabeli
ActiveSheet.ListObjects("Tabela1").DataBodyRange.Value = ""

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

cn.Open ( _
"User ID=HR" & _
";Password=oracle_test" & _
";Data Source=xe" & _
";Provider=OraOLEDB.Oracle")

'db_query = Range("B2").Value
'Odniesienie do db_query
db_query = Worksheets("Arkusz2").Range("B2").Value

rs.CursorType = adOpenForwardOnly
rs.Open (db_query), cn

mtxData = Application.Transpose(rs.GetRows)

Worksheets(1).Activate
'ActiveSheet.Range("a1:a2") = mtxData
ActiveSheet.ListObjects("Tabela1").DataBodyRange.Resize(UBound(mtxData, 1) - LBound(mtxData, 1) + 1, UBound(mtxData, 2) - LBound(mtxData, 2) + 1) = mtxData

    With Arkusz1
        .ListObjects("Tabela1").Resize .Range("A1").Offset(0, 0).Resize(UBound(mtxData, 1) - LBound(mtxData, 1) + 2, UBound(mtxData, 2) - LBound(mtxData, 2) + 1)
    End With

'Cleanup in the end
Set rs = Nothing
Set cn = Nothing

Sheets("Arkusz2").Select

ActiveSheet.PivotTables("Tabela przestawna1").PivotCache.Refresh

MsgBox ("Data updated")

End Sub


Sub refresh_pivot()
    Sheets("Arkusz2").Select

    ActiveSheet.PivotTables("Tabela przestawna1").PivotCache.Refresh
End Sub


