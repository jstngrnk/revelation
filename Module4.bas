Attribute VB_Name = "Module4"
Sub resize()
Attribute resize.VB_ProcData.VB_Invoke_Func = " \n14"
'
' resize Makro
'

'
    ActiveSheet.ListObjects("data").resize Range("$A$1:$C$7")
    ActiveCell.Select
End Sub
