Sub To_sharepoint()
refreshsharepoint
'update share point
delete_Mode
update_Mode
TriggeringRecords
End Sub

Sub delete_Mode() 'delete data in sharepoint with table list Employee List
Dim cnt As ADODB.Connection
Dim rst As ADODB.Recordset
Dim mysql As String

Set cnt = New ADODB.Connection
Set rst = New ADODB.Recordset

mysql = "DELETE * FROM [Boss list];"

With cnt
    .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=http://sharepoint;List={044FDC15-B3FE-49C5-8FB1-8753DB7542C8};"
    .Open
End With
    
cnt.Execute mysql, , adCmdText

Debug.Print "delete"

If CBool(rst.State And adStateOpen) = True Then rst.Close
Set rst = Nothing
If CBool(cnt.State And adStateOpen) = True Then cnt.Close
Set cnt = Nothing
   
End Sub


Sub update_Mode()

Dim cnt As ADODB.Connection
Dim rst As ADODB.Recordset
Dim mysql As String
Dim i As Integer

Set cnt = New ADODB.Connection
Set rst = New ADODB.Recordset

mysql = "SELECT * FROM [Boss list];"

With cnt
    .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=http://sharepoint;List={044FDC15-B3FE-49C5-8FB1-8753DB7542C8};"
    .Open
End With
    
Debug.Print Range("Table3").Rows.Count + 1

rst.Open mysql, cnt, adOpenDynamic, adLockOptimistic
For i = 2 To Range("Table3").Rows.Count + 1
    If Sheets("boss list").Cells(i, "A").Value <> "" And Sheets("boss list").Cells(i, "F").Value <> "direct" And Sheets("boss list").Cells(i, "F").Value <> "-" Then
        rst.AddNew
        rst!Boss = Sheets("boss list").Cells(i, "A").Value
        rst!Pillar = Sheets("boss list").Cells(i, "B").Value
        rst!Department = Sheets("boss list").Cells(i, "C").Value
        rst!Section = Sheets("boss list").Cells(i, "D").Value
        rst!Email = Sheets("boss list").Cells(i, "E").Value
        rst!Initial = Sheets("boss list").Cells(i, "F").Value
        rst.Update
    End If
Next i



Debug.Print "update"

If CBool(rst.State And adStateOpen) = True Then rst.Close
Set rst = Nothing
If CBool(cnt.State And adStateOpen) = True Then cnt.Close
Set cnt = Nothing
   
End Sub


Sub refreshsharepoint()
   
    'clear table
    Sheets("boss list").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ListObjects("Table3").Resize Range("$A$1:$F$2")
    Range("Table3[[Pillar]:[Initial]]").Select
    Selection.NumberFormat = "General"

    'refresh
    Sheets("employee list").Select
    Range("A2").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    'filter
    
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Range("Table_bosslist[#All]").RemoveDuplicates Columns:=9, _
        Header:=xlYes
    ActiveSheet.ListObjects("Table_bosslist").Range.AutoFilter Field:=9, _
        Criteria1:="<>"
        
    'copy and paste
    Sheets("employee list").Select
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("boss list").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    'remove filter
    'refresh
    Sheets("employee list").Select
    Range("A2").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    'Sheets("employee list").Select
    'ActiveSheet.ListObjects("Table_bosslist").Range.AutoFilter Field:=9
    
    'set fomular
    'Range("B2:F2").Select
    'Selection.NumberFormat = "General"
    Sheets("boss list").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP([@[Boss Name]],Table_bosslist[[EE Name]:[Initial]],2,0),""-"")"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP([@[Boss Name]],Table_bosslist[[EE Name]:[Initial]],3,0),""-"")"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP([@[Boss Name]],Table_bosslist[[EE Name]:[Initial]],4,0),""-"")"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP([@[Boss Name]],Table_bosslist[[EE Name]:[Initial]],10,0),""-"")"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP([@[Boss Name]],Table_bosslist[[EE Name]:[Initial]],13,0),""-"")"
    Range("F3").Select
    
    'sort A-Z
    ActiveWorkbook.Worksheets("boss list").ListObjects("Table3").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("boss list").ListObjects("Table3").Sort.SortFields. _
        Add2 Key:=Range("Table3[[#All],[Initial]]"), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("boss list").ListObjects("Table3").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Debug.Print "refresh"
End Sub



