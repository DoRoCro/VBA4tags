Attribute VB_Name = "SetupCriticalityTemplate"
Sub CreateWorksheetsFromFailureCodeList()
    Dim wb As Workbook
    Dim ws, newWs As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim rowsWithTagsCount As Integer
    
    
    Set wb = Workbooks("WND Criticality Template.xlsx")
    Set ws = wb.Sheets("FailureCodes")
    Set tbl = ws.ListObjects("ASSET_C_FailureCodesList")
    
    ' Debug.Print [ASSET_C_FailureCodesList[FailureCode]].Rows.Count
    Debug.Print tbl.ListRows.Count
    rowsWithTagsCount = 0
    For Each row In tbl.ListRows
        ' Only use code if it is used somewhere...
        ' Assumes an error code for any failure code not in use
        If WorksheetFunction.IsErr(rowCell(row, "Number found in ASSET-C WND")) Then
            rowsWithTagsCount = rowsWithTagsCount + 1
            Debug.Print rowsWithTagsCount, rowCell(row, "FailureCode"), rowCell(row, "Number found in ASSET-C WND")
            Set newWs = wb.Sheets.Add(After:= _
                 wb.Sheets(wb.Sheets.Count))
            newWs.Name = rowCell(row, "FailureCode")
        End If
        
        ' testing break after 5
        If rowsWithTagsCount >= 5 Then
            Exit For
        End If
        ' Insert default criticality assessment template here
        ' link output back to Failure codes sheet
        
        
    Next

End Sub
'From StackOverFlow, a handy function to lookup field in row by name
Function rowCell(row As ListRow, col As String) As Range
    Set rowCell = Intersect(row.Range, row.Parent.ListColumns(col).Range)
End Function

