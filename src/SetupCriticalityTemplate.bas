Attribute VB_Name = "SetupCriticalityTemplate"
Const wbCriticality = "WND Criticality Template.xlsx"

Sub CreateWorksheetsFromFailureCodeList()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim templateWs As Worksheet
    Dim tbl, fCodesTbl As ListObject
    Dim row As ListRow
    Dim rowsWithTagsCount As Integer
    
    
    Set wb = Workbooks(wbCriticality)
    Set ws = wb.Worksheets("FailureCodes")
    Set templateWs = wb.Worksheets("FailureCodeTemplate")
    Set tbl = ws.ListObjects("ASSET_C_FailureCodesList")
    
    
    ' Debug.Print [ASSET_C_FailureCodesList[FailureCode]].Rows.Count
    Debug.Print tbl.ListRows.Count
    rowsWithTagsCount = 0
    For Each row In tbl.ListRows
        ' Only use code if it is used somewhere...
        ' Assumes an error code for any failure code not in use (#REF! from GETPIVOTDATA function)
        If Not WorksheetFunction.IsErr(rowCell(row, "Number found in ASSET-C WND")) Then
            
            rowsWithTagsCount = rowsWithTagsCount + 1    ' counting found rows
            
            Debug.Print rowsWithTagsCount, rowCell(row, "FailureCode"), rowCell(row, "Number found in ASSET-C WND")
            
            'Set newWs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            templateWs.Copy After:=wb.Sheets(wb.Sheets.Count)
            Set newWs = wb.Sheets(wb.Sheets.Count)    ' can't do this on previous line as Copy is a Sub procedure (I think)
            newWs.Name = rowCell(row, "FailureCode")  ' name sheet from failure code
            
            ' Insert default criticality assessment template here
            Call CopyDefaultCriticalitiesIntoTemplateWorksheet(row, newWs)
            
            ' link output back to Failure codes sheet
        End If
        
        ' testing break after 5 to avoid deleting too many
        If rowsWithTagsCount >= 5 Then
            Exit For
        End If
       
        
    Next

End Sub
'From StackOverFlow, a handy function to lookup field in row by name
Function rowCell(row As ListRow, col As String) As Range
    Set rowCell = Intersect(row.Range, row.Parent.ListColumns(col).Range)
End Function
'This could be copied to FailureCodeDefaultCriticality worksheet code
Sub CopyDefaultCriticalitiesIntoTemplateWorksheet(codeRow As ListRow, ws As Worksheet)

    Dim wb As Workbook
    Dim fcdcWs As Worksheet
    Dim codeStr As String
    Dim fcdcTbl As ListObject
    
    Set wb = Workbooks(wbCriticality)
    Set fcdcWs = wb.Worksheets("FailurecodeDefaultCriticality")
    Set fcdcTbl = fcdcWs.ListObjects("FailurecodeDefaultCriticalities_Table")
    
    codeStr = rowCell(codeRow, "FailureCode").Value
    With ws      'remember this is the target ws
        .Range("B1").Formula = rowCell(codeRow, "FailureCode")
        .Range("B2").Formula = rowCell(codeRow, "Description")
        
    End With
End Sub
