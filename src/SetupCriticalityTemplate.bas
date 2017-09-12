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
    Dim fcdcTbl As ListObject    'Failure code default criticalities table
    
    Set wb = Workbooks(wbCriticality)
    Set ws = wb.Worksheets("FailureCodes")
    Set templateWs = wb.Worksheets("FailureCodeTemplate")
    Set tbl = ws.ListObjects("ASSET_C_FailureCodesList")
    Set fcdcWs = wb.Worksheets("FailurecodeDefaultCriticality")
    Set fcdcTbl = fcdcWs.ListObjects("FailurecodeDefaultCriticalities_Table")
    
    
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

Function getRow(Table As ListObject, ColumnName As String, Key As String) As ListRow
    On Error Resume Next
    Dim row As ListRow
    'Set GetRow = Table.ListColumns(ColumnName) _
    '    .Rows(WorksheetFunction.Match(Key, Table.ListColumns(ColumnName).Range, 0))
    
    For Each row In Table.ListRows
        If rowCell(row, ColumnName).Value = Key Then
            Set getRow = row
            Exit Function
        End If
    Next
    If Err.Number <> 0 Then
        Err.Clear
        Set getRow = Nothing
    End If
End Function

'This could be copied to FailureCodeDefaultCriticality worksheet code
Sub CopyDefaultCriticalitiesIntoTemplateWorksheet(codeRow As ListRow, _
                                                  ws As Worksheet, _
                                                  fcdcTbl As ListObject)

    Dim wb As Workbook
    Dim codeStr As String
    Dim defaultsRow As ListRow
    
    'Set wb = Workbooks(wbCriticality)
    
    codeStr = rowCell(codeRow, "FailureCode").Value
    Set defaultsRow = getRow(fcdcTbl, "FailureCode", codeStr)
    
    Debug.Print defaultsRow.Range.Address
    
    With ws      'remember this is the target ws
        .Range("B1").Formula = rowCell(codeRow, "FailureCode")
        .Range("B2").Formula = rowCell(codeRow, "Description")
        ' find row in fcdcTbl then lookup value
        .Range("B16").Formula = Left(rowCell(defaultsRow, "SC_Impact"), 1)
        .Range("C16").Formula = rowCell(defaultsRow, "SC_Likelihood")
    End With
End Sub