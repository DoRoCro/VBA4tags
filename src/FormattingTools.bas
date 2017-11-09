Attribute VB_Name = "FormattingTools"
Sub ColourSelectedABC()
    Dim selectedRange As Range
    Dim entry As Range
    
    Set selectedRange = Application.Selection
    
    'Colour A,B,C to convention
    For Each entry In selectedRange
        'Debug.Print entry.Value
        With entry.Interior
            If IsError(entry.Value) Then
                GoTo skipCase
            End If
            Select Case entry.Value
                Case "A"
                    .PatternColorIndex = xlAutomatic
                    '.Pattern = xlSolid
                    .Color = 49407
                Case "B"
                    .PatternColorIndex = xlAutomatic
                    '.Pattern = xlSolid
                    .Color = 65535
                Case "C"
                    .PatternColorIndex = xlAutomatic
                    '.Pattern = xlSolid
                    .Color = 5296274
                Case True
                    entry.Style = "Calculation"
                Case False
                    entry.Style = "Calculation"
                Case Else
                    ' leave alone
            End Select
skipCase:
        End With
    Next entry
    'Centre all text in selection
    With selectedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        With .Borders
            .LineStyle = xlContinuous
            .Color = -8421505
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub
