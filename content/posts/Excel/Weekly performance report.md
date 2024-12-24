JiraRaw_Test
```vb
Sub JiraRaw_Test()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim lastRow3 As Long
    Dim lastRow4 As Long
    ' Set the active worksheet to a variable for better readability
    Set ws = ActiveSheet
    With ws
        On Error Resume Next
        .Shapes.Range(Array("Picture 1")).Delete
        On Error GoTo 0
        .Cells.UnMerge
        ' Clear all borders
        .Cells.Borders.LineStyle = xlLineStyleNone
        ' Delete the first three rows
        .Rows("1:3").Delete
        ' Format the first row
        .Rows(1).Interior.Color = RGB(64, 64, 64)
        .Rows(1).Font.Color = RGB(255, 255, 255)
        ' Find the last rows for various columns
        lastRow = .Cells(.Rows.Count, "K").End(xlUp).Row
        lastRow1 = .Cells(.Rows.Count, "L").End(xlUp).Row
        lastRow2 = .Cells(.Rows.Count, "Q").End(xlUp).Row
        lastRow3 = .Cells(.Rows.Count, "B").End(xlUp).Row
        lastRow4 = .Cells(.Rows.Count, "J").End(xlUp).Row
        lastRow5 = .Cells(.Rows.Count, "F").End(xlUp).Row
        ' Insert new columns and add headers
        .Columns("G:H").Insert Shift:=xlToRight
        .Cells(1, "G").Value = "WP short"
        .Cells(1, "H").Value = "Type of work"
        .Columns("N:O").Insert Shift:=xlToRight
        .Cells(1, "N").Value = "Original Estimate, H"
        .Cells(1, "O").Value = "Remaining Estimate, H"
        .Columns("R").Insert Shift:=xlToRight
        .Cells(1, "R").Value = "Sprint#"
        .Columns("T").Insert Shift:=xlToRight
        .Cells(1, "T").Value = "Status2"
        .Columns("U").Insert Shift:=xlToRight
        .Cells(1, "U").Value = "Project"
        .Columns("V").Insert Shift:=xlToRight
        .Cells(1, "V").Value = "EV"
        ' Apply formulas to the respective columns
        .Range("M2:M" & lastRow).Formula = "=K2/60/60"
        .Range("N2:N" & lastRow1).Formula = "=L2/60/60"
        .Range("R2:R" & lastRow2).Formula = "=MID(Q2, FIND(""SP"", Q2), 4)"
        .Range("R2").AutoFill Destination:=.Range("R2:R" & lastRow)
        .Range("U2:U" & lastRow3).Formula = "=LEFT(B2,3)"
        .Range("V2:V" & lastRow4).Formula = "=J2*4"
        .Range("G2:G" & lastRow5).Formula = "=TEXTBEFORE(F2,"" - "")"
        .Cells(1, 1).Select
    End With
End Sub
```
CreatePivotTable_Test

```vb
Sub CreatePivotTable_Test()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As pivotCache
    Dim pivotTable As pivotTable
    Dim pivotTable2 As pivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim pSlicersCaches As SlicerCaches
    Dim sSlicerCache As SlicerCache
    Dim sSlicer As Slicer

    ' Data sheet and range
    Set wsData = ThisWorkbook.Worksheets("AC")
    Set pivotRange = wsData.Range("A1").CurrentRegion
    Set wsData2 = ThisWorkbook.Worksheets("JIRA Dec")
    Set pivotRange2 = wsData2.Range("A1").CurrentRegion

    If Not ThisWorkbook.Sheets("SP_TEST") Is Nothing Then
        ThisWorkbook.Sheets("SP_TEST").Delete
    End If

    ' New sheet
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("SP_TEST")
    'Set wsPivot2 = ThisWorkbook.Worksheets("SP_TEST")
    If wsPivot Is Nothing Then
       Set wsPivot = ThisWorkbook.Worksheets.Add
       wsPivot.Name = "SP_TEST"
    End If
    On Error GoTo 0

    Set pivotDestination = wsPivot.Range("A7")
    Set pivotDestination2 = wsPivot.Range("A55")
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotCache2 = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange2)

    ' Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="AC _MyPivotTable")
    pivotTable.TableStyle2 = "PivotStyleMedium15"
    Set pivotTable2 = pivotCache2.CreatePivotTable(TableDestination:=pivotDestination2, TableName:="JIRA _MyPivotTable2")
    pivotTable2.TableStyle2 = "PivotStyleMedium15"

    ' Fields
    With pivotTable
        .PivotFields("Function").Orientation = xlRowField
        .PivotFields("WP").Orientation = xlRowField
        On Error Resume Next
        .PivotFields("ETC JIRA SP1.1").Orientation = xlHidden
        On Error GoTo 0
        .AddDataField .PivotFields("ETC JIRA SP1.1"), " ETC JIRA SP1.1", xlMax
        .PivotFields("Week 1 ").Orientation = xlDataField
        .PivotFields("Week 2 ").Orientation = xlDataField
        .PivotFields("AC SP1.1").Orientation = xlDataField
        On Error Resume Next
         .PivotFields("Total Issue SP1.1").Orientation = xlHidden
         .PivotFields("Closed SP1.1").Orientation = xlHidden
         .PivotFields("Resolved SP1.1").Orientation = xlHidden
         .PivotFields("in progress SP1.1").Orientation = xlHidden
        On Error GoTo 0
        .AddDataField .PivotFields("Total Issue SP1.1"), " Total Issue SP1.1", xlMax
        .AddDataField .PivotFields("Closed SP1.1"), " Closed SP1.1", xlMax
        .AddDataField .PivotFields("Resolved SP1.1"), " Resolved SP1.1", xlMax
        .AddDataField .PivotFields("in progress SP1.1"), " in progress SP1.1", xlMax
        On Error Resume Next
        .PivotFields("EV SP1.1").Orientation = xlHidden
        On Error GoTo 0
        .AddDataField .PivotFields("EV SP1.1"), " EV SP1.1", xlMax
        On Error Resume Next
        .CalculatedFields.Add "EV1.1, %", "=IFERROR(EV SP1.1/ETC JIRA SP1.1,0)"
        .AddDataField .PivotFields("EV1.1, %"), " EV1.1, %", xlMin
        .CalculatedFields.Add "EV1.1, %", "=IFERROR((EV SP1.1/ETC JIRA SP1.1)*2,0)"
        .AddDataField .PivotFields("EV1.1, %"), " EV1.1, %", xlMin
        On Error Resume Next
        .PivotFields("ETC JIRA SP1.2").Orientation = xlHidden
        On Error GoTo 0
        .AddDataField .PivotFields("ETC JIRA SP1.2"), " ETC JIRA SP1.2", xlMax
        .PivotFields("Week 3 ").Orientation = xlDataField
        .PivotFields("Week 4 ").Orientation = xlDataField
        .PivotFields("Week 5 ").Orientation = xlDataField
        .PivotFields("AC SP1.2").Orientation = xlDataField
        On Error Resume Next
         .PivotFields("Total Issue SP1.2").Orientation = xlHidden
         .PivotFields("Closed SP1.2").Orientation = xlHidden
         .PivotFields("Resolved SP1.2").Orientation = xlHidden
         .PivotFields("in progress SP1.2").Orientation = xlHidden
        On Error GoTo 0
        .AddDataField .PivotFields("Total Issue SP1.2"), " Total Issue SP1.2", xlMax
        .AddDataField .PivotFields("Closed SP1.2"), " Closed SP1.2", xlMax
        .AddDataField .PivotFields("Resolved SP1.2"), " Resolved SP1.2", xlMax
        .AddDataField .PivotFields("in progress SP1.2"), " in progress SP1.2", xlMax
        On Error Resume Next
        .PivotFields("EV SP1.2").Orientation = xlHidden
        On Error GoTo 0
        .AddDataField .PivotFields("EV SP1.2"), " EV SP1.2", xlMax
        .CalculatedFields.Add "EV1.2, %", "=IFERROR(EV SP1.2/ETC JIRA SP1.2,0)"
        '= 'EV SP16'/ 'ETC JIRA SP16'
        .AddDataField .PivotFields("EV1.2, %"), " EV1.2, %", xlMin
        On Error GoTo 0
        On Error Resume Next
        .PivotFields("Week 1 ").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Week 1 ': " & Err.Description
            .PivotFields("Week 1 ").Function = xlCount
        End If
        .PivotFields("Week 2 ").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Week 2 ': " & Err.Description
            .PivotFields("Week 2 ").Function = xlCount
        End If
        .PivotFields("AC SP1.1").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'AC SP1.1': " & Err.Description
            .PivotFields("AC SP1.1").Function = xlCount
        End If
        .PivotFields("Total Issue SP1.1").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Total Issue SP1.1': " & Err.Description
            .PivotFields("Total Issue SP1.1").Function = xlCount
        End If
        .PivotFields("Closed SP1.1").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Closed SP1.1': " & Err.Description
            .PivotFields("Closed SP1.1").Function = xlCount
        End If
         .PivotFields("Resolved SP1.1").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Resolved SP1.1': " & Err.Description
            .PivotFields("Resolved SP1.1").Function = xlCount
        End If
        .PivotFields("in progress SP1.1").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'in progress SP1.1': " & Err.Description
            .PivotFields("in progress SP1.1").Function = xlCount
        End If
        .PivotFields("Week 3 ").Function = xlSum
         If Err.Number <> 0 Then
            Debug.Print "Error with 'Week 1 ': " & Err.Description
            .PivotFields("Week 3 ").Function = xlCount
        End If
        .PivotFields("Week 4 ").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Week 4 ': " & Err.Description
            .PivotFields("Week 4 ").Function = xlCount
        End If
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Week 5 ': " & Err.Description
            .PivotFields("Week 5 ").Function = xlCount
        End If
        .PivotFields("AC SP1.2").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'AC SP1.2': " & Err.Description
            .PivotFields("AC SP1.2").Function = xlCount
        End If
        .PivotFields("Total Issue SP1.2").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Total Issue SP1.2': " & Err.Description
            .PivotFields("Total Issue SP1.2").Function = xlCount
        End If
        .PivotFields("Closed SP1.2").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Closed SP1.2': " & Err.Description
            .PivotFields("Closed SP1.2").Function = xlCount
        End If
        .PivotFields("Resolved SP1.2").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Resolved SP1.2': " & Err.Description
            .PivotFields("Resolved SP1.2").Function = xlCount
        End If
        .PivotFields("in progress SP1.2").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print "Error with 'in progress SP1.2': " & Err.Description
            .PivotFields("in progress SP1.2").Function = xlCount
        End If
        On Error GoTo 0
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .ColumnGrand = False
        .RepeatAllLabels xlRepeatLabels
        .PivotFields("Type of work ").Orientation = xlPageField

'Subtotals
        For Each pf In .RowFields
            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            pf.LayoutBlankLine = False
        Next pf
        Set sSlicerCache = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Type of work ")
        Set sSlicer = sSlicerCache.Slicers.Add(wsPivot.Name, , "Type of work ", " Type of work ", 2, 2)
            With sSlicer
                .Width = 100
                .Height = 50
                .NumberOfColumns = 1
                .RowHeight = 20
                .Style = "SlicerStyleDark3"
            End With
    End With
    wsPivot.Range("A6").Value = "Total"
    wsPivot.Range("A6").Font.Bold = True
    wsPivot.Range("C6").Value = "=Sum(C9:C50)"
    wsPivot.Range("D6").Value = "=Sum(D9:D50)"
    wsPivot.Range("E6").Value = "=Sum(E9:E50)"
    wsPivot.Range("F6").Value = "=Sum(F9:F50)"
    wsPivot.Range("G6").Value = "=Sum(G9:G50)"
    wsPivot.Range("H6").Value = "=Sum(H9:H50)"
    wsPivot.Range("I6").Value = "=Sum(I9:I50)"
    wsPivot.Range("J6").Value = "=Sum(J9:J50)"
    wsPivot.Range("K6").Value = "=Sum(K9:K50)"
    wsPivot.Range("L6").Value = "=K6/C6"
    wsPivot.Range("L6").NumberFormat = "0.00%"
    wsPivot.Range("L9:L50").NumberFormat = "0.00%"
    wsPivot.Range("M6").Value = "=Sum(M9:M50)"
    wsPivot.Range("N6").Value = "=Sum(N9:N50)"
    wsPivot.Range("O6").Value = "=Sum(O9:O50)"
    wsPivot.Range("P6").Value = "=Sum(P9:P50)"
    wsPivot.Range("Q6").Value = "=Sum(Q9:Q50)"
    wsPivot.Range("R6").Value = "=Sum(R9:R50)"
    wsPivot.Range("S6").Value = "=Sum(S9:S50)"
    wsPivot.Range("T6").Value = "=Sum(T9:T50)"
    wsPivot.Range("U6").Value = "=Sum(U9:U50)"
    wsPivot.Range("V6").Value = "=Sum(V9:V50)"
    wsPivot.Range("W6").Value = "=V6/M6"
    wsPivot.Range("W6").NumberFormat = "0.00%"
    wsPivot.Range("W9:W50").NumberFormat = "0.00%"
    'wsPivot.Range("W8").Value = "=IF(AND(GETPIVOTDATA(' Week 3 ',$A$7,'WP',B9,'Function',A9)=0,GETPIVOTDATA(' ETC JIRA SP1.2',$A$7,'WP',B9,'Function',A9)=0),0,IF(AND(GETPIVOTDATA(' Week 3 ',$A$7,'WP',B9,'Function',A9)>0,GETPIVOTDATA(' ETC JIRA SP1.2',$A$7,'WP',B9,'Function',A9)=0),'NO ETC',(GETPIVOTDATA(' Week 3 ',$A$7,'WP',B9,'Function',A9)/GETPIVOTDATA(' ETC JIRA SP1.2',$A$7,'WP',B9,'Function',A9))))"

    wsPivot.Range("X8").Value = "PE reasons"
    wsPivot.Columns("X").ColumnWidth = 50
    wsPivot.Range("Y8").Value = "Corrective Action"
    wsPivot.Columns("Y").ColumnWidth = 50
    'wsPivot.Columns("F:G").Hidden = True
    With ActiveSheet.Range("A6").Font
    .Bold = True
    .Size = 12
    End With
    ' Fields for second pivot table
    With pivotTable2
        .PivotFields("Epic Link").Orientation = xlRowField
       '.CalculatedFields.Add " Story Points", "='Story Points'*4"
       .AddDataField .PivotFields("Story Points"), " ETC", xlSum
        .PivotFields("EV").Orientation = xlDataField
        On Error Resume Next
        .PivotFields("EV").Function = xlSum
        .PivotFields("EV").Caption = "EV"
            If Err.Number <> 0 Then
                Debug.Print "Error with 'EV': " & Err.Description
                .PivotFields("EV").Function = xlCount
            End If
        On Error GoTo 0
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .ColumnGrand = False
        .SubtotalHiddenPageItems = False
        .PivotFields("Status").Orientation = xlPageField
        .PivotFields("Sprint#").Orientation = xlPageField

        'Subtotals
        For Each pf In .RowFields
            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            pf.LayoutBlankLine = False
        Next pf
        '.RefreshTable
    End With

    pivotTable.TableRange1.Columns.AutoFit
    pivotTable.TableRange1.Columns("G:J").Hidden = True
    pivotTable.TableRange1.Columns("Q:T").Hidden = True
    Range("D11").Select
    Cells.Find(What:="Sum of", After:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Cells.FindNext(After:=ActiveCell).Activate
    Cells.Replace What:="Sum of", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("P7:V8").Select
    Range("V7").Activate
    Selection.Copy
    Range("X7:Y8").Select
    Range("Y7").Activate
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A8:B8").Select
    Range("B8").Activate
    Selection.Copy
    Range("A6:B6").Select
    Range("B6").Activate
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Merge True
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    'ActiveSheet.Columns.AutoFit
    MsgBox "AC Dec_Pivot Table created successfully!", vbInformation

End Sub
```
CleanWarReport
```vb
Sub CleanWarReport()

    ActiveSheet.Shapes.Range(Array("Button 4")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 6")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 5")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 7")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 3")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 2")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Button 1")).Select
    Selection.Delete
    Selection.Cut
    ActiveSheet.Shapes.Range(Array("Group 3")).Select
    Selection.Cut
    Columns("A:E").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge
    Range("A1:P253").Select
    Range("E2").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With ActiveSheet
        Dim lastRow As Long
        Dim ws As Worksheet
        Set ws = ActiveSheet
        '.Cells.EntireRow.AutoFit
        lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
        lastRow1 = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        lastRow2 = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        lastRow3 = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
        .Columns("C").Insert Shift:=xlToRight
        .Cells(1, "C").Value = "Key Code"
        .Columns("D").Insert Shift:=xlToRight
        .Cells(1, "D").Value = "WP short"
        .Columns("E").Insert Shift:=xlToRight
        .Cells(1, "E").Value = "WP"
        .Columns("F").Insert Shift:=xlToRight
        .Cells(1, "F").Value = "Functional PE"
        .Columns("G").Insert Shift:=xlToRight
        .Cells(1, "G").Value = "Type of work"
        .Columns("I:P").Delete
         .Columns("K:M").EntireColumn.Hidden = True
        ws.Range("C2:C" & lastRow).Formula = "=MID(A2, FIND(""68"", A2), 5)"
        ws.Range("D2:D" & lastRow1).Formula = "=MID(B2, FIND(""WP"", B2), 8)"
        ws.Range("E2:E" & lastRow2).Formula = "=MID(B2, FIND(""WP"", B2), 60)"
        ws.Range("F2:F" & lastRow3).Formula = ""
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
    End With
End Sub
```