**JiraRaw_Cleaning**: Something that I do way too much manually.
```vb
Sub JiraRaw_Cleaning()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim deleteRange As Range

    ' Set the active worksheet to a variable for better readability
    Set ws = ActiveSheet
    ws.Columns.AutoFit

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
        ' Find and delete rows containing "Generated at"
        Set rng = ws.UsedRange.Find(What:="Generated at", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        If Not rng Is Nothing Then
            Set deleteRange = rng
            Do
                If deleteRange Is Nothing Then
                    Set deleteRange = rng
                Else
                    Set deleteRange = Union(deleteRange, rng)
                End If
                Set rng = ws.UsedRange.FindNext(rng)
            Loop While Not rng Is Nothing And rng.Address <> deleteRange.Cells(1, 1).Address
            ' Delete the rows
            deleteRange.EntireRow.Delete
        End If

        ' AutoFit columns and rows
        .Cells(1, 1).Select
        .UsedRange.Columns.AutoFit
        .UsedRange.Rows.AutoFit
    End With
End Sub

```
DarkMode in Excel
```vb
Sub EnableDarkMode()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.Interior.Color = RGB(43, 43, 43) ' Dark background color
        ws.Cells.Font.Color = RGB(255, 255, 255) ' White font color
    Next ws
End Sub

Sub DisableDarkMode()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.Interior.ColorIndex = xlNone ' Reset to no fill
        ws.Cells.Font.ColorIndex = xlAutomatic ' Reset to automatic font color
    Next ws
End Sub

```
AutoFitAll
```vb
Sub AutoFitAll()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.UsedRange.Columns.AutoFit
        ws.UsedRange.Rows.AutoFit
    Next ws
End Sub
```

```vb
Sub CreateDropdownListFromOtherSheet()
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim dropdownRange As Range
    Dim targetCell As Range
    
    ' Set the source sheet (where the dropdown list values are located)
    Set sourceSheet = ThisWorkbook.Sheets("Sheet2")
    
    ' Set the target sheet (where the dropdown will appear)
    Set targetSheet = ThisWorkbook.Sheets("Sheet1")
    
    ' Define the range containing the dropdown list values in the source sheet
    Set dropdownRange = sourceSheet.Range("A1:A5")
    
    ' Define the target cell where the dropdown will appear in the target sheet
    Set targetCell = targetSheet.Range("B1")
    
    ' Clear any existing validation on the target cell
    targetCell.Validation.Delete
    
    ' Add data validation to the target cell
    With targetCell.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="='" & sourceSheet.Name & "'!" & dropdownRange.Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    MsgBox "Dropdown list created successfully in cell " & targetCell.Address, vbInformation
End Sub
```

```vb
Sub StandardizeRowStyles()
    Dim ws As Worksheet
    Dim row As Range
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through all rows in the worksheet
        For Each row In ws.Rows
            ' Apply consistent row height
            row.RowHeight = 20
            
            ' Apply consistent font style and size
            With row.Font
                .Name = "Calibri" ' Set font name
                .Size = 12        ' Set font size
                .Bold = False     ' Set font to not bold
            End With
            
            ' Apply consistent alignment
            With row
                .HorizontalAlignment = xlCenter ' Center align horizontally
                .VerticalAlignment = xlCenter   ' Center align vertically
            End With
            
            ' Optional: Clear any existing conditional formatting or custom styles
            row.ClearFormats
        Next row
    Next ws
    
    MsgBox "Row styles have been standardized across all sheets!"
End Sub
```

```vb 
Sub ShowAllSheetNames()
    Dim ws As Worksheet
    Dim sheetNames As String
    
    ' Loop through all sheets and concatenate their names
    For Each ws In ThisWorkbook.Sheets
        sheetNames = sheetNames & ws.Name & vbNewLine
    Next ws
    
    ' Display the list of sheet names
    MsgBox "The workbook contains the following sheets:" & vbNewLine & sheetNames, vbInformation, "Sheet Names"
End Sub
```

sumifs
```vb
Sub CalculateSumIfs()

    Dim ws As Worksheet
    Dim jiraTable As ListObject
    Dim systemsColumn As Range
    Dim storyPointsColumn As Range
    Dim warReportColumn As Range
    Dim result As Double
    Dim criteria As Variant
    
    ' Set the worksheet where the data resides
    Set ws = ThisWorkbook.Worksheets("JiraData_WeeklyPerformance_Table_1")
    
    ' Set the table and its columns
    Set jiraTable = ws.ListObjects("JiraData_WeeklyPerformance_Table_1")
    Set storyPointsColumn = jiraTable.ListColumns("Story Points").DataBodyRange
    Set systemsColumn = jiraTable.ListColumns("Systems").DataBodyRange

    ' Set the criteria range from WAR_Report_Data worksheet
    Set warReportColumn = ThisWorkbook.Worksheets("WAR_Report_Data").Range("B:B")
    
    ' Define the criteria (you might need to adjust this based on your logic)
    criteria = warReportColumn.Cells(1, 1).Value ' Example: using the first value in column B as criteria
    
    ' Perform the SUMIFS calculation
    result = Application.WorksheetFunction.SumIfs(storyPointsColumn, systemsColumn, criteria) * 4
    
    ' Output the result (you can adjust where this result is displayed or stored)
    MsgBox "The calculated result is: " & result

End Sub
```

sumifs per row
```vb
Sub CalculateSumIfsForEachRow()

    Dim wsJira As Worksheet
    Dim wsWarReport As Worksheet
    Dim jiraTable As ListObject
    Dim systemsColumn As Range
    Dim storyPointsColumn As Range
    Dim criteriaColumn As Range
    Dim resultColumn As Range
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim criteriaValue As Variant
    Dim result As Double
    
    ' Set the worksheets
    Set wsJira = ThisWorkbook.Worksheets("JiraData_WeeklyPerformance_Table_1")
    Set wsWarReport = ThisWorkbook.Worksheets("WAR_Report_Data")
    
    ' Set the Jira table and its columns
    Set jiraTable = wsJira.ListObjects("JiraData_WeeklyPerformance_Table_1")
    Set storyPointsColumn = jiraTable.ListColumns("Story Points").DataBodyRange
    Set systemsColumn = jiraTable.ListColumns("Systems").DataBodyRange
    
    ' Determine the columns in WAR_Report_Data
    Set criteriaColumn = wsWarReport.Range("B:B") ' Column B contains the criteria
    Set resultColumn = wsWarReport.Range("C:C")  ' Column C will store the results (adjust as needed)
    
    ' Find the last row in the criteria column
    lastRow = wsWarReport.Cells(wsWarReport.Rows.Count, criteriaColumn.Column).End(xlUp).Row
    
    ' Loop through each row in WAR_Report_Data
    For rowIndex = 2 To lastRow ' Start at row 2 to skip headers (adjust if no headers)
        ' Get the criteria value from the criteria column
        criteriaValue = wsWarReport.Cells(rowIndex, criteriaColumn.Column).Value
        
        ' Perform the SUMIFS calculation
        If Not IsEmpty(criteriaValue) Then
            result = Application.WorksheetFunction.SumIfs(storyPointsColumn, systemsColumn, criteriaValue) * 4
        Else
            result = 0 ' Handle empty criteria
        End If
        
        ' Place the result in the corresponding row of the result column
        wsWarReport.Cells(rowIndex, resultColumn.Column).Value = result
    Next rowIndex
    
    ' Notify the user that the macro is complete
    MsgBox "SUMIFS calculations are complete and results are stored in column " & resultColumn.Column, vbInformation

End Sub
```