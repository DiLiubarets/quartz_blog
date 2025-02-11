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

countifs for each row 
```vb
Sub CalculateCountIfsForEachRow()

    Dim wsJira As Worksheet
    Dim wsWar As Worksheet
    Dim jiraTable As ListObject
    Dim systemsColumn As Range
    Dim warReportColumn As Range
    Dim result As Long
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheets
    Set wsJira = ThisWorkbook.Worksheets("JiraData_WeeklyPerformance_Table_1")
    Set wsWar = ThisWorkbook.Worksheets("WAR_Report_Data")
    
    ' Set the table and its columns
    Set jiraTable = wsJira.ListObjects("JiraData_WeeklyPerformance_Table_1")
    Set systemsColumn = jiraTable.ListColumns("Systems").DataBodyRange

    ' Find the last row in WAR_Report_Data column B
    lastRow = wsWar.Cells(wsWar.Rows.Count, "B").End(xlUp).Row
    
    ' Loop through each row in WAR_Report_Data column B
    For i = 1 To lastRow
        Dim criteria As Variant
        
        ' Get the criteria from column B (row i)
        criteria = wsWar.Cells(i, "B").Value
        
        ' Perform the COUNTIFS calculation
        result = Application.WorksheetFunction.CountIfs(systemsColumn, criteria)
        
        ' Output the result in column C of WAR_Report_Data (or any other column you choose)
        wsWar.Cells(i, "C").Value = result
    Next i

    ' Notify the user that the process is complete
    MsgBox "CountIfs calculation completed for all rows!"

End Sub
```

countIfs for two creteria 
```vb
Sub CountIfs_Closed_Tickets()

    Dim wsJira As Worksheet
    Dim wsWar As Worksheet
    Dim jiraTable As ListObject
    Dim systemsColumn As Range
    Dim statusColumn As Range
    Dim result As Long
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheets
    Set wsJira = ThisWorkbook.Worksheets("general_report")
    Set wsWar = ThisWorkbook.Worksheets("WAR_Report_Data")

    ' Set the table and its columns
    Set jiraTable = wsJira.ListObjects("JiraData_WeeklyPerformance_Table")
    Set systemsColumn = jiraTable.ListColumns("Systems").DataBodyRange
    Set statusColumn = jiraTable.ListColumns("Status").DataBodyRange ' Use the "Status" column from the table

    ' Find the last row in WAR_Report_Data column B
    lastRow = wsWar.Cells(wsWar.Rows.Count, "B").End(xlUp).Row

    ' Loop through each row in WAR_Report_Data column B
    For i = 1 To lastRow
        Dim criteria As Variant

        ' Get the criteria from column B (row i)
        criteria = wsWar.Cells(i, "B").Value

        ' Perform the COUNTIFS calculation
        On Error Resume Next ' Handle errors gracefully
        result = Application.WorksheetFunction.CountIfs(systemsColumn, criteria, statusColumn, "Closed")
        On Error GoTo 0 ' Turn off error handling

        ' Output the result in column O of WAR_Report_Data (or any other column you choose)
        wsWar.Cells(i, "O").Value = result
    Next i

    ' Notify the user that the process is complete
    MsgBox "CountIfs calculation completed for all rows!"

End Sub
```

picture to display 
```vb 
Sub InsertPictureAndDoOtherTasks()
    Dim ws As Worksheet
    Dim picPath As String
    Dim pic As Shape
    
    ' Set the worksheet where the picture will be inserted
    Set ws = ThisWorkbook.Sheets(1) ' Change to your desired sheet
    
    ' Provide the full path to the picture file
    picPath = "C:\Path\To\Your\Picture.jpg" ' Change to the actual path of your picture
    
    ' Insert the picture into the worksheet
    On Error Resume Next
    Set pic = ws.Shapes.AddPicture(Filename:=picPath, _
                                   LinkToFile:=msoFalse, _
                                   SaveWithDocument:=msoCTrue, _
                                   Left:=100, _
                                   Top:=100, _
                                   Width:=-1, _
                                   Height:=-1)
    If Err.Number <> 0 Then
        MsgBox "Error: Unable to insert picture. Please check the file path.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Perform other tasks in the background
    ws.Range("A1").Value = "Picture inserted!" ' Example: Add a message to cell A1
    ws.Range("B1").Value = Now ' Example: Insert the current date and time
    
    ' Format a cell as an example of background work
    With ws.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 230, 255) ' Light blue background
    End With
    
    ' Move the picture to a specific location if needed
    pic.Left = ws.Range("D5").Left
    pic.Top = ws.Range("D5").Top
    
    ' Resize the picture (optional)
    pic.LockAspectRatio = msoTrue
    pic.Width = 150 ' Set the width to 150 points
    
    ' Notify the user
    MsgBox "Picture inserted and tasks completed!", vbInformation
End Sub
```

merge weekly data
```vb
Sub CopyWeeklyData_ExplicitRanges()
    Dim ws As Worksheet
    Dim tbl1Range As Range, tbl2Range As Range
    Dim row1 As Range, row2 As Range
    Dim foundRow As Range
    Dim chargeNumber As String, employeeName As String
    Dim colNum As Long, lastCol As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("CombinedData (2)") ' Change "Sheet1" to your actual sheet name
    
    ' Define the explicit ranges for Table1 and Table2 (Update these ranges accordingly)
    Set tbl1Range = ws.Range("C2:H26") ' Change "A2:F10" to your actual Table1 range (including data, excluding headers)
    Set tbl2Range = ws.Range("N2:S119") ' Change "H2:M10" to your actual Table2 range (including data, excluding headers)
    
    ' Determine the last column for weekly data (Assumes first two columns are Charge Number & Employee Name)
    lastCol = tbl1Range.Columns.Count ' Assuming both tables have the same number of columns
    
    ' Loop through each row in Table2
    For Each row2 In tbl2Range.Rows
        chargeNumber = row2.Cells(1, 1).Value ' First column (Charge Number)
        employeeName = row2.Cells(1, 2).Value ' Second column (Employee Name)
        
        ' Search for a matching row in Table1
        For Each row1 In tbl1Range.Rows
            If row1.Cells(1, 1).Value = chargeNumber And row1.Cells(1, 2).Value = employeeName Then
                ' Match found, copy the weekly data
                For colNum = 3 To lastCol ' Weekly data columns start from the 3rd column
                    row2.Cells(1, colNum).Value = row1.Cells(1, colNum).Value
                Next colNum
                Exit For ' Exit loop once a match is found
            End If
        Next row1
    Next row2
    
    MsgBox "Weekly data copied successfully!", vbInformation
End Sub
```

fill zeros
```vb
Sub CopyWeeklyData_ExplicitRanges()
    Dim ws As Worksheet
    Dim tbl1Range As Range, tbl2Range As Range
    Dim row1 As Range, row2 As Range
    Dim chargeNumber As String, employeeName As String
    Dim colNum As Long, lastCol As Long
    Dim matchFound As Boolean
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("CombinedData (2)") ' Change to your actual sheet name
    
    ' Define the explicit ranges for Table1 and Table2 (Update these ranges accordingly)
    Set tbl1Range = ws.Range("C2:H26") ' Change to your actual Table1 range (including data, excluding headers)
    Set tbl2Range = ws.Range("N2:S119") ' Change to your actual Table2 range (including data, excluding headers)
    
    ' Determine the last column for weekly data (Assumes first two columns are Charge Number & Employee Name)
    lastCol = tbl1Range.Columns.Count ' Assuming both tables have the same number of columns
    
    ' Loop through each row in Table2
    For Each row2 In tbl2Range.Rows
        chargeNumber = row2.Cells(1, 1).Value ' First column (Charge Number)
        employeeName = row2.Cells(1, 2).Value ' Second column (Employee Name)
        
        matchFound = False ' Reset match flag for each row in Table2
        
        ' Search for a matching row in Table1
        For Each row1 In tbl1Range.Rows
            If row1.Cells(1, 1).Value = chargeNumber And row1.Cells(1, 2).Value = employeeName Then
                ' Match found, copy the weekly data
                For colNum = 3 To lastCol ' Weekly data columns start from the 3rd column
                    row2.Cells(1, colNum).Value = row1.Cells(1, colNum).Value
                Next colNum
                matchFound = True ' Set flag to indicate a match was found
                Exit For ' Exit loop once a match is found
            End If
        Next row1
        
        ' If no match was found, fill weekly data columns with zeros
        If Not matchFound Then
            For colNum = 3 To lastCol
                row2.Cells(1, colNum).Value = 0
            Next colNum
        End If
    Next row2
    
    MsgBox "Weekly data copied successfully! Missing entries filled with zeros.", vbInformation
End Sub
```


```vb
Sub Summary_with_Helios()
    Dim ws As Worksheet, wsCombined As Worksheet, wsPivot As Worksheet
    Dim rng As Range, combinedLastRow As Long
    Dim pivotCache As PivotCache, pivotTable As PivotTable
    Dim sheetNames As Variant
    Dim i As Integer

    ' Disable screen updating and calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' List of sheets to combine
    sheetNames = Array("DMA", "ASP3502A", "ASP400DC", "ASP400CL", "ASP350CS", "Helios", "HD710", "HSD-P4-CORE")

    ' Delete "CombinedData" sheet if it exists
    On Error Resume Next
    Set wsCombined = ThisWorkbook.Sheets("CombinedData")
    If Not wsCombined Is Nothing Then
        Application.DisplayAlerts = False
        wsCombined.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Create new "CombinedData" sheet
    Set wsCombined = ThisWorkbook.Sheets.Add
    wsCombined.Name = "CombinedData"

    ' Loop through sheets and copy data
    combinedLastRow = 1
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        Set rng = ws.UsedRange

        ' Copy data, ensuring headers are copied only once
        If combinedLastRow = 1 Then
            rng.Copy Destination:=wsCombined.Cells(combinedLastRow, 1)
        Else
            rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count).Copy _
                Destination:=wsCombined.Cells(combinedLastRow + 1, 1)
        End If

        ' Update last row
        combinedLastRow = wsCombined.Cells(wsCombined.Rows.Count, "A").End(xlUp).Row
    Next i

    ' Format the header row
    With wsCombined.Rows(1)
        .Font.Bold = True
        .Font.Size = 13
    End With

    ' Freeze the first three rows
    wsCombined.Rows("4:4").Select
    ActiveWindow.FreezePanes = True

    ' Delete "Summary" sheet if it exists
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Sheets("Summary")
    If Not wsPivot Is Nothing Then
        Application.DisplayAlerts = False
        wsPivot.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Create new "Summary" sheet
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "Summary"

    ' Create Pivot Table
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsCombined.UsedRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=wsPivot.Range("B2"), TableName:="CombinedPivotTable")

    ' Configure Pivot Table
    With pivotTable
        .PivotFields("Program Name").Orientation = xlColumnField
        .PivotFields("Assignee").Orientation = xlRowField
        With .PivotFields("Story Points")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
        .RowGrand = False
        .ColumnGrand = False
        .TableStyle2 = "PivotStyleMedium15"
    End With

    ' Filter out "Program Name" from Pivot Table
    With pivotTable.PivotFields("Program Name")
        .ClearAllFilters
        .PivotFilters.Add Type:=xlCaptionDoesNotEqual, Value1:="Program Name"
    End With

    ' Restore screen updating and calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```

```vb
Sub Summary_with_Helios()
    Dim ws As Worksheet, wsCombined As Worksheet, wsPivot As Worksheet
    Dim rng As Range, combinedLastRow As Long
    Dim pivotCache As PivotCache, pivotTable As PivotTable
    Dim firstSheet As Boolean

    ' Disable screen updating and calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Delete "CombinedData" sheet if it exists
    On Error Resume Next
    Set wsCombined = ThisWorkbook.Sheets("CombinedData")
    If Not wsCombined Is Nothing Then
        Application.DisplayAlerts = False
        wsCombined.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Create new "CombinedData" sheet
    Set wsCombined = ThisWorkbook.Sheets.Add
    wsCombined.Name = "CombinedData"

    ' Initialize variables
    combinedLastRow = 1
    firstSheet = True

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Skip the "CombinedData" and "Summary" sheets to avoid duplication
        If ws.Name <> "CombinedData" And ws.Name <> "Summary" Then
            Set rng = ws.UsedRange

            ' Copy data, ensuring headers are copied only once
            If firstSheet Then
                rng.Copy Destination:=wsCombined.Cells(combinedLastRow, 1)
                firstSheet = False
            Else
                rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count).Copy _
                    Destination:=wsCombined.Cells(combinedLastRow + 1, 1)
            End If

            ' Update last row
            combinedLastRow = wsCombined.Cells(wsCombined.Rows.Count, "A").End(xlUp).Row
        End If
    Next ws

    ' Format the header row
    With wsCombined.Rows(1)
        .Font.Bold = True
        .Font.Size = 13
    End With

    ' Freeze the first three rows
    wsCombined.Rows("4:4").Select
    ActiveWindow.FreezePanes = True

    ' Delete "Summary" sheet if it exists
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Sheets("Summary")
    If Not wsPivot Is Nothing Then
        Application.DisplayAlerts = False
        wsPivot.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Create new "Summary" sheet
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "Summary"

    ' Create Pivot Table
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsCombined.UsedRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=wsPivot.Range("B2"), TableName:="CombinedPivotTable")

    ' Configure Pivot Table
    With pivotTable
        .PivotFields("Program Name").Orientation = xlColumnField
        .PivotFields("Assignee").Orientation = xlRowField
        With .PivotFields("Story Points")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
        .RowGrand = False
        .ColumnGrand = False
        .TableStyle2 = "PivotStyleMedium15"
    End With

    ' Filter out "Program Name" from Pivot Table
    With pivotTable.PivotFields("Program Name")
        .ClearAllFilters
        .PivotFilters.Add Type:=xlCaptionDoesNotEqual, Value1:="Program Name"
    End With

    ' Restore screen updating and calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```


```vb
Dim lastRow As Long
With wsPivot
    ' Find the last used row in column J
    lastRow = .Cells(.Rows.Count, "J").End(xlUp).Row

    ' Clear any existing conditional formatting in column J
    .Columns("J").FormatConditions.Delete

    ' Apply conditional formatting to column J (J2:JLastRow)
    With .Range("J2:J" & lastRow).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="18")
        .Interior.Color = RGB(255, 0, 0) ' Red background
        .Font.Color = RGB(255, 255, 255) ' White text for better visibility
    End With
End With
```

```vb
Dim lastRow As Long
Dim dataRange As Range

' Find the last row in column J dynamically
lastRow = wsPivot.Cells(wsPivot.Rows.Count, "J").End(xlUp).Row

' Define the data range for column J in the Pivot Table (starting from row 5 to avoid headers)
Set dataRange = wsPivot.Range("J5:J" & lastRow)

' Clear any existing conditional formatting in column J
dataRange.FormatConditions.Delete

' Apply conditional formatting: If value >= 19, make it red
With dataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="19")
    .Interior.Color = RGB(255, 0, 0) ' Red background
    .Font.Color = RGB(255, 255, 255) ' White text for better visibility
End With
```
```vb
Sub Summary_with_Helios()
    Dim ws As Worksheet, wsCombined As Worksheet, wsPivot As Worksheet
    Dim rng As Range, combinedLastRow As Long
    Dim pivotCache As PivotCache, pivotTable As PivotTable
    Dim firstSheet As Boolean
    Dim lastRow As Long
    Dim dataRange As Range

    ' Disable screen updating and calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Delete "CombinedData" sheet if it exists
    On Error Resume Next
    Set wsCombined = ThisWorkbook.Sheets("CombinedData")
    If Not wsCombined Is Nothing Then
        Application.DisplayAlerts = False
        wsCombined.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Create new "CombinedData" sheet
    Set wsCombined = ThisWorkbook.Sheets.Add
    wsCombined.Name = "CombinedData"

    ' Initialize variables
    combinedLastRow = 1
    firstSheet = True

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Skip the "CombinedData" and "Summary" sheets to avoid duplication
        If ws.Name <> "CombinedData" And ws.Name <> "Summary" Then
            Set rng = ws.UsedRange

            ' Copy data, ensuring headers are copied only once
            If firstSheet Then
                rng.Copy Destination:=wsCombined.Cells(combinedLastRow, 1)
                firstSheet = False
            Else
                rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count).Copy _
                    Destination:=wsCombined.Cells(combinedLastRow + 1, 1)
            End If

            ' Update last row
            combinedLastRow = wsCombined.Cells(wsCombined.Rows.Count, "A").End(xlUp).Row
        End If
    Next ws

    ' Format the header row
    With wsCombined.Rows(1)
        .Font.Bold = True
        .Font.Size = 13
    End With

    ' Freeze the first three rows
    wsCombined.Activate
    wsCombined.Rows("4:4").Select
    ActiveWindow.FreezePanes = True

    ' Delete "Summary" sheet if it exists
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Sheets("Summary")
    If Not wsPivot Is Nothing Then
        Application.DisplayAlerts = False
        wsPivot.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Create new "Summary" sheet
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "Summary"

    ' Create Pivot Table
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsCombined.UsedRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=wsPivot.Range("B2"), TableName:="CombinedPivotTable")

    ' Configure Pivot Table
    With pivotTable
        .PivotFields("Program Name").Orientation = xlColumnField
        .PivotFields("Assignee").Orientation = xlRowField
        With .PivotFields("Story Points")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
        .RowGrand = False
        .ColumnGrand = False
        .TableStyle2 = "PivotStyleMedium15"
    End With

    ' Filter out "Program Name" from Pivot Table
    With pivotTable.PivotFields("Program Name")
        .ClearAllFilters
        .PivotFilters.Add Type:=xlCaptionDoesNotEqual, Value1:="Program Name"
    End With

    ' Freeze the first three rows in the Summary sheet
    wsPivot.Activate
    wsPivot.Rows("4:4").Select
    ActiveWindow.FreezePanes = True

    ' Apply Conditional Formatting to Column J in the Pivot Table
    lastRow = wsPivot.Cells(wsPivot.Rows.Count, "J").End(xlUp).Row ' Find last row dynamically
    Set dataRange = wsPivot.Range("J5:J" & lastRow) ' Define the range for column J (starting from row 5)

    ' Clear any existing conditional formatting in column J
    dataRange.FormatConditions.Delete

    ' Apply conditional formatting: If value >= 19, make it red
    With dataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="19")
        .Interior.Color = RGB(255, 0, 0) ' Red background
        .Font.Color = RGB(255, 255, 255) ' White text for better visibility
    End With

    ' Restore screen updating and calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```
```vb
' Find the last row in column J dynamically
lastRow = wsPivot.Cells(wsPivot.Rows.Count, "J").End(xlUp).Row

' Ensure the last row is valid (avoid applying formatting to empty Pivot Tables)
If lastRow < 5 Then Exit Sub

' Define the data range for column J in the Pivot Table (starting from row 5)
Set dataRange = wsPivot.Range("J5:J" & lastRow)

' Clear any existing conditional formatting in column J
dataRange.FormatConditions.Delete

' Apply conditional formatting: If value >= 19, make it red
With dataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="19")
    .Interior.Color = RGB(255, 0, 0) ' Red background
    .Font.Color = RGB(255, 255, 255) ' White text for better visibility
End With
```

```vb
Sub add_Program_Name()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sprintCol As Long
    Dim programCol As Long
    Dim foundCell As Range

    For Each ws In ThisWorkbook.Worksheets
        ' Skip the sheet named "Instructions"
        If ws.Name <> "Instructions" Then
            With ws
                ' Find the column with the header "Sprint"
                Set foundCell = .Rows(1).Find(What:="Sprint", LookAt:=xlWhole, MatchCase:=False)

                ' If "Sprint" column is found
                If Not foundCell Is Nothing Then
                    sprintCol = foundCell.Column  ' Get the column number of "Sprint"
                    programCol = sprintCol + 1    ' The new column will be inserted after "Sprint"

                    ' Insert a new column to the right of "Sprint"
                    .Columns(programCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

                    ' Set the header for the new column
                    .Cells(1, programCol).Value = "Program Name"

                    ' Find the last row in column A
                    lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

                    ' Apply the formula in the new column
                    For i = 2 To lastRow
                        .Cells(i, programCol).Formula = "=IF(ISNUMBER(FIND(""_"", N" & i & ")), TEXTBEFORE(N" & i & ", ""_""), TEXTBEFORE(N" & i & ", "" ""))"
                    Next i
                End If
            End With
        End If
    Next ws
End Sub
```

```vb
Sub RenameAllSheets()
    Dim ws As Worksheet
    Dim newName As String
    Dim i As Integer
    Dim isValid As Boolean
    Dim programCol As Long
    Dim foundCell As Range

    For Each ws In ThisWorkbook.Sheets
        ' Skip the sheet named "Instructions"
        If ws.Name <> "Instructions" Then
            With ws
                ' Find the column with the header "Program Name"
                Set foundCell = .Rows(1).Find(What:="Program Name", LookAt:=xlWhole, MatchCase:=False)

                ' If "Program Name" column is found
                If Not foundCell Is Nothing Then
                    programCol = foundCell.Column  ' Get the column number of "Program Name"
                    newName = .Cells(2, programCol).Value  ' Get the value from row 2 of that column

                    isValid = True

                    ' Check for invalid characters
                    If InStr(newName, "\") > 0 Or InStr(newName, "/") > 0 Or InStr(newName, "*") > 0 Or _
                       InStr(newName, "[") > 0 Or InStr(newName, "]") > 0 Or InStr(newName, "?") > 0 Or _
                       InStr(newName, ":") > 0 Then
                        isValid = False
                    End If

                    ' Check for duplicate names
                    For i = 1 To ThisWorkbook.Sheets.Count
                        If ThisWorkbook.Sheets(i).Name = newName Then
                            isValid = False
                            Exit For
                        End If
                    Next i

                    ' Rename the sheet if the new name is valid and unique
                    If isValid And newName <> "" Then
                        On Error Resume Next
                        ws.Name = newName
                        If Err.Number <> 0 Then
                            MsgBox "Error renaming sheet to: " & newName & vbCrLf & "Error: " & Err.Description, vbExclamation
                            Err.Clear
                        End If
                        On Error GoTo 0
                    Else
                        MsgBox "Invalid or duplicate sheet name: " & newName, vbExclamation
                    End If
                Else
                    MsgBox "Column 'Program Name' not found in sheet: " & ws.Name, vbExclamation
                End If
            End With
        End If
    Next ws
End Sub
```

```vb
Sub Summary_dynamic2()
    Dim ws As Worksheet, wsCombined As Worksheet, wsPivot As Worksheet
    Dim rng As Range, combinedLastRow As Long
    Dim pivotCache As PivotCache, pivotTable As PivotTable
    Dim firstSheet As Boolean
    Dim lastRow As Long
    Dim dataRange As Range
    Dim assigneeCol As Long
    Dim tier4Col As Long
    Dim foundCell As Range

    ' Disable screen updating and calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Delete "CombinedData" sheet if it exists
    On Error Resume Next
    Set wsCombined = ThisWorkbook.Sheets("CombinedData")
    If Not wsCombined Is Nothing Then
        Application.DisplayAlerts = False
        wsCombined.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Create new "CombinedData" sheet
    Set wsCombined = ThisWorkbook.Sheets.Add
    wsCombined.Name = "CombinedData"

    combinedLastRow = 1
    firstSheet = True

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Skip the "Instructions", "CombinedData", and "Summary" sheets to avoid duplication
        If ws.Name <> "Instructions" And ws.Name <> "CombinedData" And ws.Name <> "Summary" Then
            Set rng = ws.UsedRange
            If firstSheet Then
                wsCombined.Cells(combinedLastRow, 1).Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
                firstSheet = False
            Else
                If rng.Rows.Count > 1 Then
                    wsCombined.Cells(combinedLastRow + 1, 1).Resize(rng.Rows.Count - 1, rng.Columns.Count).Value = rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count).Value
                End If
            End If
            combinedLastRow = wsCombined.Cells(wsCombined.Rows.Count, "A").End(xlUp).Row
        End If
    Next ws

    ' Find the "Assignee" column in the "CombinedData" sheet
    Set foundCell = wsCombined.Rows(1).Find(What:="Assignee", LookAt:=xlWhole, MatchCase:=False)

    If Not foundCell Is Nothing Then
        assigneeCol = foundCell.Column  ' Get the column number of "Assignee"
        tier4Col = assigneeCol + 1      ' The new column will be inserted to the right of "Assignee"

        ' Insert a new column to the right of "Assignee"
        wsCombined.Columns(tier4Col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

        ' Set the header for the new column
        wsCombined.Cells(1, tier4Col).Value = "Tier 4"

        ' Apply the XLOOKUP formula in the new column
        wsCombined.Range(wsCombined.Cells(2, tier4Col), wsCombined.Cells(combinedLastRow, tier4Col)).Formula = _
            "=XLOOKUP(" & wsCombined.Cells(2, assigneeCol).Address(False, False) & ",Instructions!W:W,Instructions!AB:AB, ""NON SATCOM"")"
    Else
        MsgBox "Column 'Assignee' not found in CombinedData sheet.", vbExclamation
    End If

    ' Delete "Summary" sheet if it exists
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Sheets("Summary")
    If Not wsPivot Is Nothing Then
        Application.DisplayAlerts = False
        wsPivot.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Create new "Summary" sheet
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "Summary"

    ' Create Pivot Table
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsCombined.UsedRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=wsPivot.Range("B2"), TableName:="CombinedPivotTable")

    With pivotTable
        .PivotFields("Program Name").Orientation = xlColumnField
        .PivotFields("Tier 4").Orientation = xlRowField
        .PivotFields("Assignee").Orientation = xlRowField
        With .PivotFields("Story Points")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0.00"
        End With
        .ColumnGrand = False
        .RowGrand = True
        .TableStyle2 = "PivotStyleMedium15"
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
    End With

    ' Filter out "Program Name" from Pivot Table
    With pivotTable.PivotFields("Program Name")
        .ClearAllFilters
        .PivotFilters.Add Type:=xlCaptionDoesNotEqual, Value1:="Program Name"
    End With

    ' Freeze the first three rows in the Summary sheet
    wsPivot.Activate
    wsPivot.Rows("4:4").Select
    ActiveWindow.FreezePanes = True

    ' Restore screen updating and calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```