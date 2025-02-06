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


extra for RTP
