Create Pivot from DataDump

```vb
Sub CreatePivot_DumpData()

    On Error Resume Next
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As pivotCache
    Dim pivotTable As pivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim pSlicersCaches As SlicerCaches
    Dim sSlicerCache1 As SlicerCache
    Dim sSlicerCache2 As SlicerCache
    Dim sSlicerCache3 As SlicerCache
    Dim sSlicerCache4 As SlicerCache
    Dim sSlicer1 As Slicer
    Dim sSlicer2 As Slicer
    Dim sSlicer3 As Slicer
    Dim sSlicer4 As Slicer
    Dim timelineSlicer As Slicer
    Dim timelineCache As SlicerCache
    Dim ws As Worksheet
    Dim currentYear As Integer
    Dim currentMonth As Integer
    Dim startDate As Date
    Dim endDate As Date
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set wsData = ThisWorkbook.Worksheets("DataDump")
    Set pivotRange = wsData.Range("A1").CurrentRegion

    ' Create or clear pivot sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PivotTable").Delete
    On Error GoTo 0
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "PivotTable"
    Application.DisplayAlerts = True

' Create pivot table
    Set pivotDestination = wsPivot.Range("E10")
    Set pivotCache = ThisWorkbook.PivotCaches.create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTable")

    With pivotTable
        .PivotFields("WorkCenter").Orientation = xlRowField
        .PivotFields("FiscalMonth").Orientation = xlColumnField
        .PivotFields("Value").Orientation = xlDataField
    End With

    ' Create Slicers
    On Error Resume Next
    ' First Slicer
    If Err.Number = 0 Then
        Set sSlicerCache1 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Type")
        If Err.Number = 0 Then
            Set sSlicer1 = sSlicerCache1.Slicers.Add(wsPivot.Name, , "Type", "Type", 5, 10)
            With sSlicer1
                .Width = 180
                .Height = 58
                .NumberOfColumns = 2
                .RowHeight = 20
            End With
        sSlicerCache1.VisibleSlicerItemsList = Array("[Type].[Labor]")
        Else
            MsgBox "Error creating first slicer: " & Err.Description
        End If
    End If

    ' Second Slicer
    Err.Clear
    If Err.Number = 0 Then
        Set sSlicerCache2 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Current vs Previous Data")
        If Err.Number = 0 Then
            Set sSlicer2 = sSlicerCache2.Slicers.Add(wsPivot.Name, , "Current/Previous Data", "Current vs Previous Data", 68, 10)

            With sSlicer2
                .Width = 180
                .Height = 58
                .NumberOfColumns = 2
                .RowHeight = 20
            End With
        Else
            MsgBox "Error creating second slicer: " & Err.Description
        End If
    End If

    ' Third Slicer
    Err.Clear
    If Err.Number = 0 Then
        Set sSlicerCache3 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "EAC Type")
        If Err.Number = 0 Then
            Set sSlicer3 = sSlicerCache3.Slicers.Add(wsPivot.Name, , "EAC Type", "EAC Type", 130, 10)
            With sSlicer3
                .Width = 180
                .Height = 58
                .NumberOfColumns = 2
                .RowHeight = 20
            End With
        Else
            MsgBox "Error creating third slicer: " & Err.Description
        End If
    End If

    ' Forth Slicer
    Err.Clear
    If Err.Number = 0 Then
        Set sSlicerCache4 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Project")
        If Err.Number = 0 Then
            Set sSlicer4 = sSlicerCache4.Slicers.Add(wsPivot.Name, , "Project", "Project", 193, 10)
            With sSlicer4
                .Width = 180
                .Height = 58
                .NumberOfColumns = 2
                .RowHeight = 20
            End With
        Else
            MsgBox "Error creating third slicer: " & Err.Description
        End If
    End If
    With wb.SlicerCaches("Slicer_Type")
        .SlicerItems("Labor").Selected = True
        .SlicerItems("NonLabor").Selected = False
    End With
    With wb.SlicerCaches("Slicer_Current_vs_Previous_Data")
        .SlicerItems("Current Month").Selected = True
        .SlicerItems("Previous Month").Selected = False
    End With
    With wb.SlicerCaches("Slicer_EAC_Type")
        .SlicerItems("ETC").Selected = True
        .SlicerItems("ACTUALS").Selected = False
    End With
    On Error GoTo 0


    ' Set references to workbook, worksheet, and PivotTable
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("PivotTable")
    Set pivotTable = ws.PivotTables("MyPivotTable")
    
    ' Create the Slicer Cache for Timeline
    Set timelineCache = wb.SlicerCaches.Add2(pivotTable, "FiscalMonth", , xlTimeline)
    timelineCache.Slicers.Add ActiveSheet, , "FiscalMonth 1", "FiscalMonth", 10, 200, 575, 108

    ' Set Timeline Slicer to Current Month and Next Month
    currentYear = Year(Date)
    currentMonth = Month(Date)
    startDate = DateSerial(currentYear, currentMonth + 1, 1)
    endDate = DateSerial(currentYear, 12, 31) - 1
    With timelineCache.TimelineState
        .SetFilterDateRange startDate, endDate
        '.ClearAllFilters
    End With
    'MsgBox "Pivot Table with Slicers created successfully!", vbInformation

End Sub

```

Start Time line from current month +1 
```vb
Sub CreateTimelineAndSetToCurrentYearMonth()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pt As pivotTable
    Dim timelineCache As SlicerCache
    Dim currentYear As Integer
    Dim currentMonth As Integer
    Dim startDate As Date
    Dim endDate As Date

    ' Set references to workbook, worksheet, and PivotTable
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("PivotTable")
    Set pt = ws.PivotTables("MyPivotTable")
    
    ' Create the Slicer Cache for Timeline
    Set timelineCache = wb.SlicerCaches.Add2(pt, "FiscalMonth", , xlTimeline)
    timelineCache.Slicers.Add ActiveSheet, , "FiscalMonth 1", "FiscalMonth", 10, 200, 575, 108

    ' Set Timeline Slicer to Current Month and Next Month
    currentYear = Year(Date)
    currentMonth = Month(Date)
    startDate = DateSerial(currentYear, currentMonth + 1, 1)
    endDate = DateSerial(currentYear, 12, 31) - 1
    With timelineCache.TimelineState
        .SetFilterDateRange startDate, endDate
        '.ClearAllFilters
    End With

End Sub
```


Demand Sheet 
```vb
Sub Demand_working_sheet()
    Dim ws As Worksheet
    Dim cell As Range
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Demand")
    ws.Columns("M").Insert Shift:=xlToRight
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    ws.Range("M2:M" & lastRow).Formula = "=MID(L2, 5, LEN(L2) - 4)"
    ws.Columns("N").Insert Shift:=xlToRight
    ws.Range("N2:N" & lastRow).Formula = "=XLOOKUP(M2:M233,PivotTable!$E$12:$E$60,PivotTable!$E$12:$E$60)"

    With ws.Range("N1:N" & ws.Cells(ws.Rows.Count, "N").End(xlUp).Row)
        .AutoFilter Field:=1, Criteria1:="<>" & "#N/A"
    End With

     ' Loop through each cell in column S
    For Each cell In ws.Range("S2:S" & lastRow)
        duplicateCount = Application.WorksheetFunction.CountIf(ws.Range("S2:S" & lastRow), cell.Value)
        If duplicateCount > 1 Then
            cell.Interior.Color = vbRed
        End If
    Next cell

    'hide col W&C
    ws.Columns("W").EntireColumn.Hidden = True
    ws.Columns("C").EntireColumn.Hidden = True
    
    'Module5.ConvertDates

    For Each cell In ws.Range("Y1:BH1")
        If IsDate("1-" & Left(cell.Value, 3) & "-20" & Right(cell.Value, 2)) Then
            cell.Value = DateValue("1-" & Left(cell.Value, 3) & "-20" & Right(cell.Value, 2))
            cell.NumberFormat = "mmm-yy"
        End If
    Next cell
End Sub
```


Converting Date to the same format with DumpData
``` vb
Sub ConvertDates()

    Dim cell As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Demand")
    For Each cell In ws.Range("Y1:BH1")
        If IsDate("1-" & Left(cell.Value, 3) & "-20" & Right(cell.Value, 2)) Then
            cell.Value = DateValue("1-" & Left(cell.Value, 3) & "-20" & Right(cell.Value, 2))
            cell.NumberFormat = "mmm-yy"
        End If
    Next cell
End Sub
```

```vb
Sub SyncSlicerWithFilter()
    Dim wsPivot As Worksheet
    Dim wsData As Worksheet
    Dim slicerCache As SlicerCache
    Dim slicerItem As SlicerItem
    Dim selectedItems As Collection
    Dim filterCriteria As String
    Dim i As Long
 
    Set wsPivot = ThisWorkbook.Sheets("Sheet1") 
    Set wsData = ThisWorkbook.Sheets("Sheet2")
    Set slicerCache = ThisWorkbook.SlicerCaches("Slicer_Project")

    Set selectedItems = New Collection
    For Each slicerItem In slicerCache.SlicerItems
        If slicerItem.Selected Then
            selectedItems.Add slicerItem.Name
        End If
    Next slicerItem

    If selectedItems.Count > 0 Then
        filterCriteria = ""
        For i = 1 To selectedItems.Count
            filterCriteria = filterCriteria & selectedItems(i) & ","
        Next i
        filterCriteria = Left(filterCriteria, Len(filterCriteria) - 1)
    Else
        MsgBox "No slicer items are selected. Please select at least one item in the slicer.", vbExclamation
        Exit Sub
    End If

    wsData.AutoFilterMode = False 
    wsData.Range("D:D").AutoFilter Field:=1, Criteria1:=Split(filterCriteria, ","), Operator:=xlFilterValues

    MsgBox "Filter applied successfully based on slicer selection!", vbInformation
End Sub
```

View all slicers
```vb
Sub ListAllSlicerCaches()
    Dim slicerCache As SlicerCache
    Dim msg As String

    For Each slicerCache In ThisWorkbook.SlicerCaches
        msg = msg & slicerCache.Name & vbNewLine
    Next slicerCache

    If msg = "" Then
        MsgBox "No slicers found in this workbook.", vbExclamation
    Else
        MsgBox "Slicer Caches in This Workbook:" & vbNewLine & msg, vbInformation
    End If
End Sub
```

Find in pivot table (two arguments)
```vb
Sub GetPivotTableValue()
    Dim wsPivot As Worksheet
    Dim wsOutput As Worksheet
    Dim pt As PivotTable
    Dim projectName As String
    Dim monthName As String
    Dim result As Variant
    
    Set wsPivot = ThisWorkbook.Sheets("PivotTableSheet") 
    Set wsOutput = ThisWorkbook.Sheets("OutputSheet") 
    Set pt = wsPivot.PivotTables("PivotTable1") 
    projectName = wsOutput.Range("A1").Value 
    monthName = wsOutput.Range("B1").Value 
    On Error Resume Next
    result = pt.GetPivotData( _
        DataField:=pt.DataFields(1).Name, _
        PivotTableField1:="Project Name", PivotItem1:=projectName, _
        PivotTableField2:="Month", PivotItem2:=monthName)
    On Error GoTo 0

    If IsError(result) Then
        MsgBox "No value found for the specified project and month.", vbExclamation
    Else

        wsOutput.Range("C1").Value = result 
        MsgBox "Value found: " & result, vbInformation
    End If
End Sub
```

Name all pivot fields
```vb
Sub ListPivotFieldNames()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim outputSheet As Worksheet
    Dim outputRow As Long
    
    Set ws = ThisWorkbook.Sheets("PivotTableSheet") 
    Set pt = ws.PivotTables("PivotTable1") 
    Set outputSheet = ThisWorkbook.Sheets("OutputSheet") 
    outputRow = 1 
    outputSheet.Cells.Clear
    
    For Each pf In pt.PivotFields
        outputSheet.Cells(outputRow, 1).Value = pf.Name 
        outputRow = outputRow + 1 
    Next pf
    
    MsgBox "Pivot field names have been listed in the output sheet.", vbInformation
End Sub
```

Debug code
```vb
Sub GetPivotTableValue()
    Dim wsPivot As Worksheet
    Dim wsOutput As Worksheet
    Dim pt As PivotTable
    Dim projectName As String
    Dim monthName As String
    Dim result As Variant
    
    ' Define the sheets
    Set wsPivot = ThisWorkbook.Sheets("PivotTableSheet")
    Set wsOutput = ThisWorkbook.Sheets("OutputSheet") 
    Set pt = wsPivot.PivotTables("PivotTable1") 
    projectName = wsOutput.Range("A1").Value 
    monthName = wsOutput.Range("B1").Value 
    If Trim(projectName) = "" Or Trim(monthName) = "" Then
        MsgBox "Please make sure both Project Name (A1) and Month (B1) are filled in.", vbExclamation
        Exit Sub
    End If
  
    If pt.DataFields.Count = 0 Then
        MsgBox "The pivot table does not have any data fields.", vbExclamation
        Exit Sub
    End If
 
    On Error Resume Next
    result = pt.GetPivotData( _
        DataField:=pt.DataFields(1).Name, _
        Field1:="Project Name", Item1:=projectName, _
        Field2:="Month", Item2:=monthName)
    On Error GoTo 0

    If IsError(result) Then
        MsgBox "No value found for the specified Project Name and Month. Please check your inputs or the pivot table structure.", vbExclamation
    ElseIf IsEmpty(result) Then
        MsgBox "The value found is empty. Please check if the combination of Project Name and Month exists in the pivot table.", vbExclamation
    Else
        wsOutput.Range("C1").Value = result 
        MsgBox "Value found: " & result, vbInformation
    End If
End Sub
```

Loop through rows
```vb
Sub GetPivotDataForAllRowsAndColumns()
    Dim wsInput As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim workCenter As String
    Dim fiscalMonth As Date
    Dim result As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim currentRow As Long
    Dim currentCol As Long

    Set wsInput = ThisWorkbook.Sheets("InputSheet") 
    Set wsPivot = ThisWorkbook.Sheets("PivotTableSheet") 
    Set pt = wsPivot.PivotTables("PivotTable1") 
    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row 
    lastCol = wsInput.Cells(1, wsInput.Columns.Count).End(xlToLeft).Column 
    For currentRow = 2 To lastRow
        workCenter = Trim(wsInput.Cells(currentRow, 1).Value)

        For currentCol = 2 To lastCol
            If IsDate(wsInput.Cells(1, currentCol).Value) Then
                fiscalMonth = DateSerial(Year(wsInput.Cells(1, currentCol).Value), Month(wsInput.Cells(1, currentCol).Value), Day(wsInput.Cells(1, currentCol).Value))
            Else

                MsgBox "Invalid date in column " & currentCol & ". Skipping.", vbExclamation
                GoTo SkipColumn
            End If

            On Error Resume Next
            result = pt.GetPivotData( _
                DataField:="Sum of Value", _
                Field1:="WorkCenter", Item1:=workCenter, _
                Field2:="FiscalMonth", Item2:=fiscalMonth)
            On Error GoTo 0
            If IsError(result) Then
                wsInput.Cells(currentRow, currentCol).Value = "N/A" 
            Else
                wsInput.Cells(currentRow, currentCol).Value = result 
            End If

SkipColumn:
        Next currentCol
    Next currentRow

    MsgBox "Data retrieval complete!", vbInformation
End Sub
```

Process visible only, no message
```vb
Sub GetPivotDataForFilteredRowsAndColumns()
    Dim wsInput As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim workCenter As String
    Dim fiscalMonth As Date
    Dim result As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim currentRow As Long
    Dim currentCol As Long

    Set wsInput = ThisWorkbook.Sheets("InputSheet") 
    Set wsPivot = ThisWorkbook.Sheets("PivotTableSheet") 
    Set pt = wsPivot.PivotTables("PivotTable1") 
    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row 
    lastCol = wsInput.Cells(1, wsInput.Columns.Count).End(xlToLeft).Column 
    For currentRow = 2 To lastRow
        If Not wsInput.Rows(currentRow).EntireRow.Hidden Then
            workCenter = Trim(wsInput.Cells(currentRow, 1).Value)

            For currentCol = 2 To lastCol
             
                If IsDate(wsInput.Cells(1, currentCol).Value) Then
                    fiscalMonth = DateSerial(Year(wsInput.Cells(1, currentCol).Value), Month(wsInput.Cells(1, currentCol).Value), Day(wsInput.Cells(1, currentCol).Value))

                    On Error Resume Next
                    result = pt.GetPivotData( _
                        DataField:="Sum of Value", _
                        Field1:="WorkCenter", Item1:=workCenter, _
                        Field2:="FiscalMonth", Item2:=fiscalMonth)
                    On Error GoTo 0
                    If IsError(result) Then
                        wsInput.Cells(currentRow, currentCol).Value = "N/A" 
                    Else
                        wsInput.Cells(currentRow, currentCol).Value = result 
                    End If
                    result = Empty
                End If
            Next currentCol
        End If
    Next currentRow

    MsgBox "Data retrieval complete for filtered rows!", vbInformation
End Sub
```

Handle n/a
```vb
' Get the WorkCenter from column N (column 14)
If IsError(wsInput.Cells(currentRow, 14).Value) Then
    Debug.Print "Skipping row " & currentRow & " because WorkCenter contains an error (#N/A)."
    GoTo NextRow
ElseIf wsInput.Cells(currentRow, 14).Value = "" Then
    Debug.Print "Skipping row " & currentRow & " because WorkCenter is blank."
    GoTo NextRow
Else
    workCenter = Trim(wsInput.Cells(currentRow, 14).Value)
End If
```

Sync with slicer and delete
```vb
Sub DeleteRowsNotMatchingSlicerSelection()
    Dim wsPivot As Worksheet
    Dim wsData As Worksheet
    Dim slicerCache As SlicerCache
    Dim slicerItem As SlicerItem
    Dim selectedItems As Collection
    Dim selectedItemDict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range

    ' Set references to the sheets and slicer
    Set wsPivot = ThisWorkbook.Sheets("Sheet1")
    Set wsData = ThisWorkbook.Sheets("Sheet2")
    Set slicerCache = ThisWorkbook.SlicerCaches("Slicer_Project")

    ' Collect selected slicer items
    Set selectedItems = New Collection
    Set selectedItemDict = CreateObject("Scripting.Dictionary") ' Use a dictionary for faster lookups

    For Each slicerItem In slicerCache.SlicerItems
        If slicerItem.Selected Then
            selectedItems.Add slicerItem.Name
            selectedItemDict.Add slicerItem.Name, True
        End If
    Next slicerItem

    ' Check if there are selected slicer items
    If selectedItems.Count = 0 Then
        MsgBox "No slicer items are selected. Please select at least one item in the slicer.", vbExclamation
        Exit Sub
    End If

    ' Find the last row in the data sheet
    lastRow = wsData.Cells(wsData.Rows.Count, "D").End(xlUp).Row

    ' Loop through the data and delete rows that don't match the slicer selection
    Application.ScreenUpdating = False
    For i = lastRow To 2 Step -1 ' Start from the bottom row to avoid skipping rows
        If Not selectedItemDict.exists(wsData.Cells(i, "D").Value) Then
            wsData.Rows(i).Delete
        End If
    Next i
    Application.ScreenUpdating = True

    MsgBox "Rows not matching the slicer selection have been deleted successfully!", vbInformation
End Sub
```

Simplifed get pivot data
```vb
Sub GetPivotDataForAllRowsAndColumns()
    Dim wsInput As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim workCenter As String
    Dim fiscalMonth As Date
    Dim result As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim currentRow As Long
    Dim currentCol As Long

    ' Set references to the input and pivot table sheets
    Set wsInput = ThisWorkbook.Sheets("InputSheet")
    Set wsPivot = ThisWorkbook.Sheets("PivotTableSheet")
    Set pt = wsPivot.PivotTables("PivotTable1")

    ' Determine the last row and column in the input sheet
    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row
    lastCol = wsInput.Cells(1, wsInput.Columns.Count).End(xlToLeft).Column

    ' Loop through each row and column to retrieve pivot table data
    For currentRow = 2 To lastRow
        workCenter = Trim(wsInput.Cells(currentRow, 1).Value)

        For currentCol = 2 To lastCol
            If IsDate(wsInput.Cells(1, currentCol).Value) Then
                fiscalMonth = DateSerial(Year(wsInput.Cells(1, currentCol).Value), _
                                         Month(wsInput.Cells(1, currentCol).Value), _
                                         Day(wsInput.Cells(1, currentCol).Value))

                ' Attempt to retrieve data from the pivot table
                On Error Resume Next
                result = pt.GetPivotData( _
                    DataField:="Sum of Value", _
                    Field1:="WorkCenter", Item1:=workCenter, _
                    Field2:="FiscalMonth", Item2:=fiscalMonth)
                On Error GoTo 0

                ' Write the result to the cell or mark it as "N/A" if not found
                If IsError(result) Then
                    wsInput.Cells(currentRow, currentCol).Value = "N/A"
                Else
                    wsInput.Cells(currentRow, currentCol).Value = result
                End If

                result = Empty
            End If
        Next currentCol
    Next currentRow

    MsgBox "Data retrieval complete for all rows!", vbInformation
End Sub
```


demand 
```vb
Sub Demand_working_sheet()
    Dim wsPivot As Worksheet
    Dim wsData As Worksheet
    Dim ws As Worksheet ' Separate variable for looping through sheets
    Dim cell As Range
    Dim slicerCache As SlicerCache
    Dim slicerItem As SlicerItem
    Dim selectedItems As Collection
    Dim i As Long
    Dim filterCriteria As String
    Dim pt As PivotTable
    Dim LastRow As Long
    Dim workCenter As String
    Dim fiscalMonth As Date
    Dim result As Variant
    Dim lastCol As Long
    Dim currentRow As Long
    Dim currentCol As Long

    ' Set the worksheet references
    Set wsPivot = ThisWorkbook.Sheets("PivotTable")
    Set wsData = ThisWorkbook.Sheets("Demand")

    ' Insert a new column and populate it
    wsData.Columns("M").Insert Shift:=xlToRight
    wsData.Cells(1, 13).Value = "NEED TO BE DELETED LATER"
    LastRow = wsData.Cells(wsData.Rows.Count, "L").End(xlUp).Row
    wsData.Range("M2:M" & LastRow).Formula = "=MID(L2, 5, LEN(L2) - 4)"

    ' Hide columns W & C
    wsData.Columns("W").EntireColumn.Hidden = True
    wsData.Columns("C").EntireColumn.Hidden = True

    ' Convert date formats in range Y1:BH1
    For Each cell In wsData.Range("Y1:BH1")
        If IsDate("1-" & Left(cell.Value, 3) & "-20" & Right(cell.Value, 2)) Then
            cell.Value = DateValue("1-" & Left(cell.Value, 3) & "-20" & Right(cell.Value, 2))
            cell.NumberFormat = "mmmyy"
        End If
    Next cell

    ' Sync slicer with filter
    Set slicerCache = ThisWorkbook.SlicerCaches("Slicer_Project")
    Set selectedItems = New Collection
    For Each slicerItem In slicerCache.SlicerItems
        If slicerItem.Selected Then
            selectedItems.Add slicerItem.Name
        End If
    Next slicerItem

    ' Build filter criteria
    If selectedItems.Count > 0 Then
        filterCriteria = ""
        For i = 1 To selectedItems.Count
            filterCriteria = filterCriteria & selectedItems(i) & ","
        Next i
        filterCriteria = Left(filterCriteria, Len(filterCriteria) - 1)
    Else
        MsgBox "No slicer items are selected. Please select at least one item in the slicer.", vbExclamation
        Exit Sub
    End If

    ' Apply filter to column D
    wsData.Range("D:D").AutoFilter Field:=1, Criteria1:=Split(filterCriteria, ","), Operator:=xlFilterValues

    ' Corrected loop to disable AutoFilter
    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
    Next ws

    ' GetPivot_Data
    Set pt = wsPivot.PivotTables("MyPivotTable")
    LastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    For currentRow = 2 To LastRow
        If Not IsError(wsData.Cells(currentRow, 14).Value) Then
            workCenter = Trim(wsData.Cells(currentRow, 14).Value)
            For currentCol = 2 To lastCol
                fiscalMonth2 = wsData.Cells(1, currentCol).Value
                If IsDate(wsData.Cells(1, currentCol).Value) Then
                    fiscalMonth = DateSerial(Year(fiscalMonth2), Month(fiscalMonth2), Day(fiscalMonth2))
                    On Error Resume Next
                    result = pt.GetPivotData( _
                        DataField:="Sum of Value", _
                        Field1:="WorkCenter", Item1:=workCenter, _
                        Field2:="FiscalMonth", Item2:=fiscalMonth)
                    On Error GoTo 0
                    If IsError(result) Then
                        wsData.Cells(currentRow, currentCol).Value = "#N/A"
                    Else
                        wsData.Cells(currentRow, currentCol).Value = result
                    End If
                    result = Empty
                End If
            Next currentCol
        End If
    Next currentRow

    ' Find the last row with data in column S
    LastRow = wsData.Cells(wsData.Rows.Count, "S").End(xlUp).Row

    MsgBox "Demand working sheet updated successfully!", vbInformation
End Sub
```


check for extra 
```vb
Sub Find_Extra_Row_Labels()
    Dim wsPivot As Worksheet
    Dim wsData As Worksheet
    Dim wsExtra As Worksheet
    Dim pt As PivotTable
    Dim pivotRange As Range
    Dim pivotRowLabels As Object
    Dim dataRowLabels As Object
    Dim cell As Range
    Dim lastRow As Long
    Dim extraRow As Long
    Dim key As Variant
    
    ' Set worksheet references
    Set wsPivot = ThisWorkbook.Sheets("PivotTable") ' Pivot table source
    Set wsData = ThisWorkbook.Sheets("Demand") ' Demand data source
    
    ' Create or clear the "Extra" sheet
    On Error Resume Next
    Set wsExtra = ThisWorkbook.Sheets("Extra")
    If wsExtra Is Nothing Then
        Set wsExtra = ThisWorkbook.Sheets.Add
        wsExtra.Name = "Extra"
    Else
        wsExtra.Cells.Clear ' Clear previous data
    End If
    On Error GoTo 0
    
    ' Set PivotTable reference
    Set pt = wsPivot.PivotTables("MyPivotTable") ' Change to your actual PivotTable name
    
    ' Define dictionaries for storing row labels
    Set pivotRowLabels = CreateObject("Scripting.Dictionary")
    Set dataRowLabels = CreateObject("Scripting.Dictionary")
    
    ' Get PivotTable row labels (assuming they are in the first column of the PivotTable)
    Set pivotRange = pt.TableRange1.Columns(1) ' First column of PivotTable
    
    For Each cell In pivotRange.Cells
        If cell.Row > pt.TableRange1.Row Then ' Avoid header row
            pivotRowLabels(cell.Value) = True
        End If
    Next cell
    
    ' Get unique values from column S in Demand sheet
    lastRow = wsData.Cells(wsData.Rows.Count, "S").End(xlUp).Row
    For Each cell In wsData.Range("S2:S" & lastRow) ' Assuming data starts from row 2
        If Not dataRowLabels.exists(cell.Value) Then
            dataRowLabels(cell.Value) = True
        End If
    Next cell
    
    ' Identify extra row labels (exist in PivotTable but not in column S)
    extraRow = 2
    wsExtra.Cells(1, 1).Value = "Extra Row Labels in PivotTable but not in Demand Column S"
    
    For Each key In pivotRowLabels.keys
        If Not dataRowLabels.exists(key) Then
            wsExtra.Cells(extraRow, 1).Value = key
            extraRow = extraRow + 1
        End If
    Next key
    
    MsgBox "Extra row labels identified and placed in 'Extra' sheet.", vbInformation
End Sub
```

```vb
Sub Find_Extra_Row_Labels_As_Table()
    Dim wsPivot As Worksheet
    Dim wsData As Worksheet
    Dim wsExtra As Worksheet
    Dim pt As PivotTable
    Dim pivotRange As Range
    Dim pivotRowLabels As Object
    Dim dataRowLabels As Object
    Dim cell As Range
    Dim lastRow As Long
    Dim extraRow As Long
    Dim key As Variant
    Dim valueCell As Range
    Dim tbl As ListObject
    Dim tblRange As Range
    
    ' Set worksheet references
    Set wsPivot = ThisWorkbook.Sheets("PivotTable") ' Pivot table source
    Set wsData = ThisWorkbook.Sheets("Demand") ' Demand data source
    
    ' Create or clear the "Extra" sheet
    On Error Resume Next
    Set wsExtra = ThisWorkbook.Sheets("Extra")
    If wsExtra Is Nothing Then
        Set wsExtra = ThisWorkbook.Sheets.Add
        wsExtra.Name = "Extra"
    Else
        wsExtra.Cells.Clear ' Clear previous data
    End If
    On Error GoTo 0
    
    ' Set PivotTable reference
    Set pt = wsPivot.PivotTables("MyPivotTable") ' Change to your actual PivotTable name
    
    ' Define dictionaries for storing row labels
    Set pivotRowLabels = CreateObject("Scripting.Dictionary")
    Set dataRowLabels = CreateObject("Scripting.Dictionary")
    
    ' Get PivotTable row labels and their corresponding values
    Set pivotRange = pt.TableRange1.Columns(1) ' First column of PivotTable (Row Labels)
    
    For Each cell In pivotRange.Cells
        If cell.Row > pt.TableRange1.Row Then ' Avoid header row
            ' Store row label and corresponding value (assuming value is in the next column)
            Set valueCell = cell.Offset(0, 1) ' Adjust if values are in a different column
            pivotRowLabels(cell.Value) = valueCell.Value
        End If
    Next cell
    
    ' Get unique values from column S in Demand sheet
    lastRow = wsData.Cells(wsData.Rows.Count, "S").End(xlUp).Row
    For Each cell In wsData.Range("S2:S" & lastRow) ' Assuming data starts from row 2
        If Not dataRowLabels.exists(cell.Value) Then
            dataRowLabels(cell.Value) = True
        End If
    Next cell
    
    ' Identify extra row labels (exist in PivotTable but not in column S)
    extraRow = 2
    wsExtra.Cells(1, 1).Value = "Extra Row Labels in PivotTable"
    wsExtra.Cells(1, 2).Value = "Corresponding Values"
    
    For Each key In pivotRowLabels.keys
        If Not dataRowLabels.exists(key) Then
            wsExtra.Cells(extraRow, 1).Value = key
            wsExtra.Cells(extraRow, 2).Value = pivotRowLabels(key) ' Get corresponding value
            extraRow = extraRow + 1
        End If
    Next key
    
    ' Convert data into a table
    lastRow = wsExtra.Cells(wsExtra.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        Set tblRange = wsExtra.Range("A1:B" & lastRow)
        Set tbl = wsExtra.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        tbl.Name = "ExtraTable"
        tbl.TableStyle = "TableStyleMedium9" ' Apply a table style
    End If
    
    MsgBox "Extra row labels with values placed in 'Extra' sheet as a table.", vbInformation
End Sub
```


```vb
Sub Copy_PivotTable_As_Table()
    Dim wsPivot As Worksheet
    Dim wsNew As Worksheet
    Dim pt As PivotTable
    Dim tblRange As Range
    Dim tbl As ListObject
    Dim lastRow As Long, lastCol As Long
    
    ' Set the worksheet containing the PivotTable
    Set wsPivot = ThisWorkbook.Sheets("PivotTable") ' Change to your actual sheet name
    
    ' Set PivotTable reference
    Set pt = wsPivot.PivotTables("MyPivotTable") ' Change to your actual PivotTable name
    
    ' Define the PivotTable range
    Set tblRange = pt.TableRange1 ' This includes the entire PivotTable
    
    ' Create or clear the "Pivot_Copy" sheet
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets("Pivot_Copy")
    If wsNew Is Nothing Then
        Set wsNew = ThisWorkbook.Sheets.Add
        wsNew.Name = "Pivot_Copy"
    Else
        wsNew.Cells.Clear ' Clear previous data
    End If
    On Error GoTo 0
    
    ' Copy PivotTable range as values to the new sheet
    tblRange.Copy
    wsNew.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsNew.Range("A1").PasteSpecial Paste:=xlPasteFormats ' Keep formatting
    Application.CutCopyMode = False ' Clear clipboard
    
    ' Determine the last row and last column
    lastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    lastCol = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column
    
    ' Convert the copied data into a table
    Set tblRange = wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(lastRow, lastCol))
    Set tbl = wsNew.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = "CopiedPivotTable"
    tbl.TableStyle = "TableStyleMedium9" ' Apply a table style
    
    MsgBox "PivotTable copied as a table in 'Pivot_Copy' sheet.", vbInformation
End Sub
```

```vb
Sub Remove_Extra_Staff_From_Pivot_Copy()
    Dim wsPivotCopy As Worksheet
    Dim wsData As Worksheet
    Dim wsExtra As Worksheet
    Dim tbl As ListObject
    Dim pivotRowLabels As Object
    Dim dataRowLabels As Object
    Dim cell As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim extraRow As Long
    Dim key As Variant
    Dim rowIndex As Long
    Dim deleteRows As Range
    Dim firstDelete As Boolean
    
    ' Set worksheet references
    Set wsPivotCopy = ThisWorkbook.Sheets("Pivot_Copy") ' Sheet with copied PivotTable
    Set wsData = ThisWorkbook.Sheets("Demand") ' Demand data source
    
    ' Get the table in Pivot_Copy
    On Error Resume Next
    Set tbl = wsPivotCopy.ListObjects("CopiedPivotTable") ' Change to actual table name if needed
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Table 'CopiedPivotTable' not found in Pivot_Copy!", vbExclamation
        Exit Sub
    End If
    
    ' Create or clear the "Extra" sheet
    On Error Resume Next
    Set wsExtra = ThisWorkbook.Sheets("Extra")
    If wsExtra Is Nothing Then
        Set wsExtra = ThisWorkbook.Sheets.Add
        wsExtra.Name = "Extra"
    Else
        wsExtra.Cells.Clear ' Clear previous data
    End If
    On Error GoTo 0
    
    ' Define dictionaries for storing row labels
    Set pivotRowLabels = CreateObject("Scripting.Dictionary")
    Set dataRowLabels = CreateObject("Scripting.Dictionary")
    
    ' Get last column in Pivot_Copy
    lastCol = tbl.Range.Columns.Count
    
    ' Get row labels from Pivot_Copy (assuming they are in the first column of the table)
    For Each cell In tbl.ListColumns(1).DataBodyRange
        If Not pivotRowLabels.exists(cell.Value) Then
            pivotRowLabels(cell.Value) = cell.Row ' Store row number
        End If
    Next cell
    
    ' Get unique values from column S in Demand sheet
    lastRow = wsData.Cells(wsData.Rows.Count, "S").End(xlUp).Row
    For Each cell In wsData.Range("S2:S" & lastRow) ' Assuming data starts from row 2
        If Not dataRowLabels.exists(cell.Value) Then
            dataRowLabels(cell.Value) = True
        End If
    Next cell
    
    ' Identify extra row labels and delete from Pivot_Copy
    extraRow = 2
    wsExtra.Range("A1").Value = "Removed Row Labels"
    
    firstDelete = True ' Track first deletion for Union function
    
    For Each key In pivotRowLabels.keys
        rowIndex = pivotRowLabels(key)
        
        ' If the label is NOT in column S, delete from Pivot_Copy and store in Extra
        If Not dataRowLabels.exists(key) Then
            ' Copy the row to Extra sheet
            wsExtra.Range(wsExtra.Cells(extraRow, 1), wsExtra.Cells(extraRow, lastCol)).Value = _
                wsPivotCopy.Range(wsPivotCopy.Cells(rowIndex, 1), wsPivotCopy.Cells(rowIndex, lastCol)).Value
            
            ' Mark the row for deletion
            If firstDelete Then
                Set deleteRows = wsPivotCopy.Rows(rowIndex)
                firstDelete = False
            Else
                Set deleteRows = Union(deleteRows, wsPivotCopy.Rows(rowIndex))
            End If
            
            extraRow = extraRow + 1
        End If
    Next key
    
    ' Delete the marked rows at once
    If Not deleteRows Is Nothing Then deleteRows.Delete Shift:=xlUp
    
    ' Convert Extra data into a table
    lastRow = wsExtra.Cells(wsExtra.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        Set tbl = wsExtra.ListObjects.Add(xlSrcRange, wsExtra.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = "ExtraTable"
        tbl.TableStyle = "TableStyleMedium9"
    End If
    
    MsgBox "Extra staff removed from Pivot_Copy and stored in 'Extra' sheet.", vbInformation
End Sub
```