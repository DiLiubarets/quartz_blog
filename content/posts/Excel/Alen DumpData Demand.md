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
Sub SyncSlicerWithFilter()
    Dim wsPivot As Worksheet
    Dim wsData As Worksheet
    Dim slicerCache As SlicerCache
    Dim slicerItem As SlicerItem
    Dim selectedItems As Collection
    Dim filterCriteria As String
    Dim i As Long
    Dim rng As Range
    Dim lastRow As Long
    
    Application.ScreenUpdating = False
    
    Set wsPivot = ThisWorkbook.Sheets("Sheet1")
    Set wsData = ThisWorkbook.Sheets("Sheet2")
    Set slicerCache = ThisWorkbook.SlicerCaches("Slicer_Project")
    
    'Get last row in the data sheet
    lastRow = wsData.Cells(wsData.Rows.Count, "D").End(xlUp).Row
    
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
    
    'Apply filter
    wsData.AutoFilterMode = False
    wsData.Range("D1:D" & lastRow).AutoFilter Field:=1, Criteria1:=Split(filterCriteria, ","), Operator:=xlFilterValues
    
    'Delete filtered rows
    On Error Resume Next
    Set rng = wsData.Range("D2:D" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        rng.EntireRow.Delete
    End If
    
    'Turn off filter
    wsData.AutoFilterMode = False
    
    Application.ScreenUpdating = True
    
    MsgBox "Rows deleted successfully based on slicer selection!", vbInformation
End Sub
```