
To automate the creation of my Pivot Table to simplify the reporting process.

```vb
Sub CreatePivotTable()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range

    ' Data sheet and range
    Set wsData = ThisWorkbook.Worksheets("Sheet1")  
    Set pivotRange = wsData.Range("A1").CurrentRegion 

    ' New sheet 
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("PivotTable")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "PivotTable"
    End If
    On Error GoTo 0

    
    Set pivotDestination = wsPivot.Range("A3")

    
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    ' Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTable")

    ' Fields 
    With pivotTable
        .PivotFields("Category").Orientation = xlRowField 
        .PivotFields("Region").Orientation = xlColumnField 
        .PivotFields("Sales").Orientation = xlDataField 
        .PivotFields("Sales").Function = xlSum 
    End With

    MsgBox "Pivot Table created successfully!", vbInformation
End Sub
```

Update pivot table with new Data 
```vb
Sub RefreshAllPivotTables()
    Dim ws As Worksheet
    Dim pt As PivotTable

    ' refresh Pivot Tables
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    MsgBox "All Pivot Tables have been refreshed!", vbInformation
End Sub
```

Two pivot tables``
```vb
Sub CreatePivotTables()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable1 As PivotTable
    Dim pivotTable2 As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination1 As Range
    Dim pivotDestination2 As Range

    
    Set wsData = ThisWorkbook.Worksheets("Sheet1") 
    Set pivotRange = wsData.Range("A1").CurrentRegion 
    
    Debug.Print "Pivot Range: " & pivotRange.Address

    
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("PivotTable")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "PivotTable"
    End If
    wsPivot.Cells.Clear 
    On Error GoTo 0

    
    Set pivotDestination1 = wsPivot.Range("A3") 
    Set pivotDestination2 = wsPivot.Range("G3") 

   
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    
    Debug.Print "Pivot Cache Source: " & pivotCache.SourceData

    
    Set pivotTable1 = pivotCache.CreatePivotTable(TableDestination:=pivotDestination1, TableName:="MyPivotTable1")
    With pivotTable1
        .PivotFields("Category").Orientation = xlRowField 
        .PivotFields("Region").Orientation = xlColumnField 

       
        On Error Resume Next
        .PivotFields("Sales").Orientation = xlDataField 
        .PivotFields("Sales").Function = xlSum 
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Sales': " & Err.Description
            .PivotFields("Sales").Function = xlCount
        End If
        On Error GoTo 0
    End With

    
    Set pivotTable2 = pivotCache.CreatePivotTable(TableDestination:=pivotDestination2, TableName:="MyPivotTable2")
    With pivotTable2
        .PivotFields("Region").Orientation = xlRowField 
        .PivotFields("Category").Orientation = xlColumnField 

       
        On Error Resume Next
        .PivotFields("Profit").Orientation = xlDataField 
        .PivotFields("Profit").Function = xlSum 
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Profit': " & Err.Description
            .PivotFields("Profit").Function = xlCount 
        End If
        On Error GoTo 0
    End With
    

    MsgBox "Two Pivot Tables created successfully!", vbInformation
End Sub
```


Pivot tables with slicer

```vb
Sub CreatePivotTablesWithSlicer()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable1 As PivotTable
    Dim pivotTable2 As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination1 As Range
    Dim pivotDestination2 As Range
    Dim pSlicers As Slicers
    Dim sSlicer As Slicer
    Dim pSlicersCaches As SlicerCaches
    Dim sSlicerCache As SlicerCache
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set wsData = ThisWorkbook.Worksheets("Sheet1")
    Set pivotRange = wsData.Range("A1").CurrentRegion
    Debug.Print "Pivot Range: " & pivotRange.Address

    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("PivotTable")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "PivotTable"
    End If
    wsPivot.Cells.Clear
    On Error GoTo 0

    Set pivotDestination1 = wsPivot.Range("A5")
    Set pivotDestination2 = wsPivot.Range("G5")

    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    Set pivotTable1 = pivotCache.CreatePivotTable(TableDestination:=pivotDestination1, TableName:="MyPivotTable1")
    With pivotTable1
        .PivotFields("Assignee").Orientation = xlRowField
        On Error Resume Next
        .PivotFields("Story Points").Orientation = xlDataField
        .PivotFields("Story Points").Function = xlSum
        On Error GoTo 0
    End With

    Set pivotTable2 = pivotCache.CreatePivotTable(TableDestination:=pivotDestination2, TableName:="MyPivotTable2")
    With pivotTable2
        .PivotFields("Assignee").Orientation = xlRowField
        .PivotFields("Project").Orientation = xlRowField
        .PivotFields("Summary").Orientation = xlRowField
        On Error Resume Next
        .PivotFields("Story Points").Orientation = xlDataField
        .PivotFields("Story Points").Function = xlSum
        On Error GoTo 0
    End With

    Set pSlicersCaches = wb.SlicerCaches
    Set sSlicerCache = pSlicersCaches.Add2(pivotTable1, "Status", "Status")
    Set sSlicer = sSlicerCache.Slicers.Add(SlicerDestination:=wsPivot.Name, _
                                           Name:="StatusSlicer", _
                                           Caption:="Status", _
                                           Top:=6, _
                                           Left:=6, _
                                           Width:=254, _
                                           Height:=109)

    sSlicer.NumberOfColumns = 3
    sSlicer.RowHeight = 28.8
    
    sSlicerCache.PivotTables.AddPivotTable pivotTable2

    MsgBox "Two Pivot Tables and a Slicer created successfully!", vbInformation
End Sub
```

Dynamic pivot table positioning

```vb
Sub CreatePivotTablesWithSlicer()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable1 As PivotTable
    Dim pivotTable2 As PivotTable
    Dim pivotTable3 As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination1 As Range
    Dim pivotDestination2 As Range
    Dim pivotDestination3 As Range
    Dim pSlicers As Slicers
    Dim sSlicer As Slicer
    Dim pSlicersCaches As SlicerCaches
    Dim sSlicerCache As SlicerCache
    Dim wb As Workbook
    Dim lastRow As Long

    Set wb = ThisWorkbook
    Set wsData = ThisWorkbook.Worksheets("Sheet1")
    Set pivotRange = wsData.Range("A1").CurrentRegion
    Debug.Print "Pivot Range: " & pivotRange.Address

    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("PivotTable")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "PivotTable"
    End If
    wsPivot.Cells.Clear
    On Error GoTo 0
   
    Set pivotDestination3 = wsPivot.Range("A7")
    Set pivotDestination2 = wsPivot.Range("E2")

    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    Set pivotTable3 = pivotCache.CreatePivotTable(TableDestination:=pivotDestination3, TableName:="MyPivotTable3")
    With pivotTable3
        .PivotFields("Project").Orientation = xlRowField
        On Error Resume Next
        .PivotFields("Story Points").Orientation = xlDataField
        .PivotFields("Story Points").Function = xlSum
        On Error GoTo 0
    End With

    lastRow = wsPivot.Cells(wsPivot.Rows.Count, "A").End(xlUp).Row + 2
    Set pivotDestination1 = wsPivot.Range("A" & lastRow)

    Set pivotTable1 = pivotCache.CreatePivotTable(TableDestination:=pivotDestination1, TableName:="MyPivotTable1")
    With pivotTable1
        .PivotFields("Assignee").Orientation = xlRowField
        On Error Resume Next
        .PivotFields("Story Points").Orientation = xlDataField
        .PivotFields("Story Points").Function = xlSum
        On Error GoTo 0
    End With

    Set pivotTable2 = pivotCache.CreatePivotTable(TableDestination:=pivotDestination2, TableName:="MyPivotTable2")
    With pivotTable2
        .PivotFields("Assignee").Orientation = xlRowField
        .PivotFields("Project").Orientation = xlRowField
        .PivotFields("Summary").Orientation = xlRowField
        On Error Resume Next
        .PivotFields("Story Points").Orientation = xlDataField
        .PivotFields("Story Points").Function = xlSum
        On Error GoTo 0
    End With

    Set pSlicersCaches = wb.SlicerCaches
    Set sSlicerCache = pSlicersCaches.Add2(pivotTable1, "Status", "Status")
    Set sSlicer = sSlicerCache.Slicers.Add(SlicerDestination:=wsPivot.Name, _
                                           Name:="StatusSlicer", _
                                           Caption:="Status", _
                                           Top:=6, _
                                           Left:=6, _
                                           Width:=254, _
                                           Height:=50)

    sSlicer.NumberOfColumns = 3
    sSlicer.RowHeight = 21.7
   
    sSlicerCache.PivotTables.AddPivotTable pivotTable2

End Sub
```
Tabular form and repeat all item Labels
```vb
Sub CreatePivotTable()

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range

    Set wsData = ThisWorkbook.Worksheets("Vacashing Data")
    Set pivotRange = wsData.Range("A1").CurrentRegion

    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("PivotTable")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "PivotTable"
    End If
    On Error GoTo 0

    wsPivot.Cells.Clear
    Set pivotDestination = wsPivot.Range("A3")
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTable")

    With pivotTable
        .PivotFields("Primary Manager").Orientation = xlRowField
        .PivotFields("Work center").Orientation = xlRowField
        .PivotFields("EID").Orientation = xlRowField
        .PivotFields("Name of Employee").Orientation = xlRowField
        .PivotFields("Hours").Orientation = xlDataField
        .PivotFields("Hours").Function = xlSum
        On Error Resume Next
        .RowAxisLayout xlTabularRow
        On Error GoTo 0
        .RepeatAllLabels xlRepeatLabels

        Dim pf As PivotField
        For Each pf In .RowFields
            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            pf.LayoutBlankLine = False
        Next pf

        .ColumnGrand = False
        .RowGrand = False
    End With

    MsgBox "Pivot Table created successfully in Tabular Form!", vbInformation

End Sub
```

```vb 
Sub CreatePivotTableWithSlicers()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim pSlicersCaches As SlicerCaches
    Dim sSlicerCache1 As SlicerCache
    Dim sSlicerCache2 As SlicerCache
    Dim sSlicer1 As Slicer
    Dim sSlicer2 As Slicer
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set wsData = ThisWorkbook.Worksheets("Vacashing Data")
    Set pivotRange = wsData.Range("A1").CurrentRegion

    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("PivotTable")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "PivotTable"
    End If
    wsPivot.Cells.Clear
    On Error GoTo 0

    Set pivotDestination = wsPivot.Range("A3")
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTable")

    With pivotTable
        .PivotFields("Month").Orientation = xlColumnField
        .PivotFields("Primary Manager").Orientation = xlRowField
        .PivotFields("Work center").Orientation = xlRowField
        .PivotFields("EID").Orientation = xlRowField
        .PivotFields("Name of Employee").Orientation = xlRowField
        .PivotFields("Hours").Orientation = xlDataField
        .PivotFields("Hours").Function = xlSum
        .ColumnGrand = False
        .RowGrand = False
    End With

    Set pSlicersCaches = wb.SlicerCaches
    Set sSlicerCache1 = pSlicersCaches.Add2(pivotTable, "Time off Description", "Time off Description")
    Set sSlicer1 = sSlicerCache1.Slicers.Add(SlicerDestination:=wsPivot.Name, _
                                           Name:="TimeOffSlicer", _
                                           Caption:="Time off Description", _
                                           Top:=6, _
                                           Left:=6, _
                                           Width:=254, _
                                           Height:=109)

    Set sSlicerCache2 = pSlicersCaches.Add2(pivotTable, "Indirect/Direct", "Indirect/Direct")
    Set sSlicer2 = sSlicerCache2.Slicers.Add(SlicerDestination:=wsPivot.Name, _
                                           Name:="IndirectSlicer", _
                                           Caption:="Indirect/Direct", _
                                           Top:=120, _
                                           Left:=6, _
                                           Width:=254, _
                                           Height:=109)

    sSlicer1.NumberOfColumns = 3
    sSlicer1.RowHeight = 28.8
    
    sSlicer2.NumberOfColumns = 3
    sSlicer2.RowHeight = 28.8

    MsgBox "Pivot Table with Slicers created successfully!", vbInformation
End Sub
```

new 
```vb 
Sub CreatePivotTableWithSlicers()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim pSlicersCaches As SlicerCaches
    Dim sSlicerCache1 As SlicerCache
    Dim sSlicerCache2 As SlicerCache
    Dim sSlicerCache3 As SlicerCache
    Dim sSlicer1 As Slicer
    Dim sSlicer2 As Slicer
    Dim sSlicer3 As Slicer
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set wsData = ThisWorkbook.Worksheets("Vacashing Data")
    Set pivotRange = wsData.Range("A1").CurrentRegion

    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("PivotTable")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "PivotTable"
    End If
    wsPivot.Cells.Clear
    On Error GoTo 0

    Set pivotDestination = wsPivot.Range("A3")
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTable")

    With pivotTable
        .PivotFields("Month").Orientation = xlColumnField
        .PivotFields("Primary Manager").Orientation = xlRowField
        .PivotFields("Work center").Orientation = xlRowField
        .PivotFields("EID").Orientation = xlRowField
        .PivotFields("Name of Employee").Orientation = xlRowField
        .PivotFields("Hours").Orientation = xlDataField
        .PivotFields("Hours").Function = xlSum
        .ColumnGrand = False
        .RowGrand = False
    End With

    Set pSlicersCaches = wb.SlicerCaches
    
    ' First Slicer - Time off Description
    Set sSlicerCache1 = pSlicersCaches.Add2(pivotTable, "Time off Description", "Time off Description")
    Set sSlicer1 = sSlicerCache1.Slicers.Add(SlicerDestination:=wsPivot.Name, _
                                           Name:="TimeOffSlicer", _
                                           Caption:="Time off Description", _
                                           Top:=6, _
                                           Left:=600, _
                                           Width:=254, _
                                           Height:=109)

    ' Second Slicer - Indirect/Direct
    Set sSlicerCache2 = pSlicersCaches.Add2(pivotTable, "Indirect/Direct", "Indirect/Direct")
    Set sSlicer2 = sSlicerCache2.Slicers.Add(SlicerDestination:=wsPivot.Name, _
                                           Name:="IndirectSlicer", _
                                           Caption:="Indirect/Direct", _
                                           Top:=120, _
                                           Left:=600, _
                                           Width:=254, _
                                           Height:=109)
                                           
    ' Third Slicer - Primary Manager
    Set sSlicerCache3 = pSlicersCaches.Add2(pivotTable, "Primary Manager", "Primary Manager")
    Set sSlicer3 = sSlicerCache3.Slicers.Add(SlicerDestination:=wsPivot.Name, _
                                           Name:="ManagerSlicer", _
                                           Caption:="Primary Manager", _
                                           Top:=234, _
                                           Left:=600, _
                                           Width:=254, _
                                           Height:=109)

    ' Format Slicers
    sSlicer1.NumberOfColumns = 3
    sSlicer1.RowHeight = 28.8
    
    sSlicer2.NumberOfColumns = 3
    sSlicer2.RowHeight = 28.8
    
    sSlicer3.NumberOfColumns = 3
    sSlicer3.RowHeight = 28.8

    MsgBox "Pivot Table with Slicers created successfully!", vbInformation
End Sub
```

2
```vb 
Sub CreatePivotTableWithSlicers()
    On Error Resume Next
    
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim pSlicersCaches As SlicerCaches
    Dim sSlicerCache1 As SlicerCache
    Dim sSlicerCache2 As SlicerCache
    Dim sSlicerCache3 As SlicerCache
    Dim sSlicer1 As Slicer
    Dim sSlicer2 As Slicer
    Dim sSlicer3 As Slicer
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    Set wsData = ThisWorkbook.Worksheets("Vacashing Data")
    Set pivotRange = wsData.Range("A1").CurrentRegion
    
    'Create or clear pivot sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PivotTable").Delete
    On Error GoTo 0
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "PivotTable"
    Application.DisplayAlerts = True
    
    'Create pivot table
    Set pivotDestination = wsPivot.Range("A3")
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTable")
    
    With pivotTable
        .PivotFields("Month").Orientation = xlColumnField
        .PivotFields("Primary Manager").Orientation = xlRowField
        .PivotFields("Work center").Orientation = xlRowField
        .PivotFields("EID").Orientation = xlRowField
        .PivotFields("Name of Employee").Orientation = xlRowField
        .PivotFields("Hours").Orientation = xlDataField
        .PivotFields("Hours").Function = xlSum
    End With
    
    'Create Slicers
    On Error Resume Next
    
    'First Slicer
    If Err.Number = 0 Then
        Set sSlicerCache1 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Time off Description")
        If Err.Number = 0 Then
            Set sSlicer1 = sSlicerCache1.Slicers.Add(wsPivot.Name, , "TimeOffSlicer", "Time off Description", 6, 600)
            With sSlicer1
                .Width = 254
                .Height = 109
                .NumberOfColumns = 3
                .RowHeight = 28.8
            End With
        Else
            MsgBox "Error creating first slicer: " & Err.Description
        End If
    End If
    
    'Second Slicer
    Err.Clear
    If Err.Number = 0 Then
        Set sSlicerCache2 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Indirect/Direct")
        If Err.Number = 0 Then
            Set sSlicer2 = sSlicerCache2.Slicers.Add(wsPivot.Name, , "IndirectSlicer", "Indirect/Direct", 120, 600)
            With sSlicer2
                .Width = 254
                .Height = 109
                .NumberOfColumns = 3
                .RowHeight = 28.8
            End With
        Else
            MsgBox "Error creating second slicer: " & Err.Description
        End If
    End If
    
    'Third Slicer
    Err.Clear
    If Err.Number = 0 Then
        Set sSlicerCache3 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Primary Manager")
        If Err.Number = 0 Then
            Set sSlicer3 = sSlicerCache3.Slicers.Add(wsPivot.Name, , "ManagerSlicer", "Primary Manager", 234, 600)
            With sSlicer3
                .Width = 254
                .Height = 109
                .NumberOfColumns = 3
                .RowHeight = 28.8
            End With
        Else
            MsgBox "Error creating third slicer: " & Err.Description
        End If
    End If
    
    On Error GoTo 0
    
    MsgBox "Pivot Table with Slicers created successfully!", vbInformation
End Sub
```

chart 
```vb 
Sub CreatePivotTable()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim pvtChart As Shape
    
    Set wsData = ThisWorkbook.Worksheets("Vacation Data")
    Set pivotRange = wsData.Range("A1").CurrentRegion

    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("VAC-HOL-BANKED Chart")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "VAC-HOL-BANKED Chart"
    End If
    On Error GoTo 0
    
    Set pivotDestination = wsPivot.Range("F16")

    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTable")

    With pivotTable
        .PivotFields("Week#").Orientation = xlRowField
        .PivotFields("Name of Employee").Orientation = xlColumnField
        .PivotFields("Month").Orientation = xlPageField    'Changed to Page Field (Filter)
        .PivotFields("Hours").Orientation = xlDataField
        .PivotFields("Hours").Function = xlSum
    End With

    'Add Chart
    wsPivot.Activate
    Set pvtChart = wsPivot.Shapes.AddChart2
    
    With pvtChart.Chart
        .SetSourceData Source:=pivotTable.TableRange1
        .ChartType = xlColumnClustered
        
        'Customize chart
        With .Parent
            .Left = pivotTable.TableRange1.Left
            .Top = pivotTable.TableRange1.Top - 200
            .Width = 800    'Increased width to accommodate employee names
            .Height = 400
        End With
        
        'Add titles
        .HasTitle = True
        .ChartTitle.Text = "Hours by Employee and Week"
        
        'Customize axes
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Hours"
        End With
        
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Week#"
        End With
        
        'Format legend
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        'Rotate category labels if needed
        .Axes(xlCategory).TickLabels.Orientation = 0
    End With

    'Auto-fit the pivot table columns
    pivotTable.TableRange1.Columns.AutoFit

    MsgBox "Pivot Table and Chart created successfully!", vbInformation
End Sub
```