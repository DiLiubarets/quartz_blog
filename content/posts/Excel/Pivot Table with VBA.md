
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

Tabular form/Slicer with VBA 
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

Pivot Table and 4Slicer
```vb 
Sub CreatePivotTableWithSlicers_VACHOL_BANKED_ChartTableSlicers()

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
    Dim wb As Workbook
    Dim targetCell As Range

	Set wb = ThisWorkbook

    Set wsData = ThisWorkbook.Worksheets("Vacation Data")
    Set pivotRange = wsData.Range("A1").CurrentRegion
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("VAC-HOL-BANKED Chart").Delete

    On Error GoTo 0

    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "VAC-HOL-BANKED Chart"
    Application.DisplayAlerts = True

    'Create pivot table
    Set pivotDestination = wsPivot.Range("I22")
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTableTable")
    With pivotTable
        .PivotFields("Name of Employee").Orientation = xlColumnField
        .PivotFields("Week#").Orientation = xlRowField
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Hours").Orientation = xlDataField

        'Sum check
        On Error Resume Next
        .PivotFields("Hours").Function = xlSum
        If Err.Number <> 0 Then
            Debug.Print
                .PivotFields("Hours").Function = xlCount
            Err.Clear
        End If

        On Error GoTo 0
        'xlTabularRow
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .ColumnGrand = False
        .SubtotalHiddenPageItems = False

         Dim pf As PivotField
        'Subtotals
        For Each pf In .RowFields
            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            pf.LayoutBlankLine = False
        Next pf

    End With

    'Create Slicers
    If Err.Number = 0 Then

        Set sSlicerCache1 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Cost Center")
        If Err.Number = 0 Then
            Set sSlicer1 = sSlicerCache1.Slicers.Add(wsPivot.Name, , "CostCenterChart", "Cost Center", 80, 25)

            With sSlicer1
                .Width = 900
                .Height = 70
                .NumberOfColumns = 8
                .RowHeight = 20
            End With

        Else
            MsgBox "Error creating first slicer: " & Err.Description
        End If

    End If

    Err.Clear

    If Err.Number = 0 Then
        Set sSlicerCache2 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Indirect/Direct")
        If Err.Number = 0 Then
            Set sSlicer2 = sSlicerCache2.Slicers.Add(wsPivot.Name, , "IndirectSlicerChart", "Indirect/Direct", 20, 25)
            With sSlicer2
                .Width = 130
                .Height = 50
                .NumberOfColumns = 2
                .RowHeight = 15
            End With

        Else

            MsgBox "Error creating second slicer: " & Err.Description
        End If

    End If

    Err.Clear

    If Err.Number = 0 Then

        Set sSlicerCache3 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Primary Manager")
        If Err.Number = 0 Then
            Set sSlicer3 = sSlicerCache3.Slicers.Add(wsPivot.Name, , "ManagerSlicerChart", "Primary Manager", 150, 25)

            With sSlicer3
                .Width = 900
                .Height = 50
                .NumberOfColumns = 8
                .RowHeight = 20
            End With

        Else
            MsgBox "Error creating third slicer: " & Err.Description
        End If
    End If

        Err.Clear
    If Err.Number = 0 Then
        Set sSlicerCache4 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Month")
        If Err.Number = 0 Then
            Set sSlicer4 = sSlicerCache4.Slicers.Add(wsPivot.Name, , "MonthChart", "Month", 200, 25)
            With sSlicer4
                .Width = 900
                .Height = 50
                .NumberOfColumns = 12
                .RowHeight = 20
            End With
        Else
            MsgBox "Error creating second slicer: " & Err.Description
        End If
    End If
    On Error GoTo 0
    MsgBox "VAC-HOL-BANKED Table created!", vbInformation
End Sub
```
Chart added 
```vb Sub CreatePivotTableWithSlicers_VACHOL_BANKED_Chart()

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

    Dim pvtChart As Shape

    Dim wb As Workbook

    Dim targetCell As Range

    Set wb = ThisWorkbook

    Set wsData = ThisWorkbook.Worksheets("Vacation Data")

    Set pivotRange = wsData.Range("A1").CurrentRegion

    Application.DisplayAlerts = False

    On Error Resume Next

    ThisWorkbook.Sheets("VAC-HOL-BANKED Chart").Delete

    On Error GoTo 0

    Set wsPivot = ThisWorkbook.Sheets.Add

    wsPivot.Name = "VAC-HOL-BANKED Chart"

    Application.DisplayAlerts = True

    'Create pivot table

    Set pivotDestination = wsPivot.Range("W22")

    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTableTable")

    With pivotTable

        .PivotFields("Name of Employee").Orientation = xlColumnField

        .PivotFields("Week#").Orientation = xlRowField

        .PivotFields("Month").Orientation = xlRowField

        .PivotFields("Hours").Orientation = xlDataField

        'Sum check

        On Error Resume Next

        .PivotFields("Hours").Function = xlSum

        If Err.Number <> 0 Then

            Debug.Print

                .PivotFields("Hours").Function = xlCount

            Err.Clear

        End If

        On Error GoTo 0

        'xlTabularRow

        .RowAxisLayout xlTabularRow

        .RowGrand = False

        .ColumnGrand = False

        .SubtotalHiddenPageItems = False

         Dim pf As PivotField

         'Subtotals

        For Each pf In .RowFields

            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

            pf.LayoutBlankLine = False

        Next pf

    End With

    'Create Slicers

    If Err.Number = 0 Then

        Set sSlicerCache1 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Cost Center")

        If Err.Number = 0 Then

            Set sSlicer1 = sSlicerCache1.Slicers.Add(wsPivot.Name, , "CostCenterChart", "Cost Center", 80, 25)

            With sSlicer1

                .Width = 900

                .Height = 70

                .NumberOfColumns = 8

                .RowHeight = 20

            End With

        Else

            MsgBox "Error creating first slicer: " & Err.Description

        End If

    End If

    Err.Clear

    If Err.Number = 0 Then

        Set sSlicerCache2 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Indirect/Direct")

        If Err.Number = 0 Then

            Set sSlicer2 = sSlicerCache2.Slicers.Add(wsPivot.Name, , "IndirectSlicerChart", "Indirect/Direct", 20, 25)

            With sSlicer2

                .Width = 130

                .Height = 50

                .NumberOfColumns = 2

                .RowHeight = 15

            End With

        Else

            MsgBox "Error creating second slicer: " & Err.Description

        End If

    End If

    Err.Clear

    If Err.Number = 0 Then

        Set sSlicerCache3 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Primary Manager")

        If Err.Number = 0 Then

            Set sSlicer3 = sSlicerCache3.Slicers.Add(wsPivot.Name, , "ManagerSlicerChart", "Primary Manager", 150, 25)

            With sSlicer3

                .Width = 900

                .Height = 50

                .NumberOfColumns = 8

                .RowHeight = 20

            End With

        Else

            MsgBox "Error creating third slicer: " & Err.Description

        End If

    End If

        Err.Clear

    If Err.Number = 0 Then

        Set sSlicerCache4 = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Month")

        If Err.Number = 0 Then

            Set sSlicer4 = sSlicerCache4.Slicers.Add(wsPivot.Name, , "MonthChart", "Month", 200, 25)

            With sSlicer4

                .Width = 900

                .Height = 50

                .NumberOfColumns = 12

                .RowHeight = 20

            End With

        Else

            MsgBox "Error creating second slicer: " & Err.Description

        End If

    End If

    On Error GoTo 0

     'Add Chart

    wsPivot.Activate

    Set pvtChart = wsPivot.Shapes.AddChart2

    With pvtChart.Chart

        .SetSourceData Source:=pivotTable.TableRange1

        .ChartType = xlColumnClustered

        With .Parent

            .Left = pivotTable.TableRange1.Left

            .Top = pivotTable.TableRange1.Top

            .Width = 1000    'Increased width to accommodate employee names

            .Height = 800

        End With

        '

        .HasTitle = True

        .ChartTitle.Text = "SATCOM Vacation by Primary Manager"

        'Customize axes

        With .Axes(xlValue, xlPrimary)

            .HasTitle = False

        End With

        With .Axes(xlCategory, xlPrimary)

            .HasTitle = True

            .AxisTitle.Text = "Week#"

        End With

        .HasLegend = True

        .Legend.Position = xlLegendPositionBottom

        .Axes(xlCategory).TickLabels.Orientation = 0

    End With

    'Auto-fit the pivot table columns

    pivotTable.TableRange1.Columns.AutoFit

    MsgBox "VAC-HOL-BANKED Table created!", vbInformation

End Sub
```

from two data sheets
```vb
Sub CreatePivotTableFromTwoSheets()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim wsPivot As Worksheet
    Dim LastRow1 As Long, LastRow2 As Long
    Dim PvtCache As PivotCache
    Dim PvtTable As PivotTable
    Dim TempSheet As Worksheet
    
    'Define worksheets
    Set ws1 = ThisWorkbook.Sheets("Sheet1") 'Change to your first sheet name
    Set ws2 = ThisWorkbook.Sheets("Sheet2") 'Change to your second sheet name
    
    'Create a new sheet for the pivot table
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("PivotTable").Delete
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "PivotTable"
    Application.DisplayAlerts = True
    
    'Create temporary sheet to combine data
    Set TempSheet = ThisWorkbook.Sheets.Add
    TempSheet.Name = "TempData"
    
    'Find last rows of both sheets
    LastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    LastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    'Copy headers from first sheet
    ws1.Range("A1:D1").Copy TempSheet.Range("A1") 'Adjust range as needed
    
    'Copy data from first sheet
    ws1.Range("A2:D" & LastRow1).Copy TempSheet.Range("A2") 'Adjust range as needed
    
    'Copy data from second sheet (excluding headers)
    ws2.Range("A2:D" & LastRow2).Copy _
        TempSheet.Range("A" & LastRow1 + 1) 'Adjust range as needed
    
    'Create Pivot Cache
    Set PvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=TempSheet.UsedRange)
    
    'Create Pivot Table
    Set PvtTable = PvtCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="CombinedPivotTable")
    
    'Add fields to pivot table (adjust field names as needed)
    With PvtTable
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Category").Position = 1
        
        .PivotFields("Sales").Orientation = xlDataField
        .PivotFields("Sales").Position = 1
        .PivotFields("Sales").Function = xlSum
    End With
    
    'Delete temporary sheet
    Application.DisplayAlerts = False
    TempSheet.Delete
    Application.DisplayAlerts = True
    
    'Format pivot table
    wsPivot.Cells.EntireColumn.AutoFit
    
    MsgBox "Pivot Table created successfully!", vbInformation
    
End Sub
```