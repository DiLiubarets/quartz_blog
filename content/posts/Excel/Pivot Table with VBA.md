
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
Pivot Table with calculated field
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
        
        ' Add calculated field
        On Error Resume Next
        .CalculatedFields("Sales Percentage").Delete  ' Delete if exists
        On Error GoTo 0
        
        ' Add new calculated field
        With .CalculatedFields.Add( _
            Name:="Sales Percentage", _
            Formula:="=Sales/Sum(Sales)", _
            UseStandardFormula:=True)
        End With
        
        ' Add calculated field to the pivot table
        .PivotFields("Sales Percentage").Orientation = xlDataField
        
        ' Format the calculated field as percentage
        With .PivotFields("Sales Percentage")
            .NumberFormat = "0.00%"
            .Function = xlSum
        End With
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
        .PivotFields("Month").Orientation = xlPageField    
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
            .Width = 800    
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
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
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

    wsPivot.Activate
    Set pvtChart = wsPivot.Shapes.AddChart2
    With pvtChart.Chart
        .SetSourceData Source:=pivotTable.TableRange1
        .ChartType = xlColumnClustered
        With .Parent
            .Left = pivotTable.TableRange1.Left
            .Top = pivotTable.TableRange1.Top
            .Width = 1000    
            .Height = 800
        End With
 
        .HasTitle = True
        .ChartTitle.Text = "SATCOM Vacation by Primary Manager"

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
  
    Set ws1 = ThisWorkbook.Sheets("Sheet1") 
    Set ws2 = ThisWorkbook.Sheets("Sheet2") 
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("PivotTable").Delete
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "PivotTable"
    Application.DisplayAlerts = True
  
    Set TempSheet = ThisWorkbook.Sheets.Add
    TempSheet.Name = "TempData"
    
       LastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    LastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
 
    ws1.Range("A1:D1").Copy TempSheet.Range("A1") 
    
     ws1.Range("A2:D" & LastRow1).Copy TempSheet.Range("A2") 
    
  
    ws2.Range("A2:D" & LastRow2).Copy _
        TempSheet.Range("A" & LastRow1 + 1) '    
   
    Set PvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=TempSheet.UsedRange)
    
    Set PvtTable = PvtCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="CombinedPivotTable")
    
    '    With PvtTable
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
Pivot with Slicer and Conditional formatting 
```vb
Sub CreatePivotTable_Day_Hours_Available1()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As pivotCache
    Dim pivotTable As pivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim pSlicersCaches As SlicerCaches
    Dim sSlicerCache As SlicerCache
    Dim sSlicer As Slicer
 
    Set wsData = ThisWorkbook.Worksheets("Quota RPT")
    Set pivotRange = wsData.Range("A1").CurrentRegion

    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("Days-Hours1")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "Days-Hours1"
    End If
    On Error GoTo 0
    Set pivotDestination = wsPivot.Range("D6")
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    ' Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="Days-Hours1")
    ' Fields
    With pivotTable
        .PivotFields("Personnel Number").Orientation = xlRowField
        .PivotFields("Name (Last, First)").Orientation = xlRowField
        .PivotFields("Quota Description").Orientation = xlColumnField
        .PivotFields("Requested").Orientation = xlDataField
        If Err.Number <> 0 Then

         .PivotFields("Requested").Function = xlSum
            Debug.Print "Error with 'Requested': " & Err.Description
            .PivotFields("Requested").Function = xlCount
        End If
        On Error GoTo 0
        .PivotFields("Total Remaining").Orientation = xlDataField
        If Err.Number <> 0 Then
         .PivotFields("Total Remaining").Function = xlSum
            Debug.Print "Error with 'Total Remaining': " & Err.Description
            .PivotFields("Total Remaining").Function = xlCount
        End If
        On Error GoTo 0
            .RowAxisLayout xlTabularRow
            .RowGrand = False
            .ColumnGrand = False
            .SubtotalHiddenPageItems = False
         Dim pf As PivotField
   

        For Each pf In .RowFields
            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            pf.LayoutBlankLine = False
        Next pf
    End With
    If Err.Number = 0 Then
        Set sSlicerCache = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Quota Description")
        If Err.Number = 0 Then
            Set sSlicer = sSlicerCache.Slicers.Add(wsPivot.Name, , "Quota Description", "Quota Description", 15, 140)
            With sSlicer
                .Width = 350
                .Height = 50
                .NumberOfColumns = 3
                .RowHeight = 20
            End With
        Else
            MsgBox "Error creating slicer: " & Err.Description
        End If
    End If
    Set DataRange = pivotTable.DataBodyRange
    With DataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:=20)
        .Interior.Color = RGB(255, 192, 203) 
        .Font.Bold = True  
    End With
    MsgBox "Pivot Table created successfully!", vbInformation
End Sub

```

```vb  
Sub CreatePivotTable_Day_Hours_Available1()
    ' ... (keep your existing variable declarations)
    
    ' Create a combined range from multiple sheets
    Dim wsData2 As Worksheet
    Dim combinedRange As Range
    
    ' First data range (your existing one)
    Set wsData = ThisWorkbook.Worksheets("Quota RPT")
    Set pivotRange = wsData.Range("A1").CurrentRegion
    
    ' Second data range (from different sheet)
    Set wsData2 = ThisWorkbook.Worksheets("Your_Second_Sheet_Name") ' Change this to your second sheet name
    Dim secondRange As Range
    Set secondRange = wsData2.Range("A1").CurrentRegion
    
    ' Create a temporary sheet to combine data
    Dim wsCombined As Worksheet
    Set wsCombined = ThisWorkbook.Worksheets.Add
    
    ' Copy both ranges to the temporary sheet
    pivotRange.Copy wsCombined.Range("A1")
    secondRange.Copy wsCombined.Range("A" & pivotRange.Rows.Count + 2)
    
    ' Set the combined range
    Set combinedRange = wsCombined.Range("A1").CurrentRegion
    
    ' Create pivot cache using combined range
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=combinedRange)
    
    ' ... (keep your existing pivot table creation code)
    
    ' Add additional fields from the second dataset
    With pivotTable
        ' Existing fields
        .PivotFields("Personnel Number").Orientation = xlRowField
        .PivotFields("Name (Last, First)").Orientation = xlRowField
        .PivotFields("Quota Description").Orientation = xlColumnField
        .PivotFields("Requested").Orientation = xlDataField
        
        ' Add new fields from second dataset
        ' Example (modify field names as needed):
        '.PivotFields("New_Column1").Orientation = xlRowField
        '.PivotFields("New_Column2").Orientation = xlColumnField
        '.PivotFields("New_Column3").Orientation = xlDataField
        
        ' ... (rest of your existing pivot table formatting)
    End With
    
    ' Clean up - delete temporary sheet
    Application.DisplayAlerts = False
    wsCombined.Delete
    Application.DisplayAlerts = True
    
    ' ... (rest of your existing code)
End Sub
```


Combine All Data from two sheets 
```vb 
Sub CombineAllData()
    Dim ws1 As Worksheet, ws2 As Worksheet, wsNew As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim lastCol1 As Long, lastCol2 As Long
    Dim i As Long, j As Long
    Dim columnDict As Object
    Dim colName1 As String, colName2 As String
    Dim maxRows As Long
    
    Set columnDict = CreateObject("Scripting.Dictionary")

    ' Set your worksheets
    Set ws1 = ThisWorkbook.Sheets("Vacation Data")
    Set ws2 = ThisWorkbook.Sheets("Quota RPT")
   
    ' Create new sheet for combined data
    On Error Resume Next
    ThisWorkbook.Sheets("Combined").Delete
    On Error GoTo 0
    Set wsNew = ThisWorkbook.Sheets.Add
    wsNew.Name = "Combined"

    ' Find last rows and columns
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    ' Find the maximum number of rows between the two sheets
    maxRows = IIf(lastRow1 > lastRow2, lastRow1, lastRow2)

    ' Create headers and build column dictionary
    For i = 1 To lastCol1
        colName1 = ws1.Cells(1, i).Value
        If Not columnDict.Exists(colName1) And colName1 <> "" Then
            columnDict.Add colName1, columnDict.Count + 1
            wsNew.Cells(1, columnDict(colName1)).Value = colName1
        End If
    Next i

    For i = 1 To lastCol2
        colName2 = ws2.Cells(1, i).Value
        If Not columnDict.Exists(colName2) And colName2 <> "" Then
            columnDict.Add colName2, columnDict.Count + 1
            wsNew.Cells(1, columnDict(colName2)).Value = colName2
        End If
    Next i

    ' Copy data from both sheets simultaneously
    For i = 2 To maxRows
        ' Copy data from first sheet if row exists
        If i <= lastRow1 Then
            For j = 1 To lastCol1
                colName1 = ws1.Cells(1, j).Value
                If columnDict.Exists(colName1) Then
                    wsNew.Cells(i, columnDict(colName1)).Value = ws1.Cells(i, j).Value
                End If
            Next j
        End If
        
        ' Copy data from second sheet if row exists
        If i <= lastRow2 Then
            For j = 1 To lastCol2
                colName2 = ws2.Cells(1, j).Value
                If columnDict.Exists(colName2) Then
                    wsNew.Cells(i, columnDict(colName2)).Value = ws2.Cells(i, j).Value
                End If
            Next j
        End If
    Next i

    ' AutoFit columns
    wsNew.Columns.AutoFit

    ' Debug information
    Debug.Print "Last Row Sheet 1: " & lastRow1
    Debug.Print "Last Row Sheet 2: " & lastRow2
    Debug.Print "Max Rows: " & maxRows

    MsgBox "All data combined successfully!" & vbNewLine & _
           "Sheet 1 Rows: " & (lastRow1 - 1) & vbNewLine & _
           "Sheet 2 Rows: " & (lastRow2 - 1) & vbNewLine & _
           "Total Combined Rows: " & (maxRows - 1), vbInformation

End Sub
```

```vb
Sub CombineAllData()
    Dim ws1 As Worksheet, ws2 As Worksheet, wsNew As Worksheet, wsPivot As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastColNew As Long
    Dim lastCol1 As Long, lastCol2 As Long
    Dim i As Long, j As Long
    Dim columnDict As Object
    Dim colName1 As String, colName2 As String
    Dim maxRows As Long
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim sSlicerCache As SlicerCache
    Dim sSlicer As Slicer
    Dim DataRange As Range
    
    Set columnDict = CreateObject("Scripting.Dictionary")

    ' Set your worksheets
    Set ws1 = ThisWorkbook.Sheets("Vacation Data")
    Set ws2 = ThisWorkbook.Sheets("Quota RPT")
   
    ' Create new sheet for combined data
    On Error Resume Next
    ThisWorkbook.Sheets("TempCombined").Delete
    On Error GoTo 0
    Set wsNew = ThisWorkbook.Sheets.Add
    wsNew.Name = "TempCombined"

    ' Find last rows and columns
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    maxRows = IIf(lastRow1 > lastRow2, lastRow1, lastRow2)

    ' Create headers and build column dictionary
    For i = 1 To lastCol1
        colName1 = ws1.Cells(1, i).Value
        If Not columnDict.Exists(colName1) And colName1 <> "" Then
            columnDict.Add colName1, columnDict.Count + 1
            wsNew.Cells(1, columnDict(colName1)).Value = colName1
        End If
    Next i

    For i = 1 To lastCol2
        colName2 = ws2.Cells(1, i).Value
        If Not columnDict.Exists(colName2) And colName2 <> "" Then
            columnDict.Add colName2, columnDict.Count + 1
            wsNew.Cells(1, columnDict(colName2)).Value = colName2
        End If
    Next i

    ' Copy data from both sheets simultaneously
    For i = 2 To maxRows
        If i <= lastRow1 Then
            For j = 1 To lastCol1
                colName1 = ws1.Cells(1, j).Value
                If columnDict.Exists(colName1) Then
                    wsNew.Cells(i, columnDict(colName1)).Value = ws1.Cells(i, j).Value
                End If
            Next j
        End If
        
        If i <= lastRow2 Then
            For j = 1 To lastCol2
                colName2 = ws2.Cells(1, j).Value
                If columnDict.Exists(colName2) Then
                    wsNew.Cells(i, columnDict(colName2)).Value = ws2.Cells(i, j).Value
                End If
            Next j
        End If
    Next i

    ' Get the last column in the new sheet
    lastColNew = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column
    
    ' Set the range for pivot table source
    Set pivotRange = wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(maxRows, lastColNew))

    ' Create or activate pivot sheet
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("Days-Hours1")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "Days-Hours1"
    End If
    On Error GoTo 0

    ' Set pivot table destination
    Set pivotDestination = wsPivot.Range("D6")
    
    ' Create pivot cache and table
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="Days-Hours1")

    ' Configure pivot table
    With pivotTable
        .PivotFields("Personnel Number").Orientation = xlRowField
        .PivotFields("Name (Last, First)").Orientation = xlRowField
        .PivotFields("Quota Description").Orientation = xlColumnField
        
        On Error Resume Next
        .PivotFields("Requested").Orientation = xlDataField
        If Err.Number = 0 Then
            .PivotFields("Requested").Function = xlSum
        End If
        On Error GoTo 0

        On Error Resume Next
        .PivotFields("Total Remaining").Orientation = xlDataField
        If Err.Number = 0 Then
            .PivotFields("Total Remaining").Function = xlSum
        End If
        On Error GoTo 0

        .RowAxisLayout = xlTabularRow
        .RowGrand = False
        .ColumnGrand = False
        .SubtotalHiddenPageItems = False

        ' Remove subtotals
        Dim pf As PivotField
        For Each pf In .RowFields
            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            pf.LayoutBlankLine = False
        Next pf
    End With

    ' Add slicer
    On Error Resume Next
    Set sSlicerCache = ActiveWorkbook.SlicerCaches.Add2(pivotTable, "Quota Description")
    If Err.Number = 0 Then
        Set sSlicer = sSlicerCache.Slicers.Add(wsPivot.Name, , "Quota Description", "Quota Description", 15, 140)
        With sSlicer
            .Width = 350
            .Height = 50
            .NumberOfColumns = 3
            .RowHeight = 20
        End With
    End If
    On Error GoTo 0

    ' Add conditional formatting
    Set DataRange = pivotTable.DataBodyRange
    With DataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="20")
        .Interior.Color = RGB(255, 192, 203)
        .Font.Bold = True
    End With

    ' Delete the temporary combined data sheet
    Application.DisplayAlerts = False
    wsNew.Delete
    Application.DisplayAlerts = True

    MsgBox "Process completed successfully!", vbInformation
End Sub
```

```vb
With pivotTable
    .PivotFields("Personnel Number").Orientation = xlRowField
    .PivotFields("Name (Last, First)").Orientation = xlRowField
    .PivotFields("Quota Description").Orientation = xlColumnField
    
    ' Add the individual fields first
    .PivotFields("Requested").Orientation = xlDataField
    .PivotFields("Total Remaining").Orientation = xlDataField
    
    ' Create the calculated field
    On Error Resume Next
    .CalculatedFields.Add "Total Hours", "='Requested'+'Total Remaining'", True
    If Err.Number = 0 Then
        ' Add the calculated field to the pivot table
        .PivotFields("Total Hours").Orientation = xlDataField
    Else
        Debug.Print "Error creating calculated field: " & Err.Description
    End If
    On Error GoTo 0
    
    ' Continue with the rest of your pivot table formatting
    .RowAxisLayout xlTabularRow
    .RowGrand = False
    .ColumnGrand = False
    .SubtotalHiddenPageItems = False
End With
```

```vb
Sub CreatePivotTable()

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotTable2 As PivotTable
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
  
    ' Delete existing sheet if it exists
    On Error Resume Next
    ThisWorkbook.Sheets("SP_TEST").Delete
    On Error GoTo 0

    ' Create a new sheet for the pivot table
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("SP_TEST")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "SP_TEST"
    End If
    On Error GoTo 0

    ' Destination for the pivot table
    Set pivotDestination = wsPivot.Range("A7")
    Set pivotDestination2 = wsPivot.Range("A55")

    ' Create PivotCache and PivotTables
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotCache2 = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange2)

    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="AC_MyPivotTable")
    pivotTable.TableStyle2 = "PivotStyleMedium15"

    Set pivotTable2 = pivotCache2.CreatePivotTable(TableDestination:=pivotDestination2, TableName:="JIRA_MyPivotTable2")
    pivotTable2.TableStyle2 = "PivotStyleMedium15"

    ' Fields for the first pivot table
    With pivotTable
        .PivotFields("Function").Orientation = xlRowField
        .PivotFields("WP").Orientation = xlRowField
        .AddDataField .PivotFields("ETC JIRA SP1.1"), "ETC JIRA SP1.1", xlMax
        .AddDataField .PivotFields("Week 1 "), "Week 1", xlSum
        .AddDataField .PivotFields("Week 2 "), "Week 2", xlSum
        .AddDataField .PivotFields("AC SP1.1"), "AC SP1.1", xlSum

        ' Add blank column after the main data fields
        wsPivot.Columns("F:F").Insert Shift:=xlToRight
        wsPivot.Range("F6").Value = " " ' Add a header for the blank column (optional)

        ' Add additional fields after the blank column
        .AddDataField .PivotFields("Total Issue SP1.1"), "Total Issue SP1.1", xlSum
        .AddDataField .PivotFields("Closed SP1.1"), "Closed SP1.1", xlSum
        .AddDataField .PivotFields("Resolved SP1.1"), "Resolved SP1.1", xlSum

        ' Add another blank column
        wsPivot.Columns("J:J").Insert Shift:=xlToRight
        wsPivot.Range("J6").Value = " " ' Add a header for the blank column (optional)

        ' Add remaining fields
        .AddDataField .PivotFields("Week 3 "), "Week 3", xlSum
        .AddDataField .PivotFields("Week 4 "), "Week 4", xlSum
        .AddDataField .PivotFields("AC SP1.2"), "AC SP1.2", xlSum
    End With

    ' Add blank column for the second pivot table
    wsPivot.Columns("Z:Z").Insert Shift:=xlToRight
    wsPivot.Range("Z6").Value = " " ' Add a header for the blank column (optional)

    ' Fields for the second pivot table
    With pivotTable2
        .PivotFields("Epic Link").Orientation = xlRowField
        .AddDataField .PivotFields("EV"), "EV", xlSum
        .RowAxisLayout xlTabularRow
    End With

    ' Final formatting and messages
    MsgBox "AC Dec_Pivot Table created successfully!", vbInformation

End Sub
```

```vb
Sub CreatePivot_DumpData()

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
    Dim timelineCache As SlicerCache
    Dim timeline As Slicer
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
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
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
                .Width = 150
                .Height = 58
                .NumberOfColumns = 2
                .RowHeight = 20
            End With
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
                .Width = 256
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
                .Width = 150
                .Height = 58
                .NumberOfColumns = 2
                .RowHeight = 20
            End With
        Else
            MsgBox "Error creating third slicer: " & Err.Description
        End If
    End If

    On Error GoTo 0

    ' Add the timeline slicer for FiscalMonth
    On Error Resume Next
    Set timelineCache = ThisWorkbook.SlicerCaches.Add2(pivotTable, "FiscalMonth")
    If Not timelineCache Is Nothing Then
        ' Add the timeline slicer to the wsPivot worksheet
        Set timeline = timelineCache.Slicers.Add(wsPivot, , "FiscalMonth", "FiscalMonth", 0, 0, 400, 50)
        If Not timeline Is Nothing Then
            ' Position the timeline slicer at H2
            With timeline
                .Top = wsPivot.Range("H2").Top
                .Left = wsPivot.Range("H2").Left
            End With
        Else
            MsgBox "Error creating the timeline slicer for FiscalMonth.", vbExclamation
        End If
    Else
        MsgBox "Error creating the timeline cache for FiscalMonth.", vbExclamation
    End If
    On Error GoTo 0

    MsgBox "Pivot Table with Slicers created successfully!", vbInformation

End Sub
```