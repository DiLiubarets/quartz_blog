
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
    Dim slicerCache1 As SlicerCache
    Dim slicerCache2 As SlicerCache

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

    ' Add Slicers
    On Error Resume Next
    Set slicerCache1 = ThisWorkbook.SlicerCaches.Add(pivotTable, "Time off Description")
    slicerCache1.Slicers.Add wsPivot, , "Time off Description Slicer", wsPivot.Range("H3"), 150, 200, 100, 200

    Set slicerCache2 = ThisWorkbook.SlicerCaches.Add(pivotTable, "Indirect/Direct")
    slicerCache2.Slicers.Add wsPivot, , "Indirect/Direct Slicer", wsPivot.Range("H10"), 150, 200, 100, 200
    On Error GoTo 0

    MsgBox "Pivot Table with Slicers created successfully!", vbInformation

End Sub
```