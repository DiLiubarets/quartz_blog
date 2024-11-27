
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

    ' Set the data sheet and range
    Set wsData = ThisWorkbook.Worksheets("Sheet1") ' Replace "Sheet1" with your data sheet name
    Set pivotRange = wsData.Range("A1").CurrentRegion ' Adjust range if needed

    ' Debug: Check the range
    Debug.Print "Pivot Range: " & pivotRange.Address

    ' Add a new sheet for the Pivot Tables
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("PivotTable")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Worksheets.Add
        wsPivot.Name = "PivotTable"
    End If
    wsPivot.Cells.Clear ' Clear the sheet to avoid conflicts
    On Error GoTo 0

    ' Set the destinations for the Pivot Tables
    Set pivotDestination1 = wsPivot.Range("A3") ' First Pivot Table starts at A3
    Set pivotDestination2 = wsPivot.Range("G3") ' Second Pivot Table starts at G3 (adjust as needed)

    ' Create the Pivot Cache (shared for both Pivot Tables)
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    ' Debug: Check the cache source
    Debug.Print "Pivot Cache Source: " & pivotCache.SourceData

    ' Create the First Pivot Table
    Set pivotTable1 = pivotCache.CreatePivotTable(TableDestination:=pivotDestination1, TableName:="MyPivotTable1")
    With pivotTable1
        .PivotFields("Category").Orientation = xlRowField ' Replace "Category" with your column name
        .PivotFields("Region").Orientation = xlColumnField ' Replace "Region" with your column name

        ' Check if "Sales" is numeric before applying xlSum
        On Error Resume Next
        .PivotFields("Sales").Orientation = xlDataField ' Replace "Sales" with your column name
        .PivotFields("Sales").Function = xlSum ' Summarize as SUM
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Sales': " & Err.Description
            .PivotFields("Sales").Function = xlCount ' Use Count as fallback
        End If
        On Error GoTo 0
    End With

    ' Create the Second Pivot Table
    Set pivotTable2 = pivotCache.CreatePivotTable(TableDestination:=pivotDestination2, TableName:="MyPivotTable2")
    With pivotTable2
        .PivotFields("Region").Orientation = xlRowField ' Replace "Region" with your column name
        .PivotFields("Category").Orientation = xlColumnField ' Replace "Category" with your column name

        ' Check if "Profit" is numeric before applying xlSum
        On Error Resume Next
        .PivotFields("Profit").Orientation = xlDataField ' Replace "Profit" with your column name
        .PivotFields("Profit").Function = xlSum ' Summarize as SUM
        If Err.Number <> 0 Then
            Debug.Print "Error with 'Profit': " & Err.Description
            .PivotFields("Profit").Function = xlCount ' Use Count as fallback
        End If
        On Error GoTo 0
    End With

    MsgBox "Two Pivot Tables created successfully!", vbInformation
End Sub
```