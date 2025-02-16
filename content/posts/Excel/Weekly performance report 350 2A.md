Pivot_Report
```vb
Sub CreatePivotTableAndTotal()

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim monthName As String
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim cell As Range
    Dim field As PivotField
    Dim lastCol As Long

    ' Get the current month name
    monthName = Format(Date, "mmm")

    ' Set the data worksheet
    Set wsData = ThisWorkbook.Worksheets("WAR" & " " & monthName)
    Set pivotRange = wsData.Range("A1").CurrentRegion

    ' Delete existing "Report" sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsPivot = ThisWorkbook.Worksheets("Report")
    If Not wsPivot Is Nothing Then
        wsPivot.Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Add new "Report" sheet
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "Report"

    ' Set pivot table destination (Now starts at B4)
    Set pivotDestination = wsPivot.Range("B4")

    ' Create Pivot Cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    ' Create PivotTable
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTable")

    ' Add fields to PivotTable
    With pivotTable
        .PivotFields("WP").Orientation = xlRowField
        .PivotFields("WP Description").Orientation = xlRowField
        .PivotFields("ETC JIRA").Orientation = xlDataField
        .PivotFields("EV").Orientation = xlDataField
        .PivotFields("AC week1-2").Orientation = xlDataField
        .PivotFields("AC week3-4").Orientation = xlDataField
        .PivotFields("AC month").Orientation = xlDataField
    End With

    ' Add calculated field for EV,%
    pivotTable.CalculatedFields.Add "EV,%", "=IFERROR(EV/ETC JIRA, 0)"
    pivotTable.PivotFields("EV,%").Orientation = xlDataField
    pivotTable.PivotFields("EV,%").NumberFormat = "0.00%"

    ' Add calculated field for AC/ETC week1-2
    pivotTable.CalculatedFields.Add "AC/ETC week1-2", "=IFERROR('AC week1-2'/'ETC JIRA', 0)"
    pivotTable.PivotFields("AC/ETC week1-2").Orientation = xlDataField
    pivotTable.PivotFields("AC/ETC week1-2").NumberFormat = "0.00%"

    ' Format PivotTable
    With pivotTable
        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = False
        .ShowTableStyleColumnStripes = False
        .ShowTableStyleLastColumn = False
        .ShowTableStyleRowHeaders = False
        .ShowTableStyleColumnHeaders = True
        .ColumnGrand = False
        .RowGrand = False
        .TableStyle2 = "PivotStyleMedium15"
        .DisplayFieldCaptions = False ' Hide field captions
    End With

    ' Remove subtotals
    With pivotTable
        .PivotFields("WP").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("WP Description").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With

    ' Rename fields to remove "Sum of"
    pivotTable.PivotFields("Sum of ETC JIRA").Caption = "ETC JIRA "
    pivotTable.PivotFields("Sum of EV").Caption = "EV "
    pivotTable.PivotFields("Sum of AC week1-2").Caption = "AC week1-2 "
    pivotTable.PivotFields("Sum of AC week3-4").Caption = "AC week3-4 "
    pivotTable.PivotFields("Sum of AC month").Caption = "AC month "
    pivotTable.PivotFields("Sum of EV,%").Caption = "EV,% "
    pivotTable.PivotFields("Sum of AC/ETC week1-2").Caption = "AC/ETC week1-2 " ' Added space at the end

    ' Explicitly set the number format again to ensure it's applied
    For Each field In pivotTable.DataFields
        If field.Name = "EV,% " Or field.Name = "AC/ETC week1-2 " Or field.Name = "AC/ETC week3-4 " Then
            field.NumberFormat = "0.00%"
        End If
    Next field

    ' Find last column in the pivot table
    lastColumn = wsPivot.Cells(4, wsPivot.Columns.Count).End(xlToLeft).Column ' Adjusted for new starting position

    ' Set the value and style for the "Total" label (Now in row 3)
    With wsPivot.Cells(3, 2)
        .Value = "Total"
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0) ' Black background
        .Font.Color = RGB(255, 255, 255) ' White font color
    End With

    ' Set the formula and style for the total cells dynamically (Now in row 3)
    For Each cell In wsPivot.Range(wsPivot.Cells(3, 3), wsPivot.Cells(3, lastColumn))
        lastRow = wsPivot.Cells(wsPivot.Rows.Count, cell.Column).End(xlUp).Row ' Find the last row with data in the current column
        cell.Formula = "=SUM(" & wsPivot.Cells(4, cell.Column).Address & ":" & wsPivot.Cells(lastRow, cell.Column).Address & ")"
        cell.Font.Bold = True
        cell.Interior.Color = RGB(0, 0, 0) ' Black background
        cell.Font.Color = RGB(255, 255, 255) ' White font color
    Next cell

    ' Insert a column named "Notes" after the last column of the pivot table in row 4
    lastCol = pivotTable.TableRange2.Columns.Count
    wsPivot.Cells(4, lastCol + 2).EntireColumn.Insert
    wsPivot.Cells(4, lastCol + 2).Value = "Notes"

    ' Copy the style from the last column of the pivot table to the "Notes" column
    wsPivot.Cells(3, lastCol).EntireColumn.Copy
    wsPivot.Cells(3, lastCol + 1).EntireColumn.PasteSpecial Paste:=xlPasteFormats

    Application.CutCopyMode = False

    ' Display success message
    MsgBox "Pivot Table and Total row created successfully!", vbInformation

End Sub
```
