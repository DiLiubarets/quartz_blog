Pivot_Report
```vb
Sub CreatePivotTableAndTotal()

    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim pivotCache As PivotCache, pivotTable As PivotTable
    Dim pivotRange As Range, pivotDestination As Range
    Dim monthName As String
    Dim lastColumn As Long, lastRow As Long, lastCol As Long
    Dim cell As Range, field As PivotField, pf As PivotField

    ' Get the current month name
    monthName = Format(Date, "mmm")

    ' Set the data worksheet & range
    Set wsData = ThisWorkbook.Worksheets("WAR " & monthName)
    Set pivotRange = wsData.Range("A1").CurrentRegion

    ' Delete existing "Report" sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Report").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Add new "Report" sheet
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "Report"

    ' Set pivot table destination
    Set pivotDestination = wsPivot.Range("B4")

    ' Create Pivot Cache & PivotTable
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
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

        ' ✅ Add calculated fields
        .CalculatedFields.Add "EV,%", "=IFERROR(EV/ETC JIRA, 0)"
        .CalculatedFields.Add "AC/ETC week1-2", "=IFERROR('AC week1-2'/'ETC JIRA', 0)"
        .CalculatedFields.Add "AC/ETC week3-4", "=IFERROR('AC week3-4'/'ETC JIRA', 0)" ' ✅ New field added

        ' Set calculated fields orientation
        .PivotFields("EV,%").Orientation = xlDataField
        .PivotFields("AC/ETC week1-2").Orientation = xlDataField
        .PivotFields("AC/ETC week3-4").Orientation = xlDataField ' ✅ Display new field

        ' Format PivotTable
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleMedium15"
        .DisplayFieldCaptions = False
        .ColumnGrand = False
        .RowGrand = False
    End With

    ' ✅ Disable subtotals only for applicable fields
    For Each pf In pivotTable.RowFields
        pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next pf

    ' Rename fields
    Dim fieldNames As Variant, newNames As Variant, i As Integer
    fieldNames = Array("Sum of ETC JIRA", "Sum of EV", "Sum of AC week1-2", "Sum of AC week3-4", "Sum of AC month", "Sum of EV,%", "Sum of AC/ETC week1-2", "Sum of AC/ETC week3-4")
    newNames = Array("ETC JIRA ", "EV ", "AC week1-2 ", "AC week3-4 ", "AC month ", "EV,% ", "AC/ETC week1-2 ", "AC/ETC week3-4 ")

    For i = LBound(fieldNames) To UBound(fieldNames)
        On Error Resume Next
        pivotTable.PivotFields(fieldNames(i)).Caption = newNames(i)
        On Error GoTo 0
    Next i

    ' Explicitly set number format for percentage fields
    For Each field In pivotTable.DataFields
        Select Case field.Name
            Case "EV,% ", "AC/ETC week1-2 ", "AC/ETC week3-4 "
                field.NumberFormat = "0.00%"
        End Select
    Next field

    ' Find last column in the PivotTable
    lastColumn = wsPivot.Cells(4, wsPivot.Columns.Count).End(xlToLeft).Column

    ' Insert "Total" row above PivotTable (Row 3)
    With wsPivot.Cells(3, 2)
        .Value = "Total"
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0) ' Black background
        .Font.Color = RGB(255, 255, 255) ' White font color
    End With

    ' ✅ Adjust total formulas: SUM for numbers, leave percentage fields empty
    For Each cell In wsPivot.Range(wsPivot.Cells(3, 3), wsPivot.Cells(3, lastColumn))
        lastRow = wsPivot.Cells(wsPivot.Rows.Count, cell.Column).End(xlUp).Row
        
        ' Get the column header name
        Dim columnHeader As String
        columnHeader = wsPivot.Cells(4, cell.Column).Value
        
        ' Check if column is a percentage field
        Select Case columnHeader
            Case "EV,% ", "AC/ETC week1-2 ", "AC/ETC week3-4 "
                ' ✅ LEAVE EMPTY (No formula)
                cell.Value = ""
            Case Else
                ' ✅ Use SUM for numeric fields
                cell.Formula = "=SUM(" & wsPivot.Cells(4, cell.Column).Address & ":" & wsPivot.Cells(lastRow, cell.Column).Address & ")"
        End Select

        ' Apply formatting
        cell.Font.Bold = True
        cell.Interior.Color = RGB(0, 0, 0)
        cell.Font.Color = RGB(255, 255, 255)
    Next cell

    ' Insert "Notes" column after last pivot column
    lastCol = pivotTable.TableRange2.Columns.Count
    wsPivot.Cells(4, lastCol + 2).EntireColumn.Insert
    wsPivot.Cells(4, lastCol + 2).Value = "Notes"

    ' Copy formatting from last column to "Notes" column
    wsPivot.Cells(3, lastCol).EntireColumn.Copy
    wsPivot.Cells(3, lastCol + 1).EntireColumn.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' ✅ AutoFit all columns
    wsPivot.Cells.EntireColumn.AutoFit

    ' Success message
    MsgBox "Pivot Table and Total row created successfully!", vbInformation

End Sub
```
