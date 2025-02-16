Pivot_Report
```vb
Sub CreatePivotTableAndTotal()

    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim pivotCache As PivotCache, pivotTable As PivotTable
    Dim pivotRange As Range, pivotDestination As Range
    Dim monthName As String
    Dim lastColumn As Long, lastRow As Long, lastCol As Long
    Dim cell As Range, field As PivotField, pf As PivotField
    Dim totalValue As Double
    
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
        .CalculatedFields.Add "AC/ETC week3-4", "=IFERROR('AC week3-4'/'ETC JIRA', 0)"

        ' Set calculated fields orientation
        .PivotFields("EV,%").Orientation = xlDataField
        .PivotFields("AC/ETC week1-2").Orientation = xlDataField
        .PivotFields("AC/ETC week3-4").Orientation = xlDataField

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

    ' ✅ Adjust total formulas: SUM for numbers, leave percentage fields empty, hide zeros
    For Each cell In wsPivot.Range(wsPivot.Cells(3, 3), wsPivot.Cells(3, lastColumn))
        lastRow = wsPivot.Cells(wsPivot.Rows.Count, cell.Column).End(xlUp).Row
        
        ' Get the column header name
        Dim columnHeader As String
        columnHeader = wsPivot.Cells(4, cell.Column).Value
        
        ' Check if column is a percentage field
        Select Case columnHeader
            Case "EV,% ", "AC/ETC week1-2 ", "AC/ETC week3-4 "
                ' ✅ LEAVE EMPTY (No formula for percentage fields)
                cell.Value = ""
            Case Else
                ' ✅ Use SUM for numeric fields
                cell.Formula = "=SUM(" & wsPivot.Cells(4, cell.Column).Address & ":" & wsPivot.Cells(lastRow, cell.Column).Address & ")"
                
                ' ✅ Hide 0 values
                totalValue = cell.Value
                If totalValue = 0 Then cell.Value = ""
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
War to DATA
```vb
Sub WAR_Pivot_To_Data()

    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim pt As PivotTable, newPt As PivotTable
    Dim pf As PivotField, dataRange As Range
    Dim tbl As ListObject, cell As Range
    Dim monthName As String
    Dim lastRow As Long, i As Integer

    ' Get the current month name
    monthName = Format(Date, "mmm")

    ' Set source worksheet
    Set wsSource = ThisWorkbook.Sheets("Current Month ETC vs ACWP")
    Set pt = wsSource.PivotTables("PivotTable4")

    ' Delete existing WAR sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("WAR " & monthName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Add new WAR sheet
    Set wsTarget = ThisWorkbook.Sheets.Add
    wsTarget.Name = "WAR " & monthName

    ' Copy PivotTable data
    pt.TableRange2.Copy
    wsTarget.Range("A1").PasteSpecial Paste:=xlPasteAll

    ' Get new pivot table if it exists
    On Error Resume Next
    Set newPt = wsTarget.PivotTables(1)
    On Error GoTo 0

    ' Format PivotTable
    If Not newPt Is Nothing Then
        With newPt
            .RowAxisLayout xlTabularRow
            .RepeatAllLabels xlRepeatLabels
            .ShowTableStyleRowStripes = False
            .ShowTableStyleColumnStripes = False
            .ColumnGrand = False
            .RowGrand = False
            For Each pf In .RowFields
                pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            Next pf
        End With
    End If

    ' Unmerge cells
    wsTarget.Cells.UnMerge
    wsTarget.Columns("A").UnMerge
    wsTarget.Columns("A:B").UnMerge

    ' Align text to the left
    wsTarget.Columns("A:C").HorizontalAlignment = xlLeft

    ' Freeze panes at row 4
    wsTarget.Rows("4:4").Select
    ActiveWindow.FreezePanes = True

    ' Convert PivotTable to values
    Set dataRange = wsTarget.Range("A1").CurrentRegion
    dataRange.Copy
    dataRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Delete the first two rows
    wsTarget.Rows("1:2").Delete

    ' Delete unnecessary columns
    wsTarget.Range("D:K,Q:X").Delete

    ' Convert the remaining data to a table
    Set tbl = wsTarget.ListObjects.Add(xlSrcRange, wsTarget.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = "WARDataTable"
    tbl.TableStyle = "TableStyleLight8"

    ' Insert columns after "Charge Number"
    Dim chargeNumberCol As ListColumn
    Set chargeNumberCol = tbl.ListColumns("Charge Number")
    chargeNumberCol.Range.Offset(0, 1).EntireColumn.Insert
    chargeNumberCol.Range.Offset(0, 1).EntireColumn.Insert

    ' Name the new columns "WP" and "WP Description"
    chargeNumberCol.Range.Offset(0, 1).Cells(1, 1).Value = "WP"
    chargeNumberCol.Range.Offset(0, 2).Cells(1, 1).Value = "WP Description"

    ' Insert columns after "Week 5"
    Dim week5Col As ListColumn
    Set week5Col = tbl.ListColumns("Week 5 ")
    week5Col.Range.Offset(0, 1).EntireColumn.Insert
    week5Col.Range.Offset(0, 1).EntireColumn.Insert
    week5Col.Range.Offset(0, 1).EntireColumn.Insert
    week5Col.Range.Offset(0, 1).EntireColumn.Insert
    week5Col.Range.Offset(0, 1).EntireColumn.Insert
    week5Col.Range.Offset(0, 1).EntireColumn.Insert

    ' Name the new columns
    week5Col.Range.Offset(0, 1).Cells(1, 1).Value = "AC week1-2"
    week5Col.Range.Offset(0, 2).Cells(1, 1).Value = "AC week3-4"
    week5Col.Range.Offset(0, 3).Cells(1, 1).Value = "AC month"
    week5Col.Range.Offset(0, 4).Cells(1, 1).Value = "ETC JIRA"
    week5Col.Range.Offset(0, 5).Cells(1, 1).Value = "EV"
    week5Col.Range.Offset(0, 6).Cells(1, 1).Value = "Notes from PEs"

    ' Find column indexes dynamically
    Dim week1ColIndex As Integer, week2ColIndex As Integer
    Dim week3ColIndex As Integer, week4ColIndex As Integer
    Dim totalWeek12SPColIndex As Integer, totalWeek34SPColIndex As Integer
    Dim totalMonthSPColIndex As Integer
    Dim headerRow As Range
    Set headerRow = wsTarget.Rows(1)

    ' Get column indexes
    Dim col As Range
    For Each col In headerRow.Cells
        Select Case col.Value
            Case "Week 1 ": week1ColIndex = col.Column
            Case "Week 2 ": week2ColIndex = col.Column
            Case "Week 3 ": week3ColIndex = col.Column
            Case "Week 4 ": week4ColIndex = col.Column
            Case "AC week1-2": totalWeek12SPColIndex = col.Column
            Case "AC week3-4": totalWeek34SPColIndex = col.Column
            Case "AC month": totalMonthSPColIndex = col.Column
        End Select
    Next col

    ' Ensure required columns exist
    If week1ColIndex = 0 Or week2ColIndex = 0 Or totalWeek12SPColIndex = 0 Then Exit Sub
    If week3ColIndex = 0 Or week4ColIndex = 0 Or totalWeek34SPColIndex = 0 Then Exit Sub
    If totalMonthSPColIndex = 0 Then Exit Sub

    ' Get last row
    lastRow = wsTarget.Cells(Rows.Count, "A").End(xlUp).Row

    ' Insert formulas
    wsTarget.Range(wsTarget.Cells(2, totalWeek12SPColIndex), wsTarget.Cells(lastRow, totalWeek12SPColIndex)).FormulaR1C1 = "=RC" & week1ColIndex & " + RC" & week2ColIndex
    wsTarget.Range(wsTarget.Cells(2, totalWeek34SPColIndex), wsTarget.Cells(lastRow, totalWeek34SPColIndex)).FormulaR1C1 = "=RC" & week3ColIndex & " + RC" & week4ColIndex
    wsTarget.Range(wsTarget.Cells(2, totalMonthSPColIndex), wsTarget.Cells(lastRow, totalMonthSPColIndex)).FormulaR1C1 = "=RC" & week1ColIndex & " + RC" & week2ColIndex & " + RC" & week3ColIndex & " + RC" & week4ColIndex

    ' Convert formulas to values
    wsTarget.Range(wsTarget.Cells(2, totalWeek12SPColIndex), wsTarget.Cells(lastRow, totalWeek12SPColIndex)).Value = wsTarget.Range(wsTarget.Cells(2, totalWeek12SPColIndex), wsTarget.Cells(lastRow, totalWeek12SPColIndex)).Value
    wsTarget.Range(wsTarget.Cells(2, totalWeek34SPColIndex), wsTarget.Cells(lastRow, totalWeek34SPColIndex)).Value = wsTarget.Range(wsTarget.Cells(2, totalWeek34SPColIndex), wsTarget.Cells(lastRow, totalWeek34SPColIndex)).Value
    wsTarget.Range(wsTarget.Cells(2, totalMonthSPColIndex), wsTarget.Cells(lastRow, totalMonthSPColIndex)).Value = wsTarget.Range(wsTarget.Cells(2, totalMonthSPColIndex), wsTarget.Cells(lastRow, totalMonthSPColIndex)).Value

    ' AutoFit columns
    wsTarget.Columns("A:M").AutoFit

    ' Success message
    MsgBox "WAR Pivot to Data process completed successfully!", vbInformation

End Sub
```