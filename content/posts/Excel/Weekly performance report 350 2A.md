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

    Dim wsSource As Worksheet

    Dim wsTarget As Worksheet

    Dim pt As pivotTable

    Dim pf As PivotField

    Dim dataRange As Range

    Dim newPt As pivotTable

    Dim tbl As ListObject

    Dim cell As Range

    Dim monthName As String

    ' Get the month name first

    monthName = Format(Date, "mmm")

    ' Set source worksheet

    Set wsSource = ThisWorkbook.Sheets("Current Month ETC vs ACWP")

    Set pt = wsSource.PivotTables("PivotTable4")

    ' Delete the sheet if it exists

    Application.DisplayAlerts = False

    On Error Resume Next

    ThisWorkbook.Sheets("WAR" & " " & monthName).Delete

    On Error GoTo 0

    Application.DisplayAlerts = True

    ' Add new sheet

    Set wsTarget = ThisWorkbook.Sheets.Add

    wsTarget.Name = "WAR" & " " & monthName

    If wsTarget Is Nothing Then

        MsgBox "Error: Target worksheet not set."

        Exit Sub

    End If

    ' Copy PivotTable data

    pt.TableRange2.Copy

    wsTarget.Range("A1").PasteSpecial Paste:=xlPasteAll

    ' Get new pivot table if it exists

    On Error Resume Next

    Set newPt = wsTarget.PivotTables(1)

    On Error GoTo 0

    If newPt Is Nothing Then

        MsgBox "Error: PivotTable not set in the new worksheet."

        Exit Sub

    End If

    ' Format PivotTable

    With newPt

        .RowAxisLayout xlTabularRow

        .RepeatAllLabels xlRepeatLabels

        .ShowTableStyleRowStripes = False

        .ShowTableStyleColumnStripes = False

        For Each pf In .RowFields

            pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

        Next pf

        .ColumnGrand = False

        .RowGrand = False

    End With

    wsTarget.Cells.UnMerge

    wsTarget.Columns("A").UnMerge

    wsTarget.Columns("A:B").UnMerge

    Application.CutCopyMode = False

    wsTarget.Columns("A:C").HorizontalAlignment = xlLeft

    ' Freeze panes

    wsTarget.Rows("4:4").Select

    ActiveWindow.FreezePanes = True

    ' Copy PivotTable data and paste as values

    Set dataRange = newPt.TableRange2

    dataRange.Copy

    dataRange.PasteSpecial Paste:=xlPasteValues

    ' Ensure the data range is set

    If dataRange Is Nothing Then

        MsgBox "Error: Data range not set."

        Exit Sub

    End If

    ' Delete the first two rows

    wsTarget.Rows("1:2").Delete

    ' Delete columns D:K and Q:X

    wsTarget.Range("D:K,Q:X").Delete

    ' Convert the remaining data to a table

    Set tbl = wsTarget.ListObjects.Add(xlSrcRange, wsTarget.Range("A1").CurrentRegion, , xlYes)

    tbl.Name = "WARDataTable"

    tbl.TableStyle = "TableStyleLight8"

    ' Insert column after "Charge Number"

    Dim chargeNumberCol As ListColumn

    Set chargeNumberCol = tbl.ListColumns("Charge Number")

    chargeNumberCol.Range.Offset(0, 1).EntireColumn.Insert

    chargeNumberCol.Range.Offset(0, 1).EntireColumn.Insert

    ' Name the new columns "WP" and "Type of Work"

    chargeNumberCol.Range.Offset(0, 1).Cells(1, 1).Value = "WP"

    chargeNumberCol.Range.Offset(0, 2).Cells(1, 1).Value = "WP Description"

    ' Insert three columns after "Week 5"

    Dim week5Col As ListColumn

    Set week5Col = tbl.ListColumns("Week 5 ")

    week5Col.Range.Offset(0, 1).EntireColumn.Insert

    week5Col.Range.Offset(0, 1).EntireColumn.Insert

    week5Col.Range.Offset(0, 1).EntireColumn.Insert

    week5Col.Range.Offset(0, 1).EntireColumn.Insert

    week5Col.Range.Offset(0, 1).EntireColumn.Insert

    week5Col.Range.Offset(0, 1).EntireColumn.Insert

    ' Name the new columns "AC week1-2", "AC week3-4", and "AC month"

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

    Dim lastRow As Long

    Dim headerRow As Range

    ' Get the header row

    Set headerRow = wsTarget.Rows(1)

    ' Find column indexes

    Dim col As Range

    For Each col In headerRow.Cells

        Select Case col.Value

            Case "Week 1 "

                week1ColIndex = col.Column

            Case "Week 2 "

                week2ColIndex = col.Column

            Case "Week 3 "

                week3ColIndex = col.Column

            Case "Week 4 "

                week4ColIndex = col.Column

            Case "AC week1-2"

                totalWeek12SPColIndex = col.Column

            Case "AC week3-4"

                totalWeek34SPColIndex = col.Column

            Case "AC month"

                totalMonthSPColIndex = col.Column

        End Select

    Next col

    ' Check if required columns exist

    If week1ColIndex = 0 Or week2ColIndex = 0 Or totalWeek12SPColIndex = 0 Then

        MsgBox "Error: Could not find required columns for AC week1-2.", vbCritical

        Exit Sub

    End If

    If week3ColIndex = 0 Or week4ColIndex = 0 Or totalWeek34SPColIndex = 0 Then

        MsgBox "Error: Could not find required columns for AC week3-4.", vbCritical

        Exit Sub

    End If

    If totalMonthSPColIndex = 0 Then

        MsgBox "Error: Could not find required column for AC month.", vbCritical

        Exit Sub

    End If

    ' Get last row

    lastRow = wsTarget.Cells(Rows.Count, "A").End(xlUp).Row

    ' Insert formula in "AC week1-2" column

    Dim formulaRange12 As Range

    Set formulaRange12 = wsTarget.Range(wsTarget.Cells(2, totalWeek12SPColIndex), wsTarget.Cells(lastRow, totalWeek12SPColIndex))

    formulaRange12.FormulaR1C1 = "=(RC" & week1ColIndex & " + RC" & week2ColIndex & ") " '/ 4

    ' Insert formula in "AC week3-4" column

    Dim formulaRange34 As Range

    Set formulaRange34 = wsTarget.Range(wsTarget.Cells(2, totalWeek34SPColIndex), wsTarget.Cells(lastRow, totalWeek34SPColIndex))

    formulaRange34.FormulaR1C1 = "=(RC" & week3ColIndex & " + RC" & week4ColIndex & ") " '/ 4

    ' Insert formula in "AC month" column

    Dim formulaRangeMonth As Range

    Set formulaRangeMonth = wsTarget.Range(wsTarget.Cells(2, totalMonthSPColIndex), wsTarget.Cells(lastRow, totalMonthSPColIndex))

    formulaRangeMonth.FormulaR1C1 = "=(RC" & week1ColIndex & " + RC" & week2ColIndex & " + RC" & week3ColIndex & " + RC" & week4ColIndex & ") " '/ 4

    ' Convert formulas to values

    formulaRange12.Copy

    formulaRange12.PasteSpecial Paste:=xlPasteValues

    formulaRange34.Copy

    formulaRange34.PasteSpecial Paste:=xlPasteValues

    formulaRangeMonth.Copy

    formulaRangeMonth.PasteSpecial Paste:=xlPasteValues

    Application.CutCopyMode = False

    InsertXLOOKUP_WP.InsertXLOOKUP_WP

    InsertXLOOKUP_WP.InsertXLOOKUP_WP_Description

    InsertFormula.InsertFormula_ETC

    InsertFormula.InsertFormula_EV

    For i = tbl.ListRows.Count To 1 Step -1

        With tbl.ListRows(i).Range

            If .Cells(1, tbl.ListColumns("Employee Name").Index).Value = "ETC 1" Or _

               .Cells(1, tbl.ListColumns("Employee Name").Index).Value = "ETC 2" Then

                .Delete

            End If

        End With

    Next i

     wsTarget.Columns("A:M").AutoFit

    MsgBox "Process completed successfully!", vbInformation

End Sub
```

test for notes 
```vb 
Sub InsertXLOOKUPFormula()

    Dim wsWarFeb As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim formula As String
    Dim notesCol As Range
    Dim wpCol As Range
    Dim pivotTable As PivotTable
    Dim firstPivotCol As Range
    Dim notesColNum As Long
    Dim wpColNum As Long
    Dim firstPivotColNum As Long
    Dim notesReportCol As Range
    Dim notesReportColNum As Long

    ' Set worksheet references
    Set wsWarFeb = ThisWorkbook.Sheets("War Feb")
    Set wsReport = ThisWorkbook.Sheets("Report")

    ' Find the last row in column A of "War Feb"
    lastRow = wsWarFeb.Cells(wsWarFeb.Rows.Count, "A").End(xlUp).Row

    ' Find the column with the name "Notes from PEs" in "War Feb" (Row 1)
    Set notesCol = wsWarFeb.Rows(1).Find("Notes from PEs", LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Find the column with the name "WP" in "War Feb" (Row 1)
    Set wpCol = wsWarFeb.Rows(1).Find("WP", LookIn:=xlValues, LookAt:=xlWhole)

    ' Find the first column of the pivot table in the "Report" sheet
    For Each pivotTable In wsReport.PivotTables
        Set firstPivotCol = pivotTable.TableRange1.Columns(1)
        firstPivotColNum = firstPivotCol.Column
        Exit For
    Next pivotTable

    ' Find the column with the name "Notes" in "Report" (Row 4)
    Set notesReportCol = wsReport.Rows(4).Find("Notes", LookIn:=xlValues, LookAt:=xlWhole)

    ' Ensure all necessary columns are found
    If Not notesCol Is Nothing And Not wpCol Is Nothing And firstPivotColNum > 0 And Not notesReportCol Is Nothing Then
        notesColNum = notesCol.Column
        wpColNum = wpCol.Column
        notesReportColNum = notesReportCol.Column

        ' Define the XLOOKUP formula dynamically using R1C1 notation
        formula = "=XLOOKUP(RC" & wpColNum & ", 'Report'!R4C" & firstPivotColNum & ":R1000C" & firstPivotColNum & ", 'Report'!R4C" & notesReportColNum & ":R1000C" & notesReportColNum & ")"

        ' Debugging: Print the formula to the Immediate Window
        Debug.Print formula

        ' Insert the formula into the "Notes from PEs" column in War Feb sheet
        wsWarFeb.Range(wsWarFeb.Cells(2, notesColNum), wsWarFeb.Cells(lastRow, notesColNum)).FormulaR1C1 = formula
    Else
        MsgBox "Column 'Notes from PEs', 'WP', 'Notes' in Report (Row 4), or the first column of the pivot table not found.", vbExclamation
    End If

End Sub
```