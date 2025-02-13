```vb
Sub Test_War_PivotTable()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim pivotRange As Range
    Dim pivotDestination As Range
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets("WAR_Report_Data")
    Set pivotRange = wsData.Range("A1").CurrentRegion

    ' Delete existing "Report" sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Report").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create new "Report" sheet
    Set wsPivot = wb.Sheets.Add
    wsPivot.Name = "Report"

    ' Set pivot table destination
    Set pivotDestination = wsPivot.Range("B6")

    ' Create Pivot Cache
    Set pivotCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)

    ' Create Pivot Table
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="WAR_PivotTable")

    ' Configure Pivot Table
    With pivotTable
        .PivotFields("Components").Orientation = xlRowField

        ' Add Data Fields
        On Error Resume Next
        .PivotFields("AC, Week 1&2").Orientation = xlDataField
        .PivotFields("AC, Week 1&2").Function = xlSum
        .PivotFields("AC, Week 1&2").Caption = "AC, Week 1&2"

        .PivotFields("AC, Week 3&4").Orientation = xlDataField
        .PivotFields("AC, Week 3&4").Function = xlSum
        .PivotFields("AC, Week 3&4").Caption = "AC, Week 3&4"

        .PivotFields("ETC in hrs").Orientation = xlDataField
        .PivotFields("ETC in hrs").Function = xlMax
        .PivotFields("ETC in hrs").Caption = "ETC in hrs"

        .PivotFields("EV in hrs").Orientation = xlDataField
        .PivotFields("EV in hrs").Function = xlMax
        .PivotFields("EV in hrs").Caption = "EV in hrs"

        .PivotFields("Closed Tickets").Orientation = xlDataField
        .PivotFields("Closed Tickets").Function = xlMax
        .PivotFields("Closed Tickets").Caption = "Closed Tickets"

        .PivotFields("EV Closed").Orientation = xlDataField
        .PivotFields("EV Closed").Function = xlMax
        .PivotFields("EV Closed").Caption = "EV Closed"

        .PivotFields("AC Sum Week 1&2").Orientation = xlDataField
        .PivotFields("AC Sum Week 1&2").Function = xlMax
        .PivotFields("AC Sum Week 1&2").Caption = "AC Sum Week 1&2"

        .PivotFields("AC Sum Week 3&4").Orientation = xlDataField
        .PivotFields("AC Sum Week 3&4").Function = xlMax
        .PivotFields("AC Sum Week 3&4").Caption = "AC Sum Week 3&4"
        On Error GoTo 0

        ' Refresh Pivot Table before adding calculated fields
        .PivotCache.Refresh

        ' Add Calculated Fields
        On Error Resume Next
        .CalculatedFields.Add Name:="EV,% Week 1&2", Formula:="=IF([EV Closed]=0, 0, [EV Closed]/[ETC in hrs])"
        .CalculatedFields.Add Name:="EV,% Week 3&4", Formula:="=IF([EV in hrs]=0, 0, [EV in hrs]/[ETC in hrs])"
        .CalculatedFields.Add Name:="AC/ETC, Week 1&2", Formula:="=IFERROR(IF([AC Sum Week 1&2]=0, 0, [AC Sum Week 1&2]/[ETC in hrs]), 0)"
        .CalculatedFields.Add Name:="AC/ETC, Week 3&4", Formula:="=IFERROR(IF([AC Sum Week 3&4]=0, 0, [AC Sum Week 3&4]/[ETC in hrs]), 0)"
        On Error GoTo 0

        ' Refresh Pivot Table again
        .PivotCache.Refresh

        ' Check if the calculated field exists before configuring it
        Dim pf As PivotField
        On Error Resume Next
        Set pf = .PivotFields("EV,% Week 1&2")
        On Error GoTo 0

        If Not pf Is Nothing Then
            With pf
                .Orientation = xlDataField
                .NumberFormat = "0.00%"
                .Function = xlSum
                .Caption = "EV,% Week 1&2"
            End With
        Else
            MsgBox "Calculated field 'EV,% Week 1&2' was not created successfully.", vbExclamation, "Error"
        End If

        ' Repeat for other calculated fields
        On Error Resume Next
        Set pf = .PivotFields("EV,% Week 3&4")
        On Error GoTo 0

        If Not pf Is Nothing Then
            With pf
                .Orientation = xlDataField
                .NumberFormat = "0.00%"
                .Function = xlSum
                .Caption = "EV,% Week 3&4"
            End With
        Else
            MsgBox "Calculated field 'EV,% Week 3&4' was not created successfully.", vbExclamation, "Error"
        End If

        ' Apply Pivot Table Style
        .TableStyle2 = "PivotStyleDark8"

        ' Remove Grand Totals
        .ColumnGrand = False
        .RowGrand = False

        ' Set Pivot Table to Tabular Form and Remove Subtotals
        .RowAxisLayout xlTabularRow
        Dim pfRow As PivotField
        For Each pfRow In .RowFields
            pfRow.Subtotals(1) = False
        Next pfRow
    End With

    ' Autofit Columns
    wsPivot.Columns.AutoFit
End Sub
```