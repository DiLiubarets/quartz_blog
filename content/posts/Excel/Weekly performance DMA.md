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
        With .PivotFields("AC, Week 1&2")
            .Orientation = xlDataField
            .Function = xlSum
            .Caption = "AC, Week 1&2"
        End With

        With .PivotFields("AC, Week 3&4")
            .Orientation = xlDataField
            .Function = xlSum
            .Caption = "AC, Week 3&4"
        End With

        With .PivotFields("ETC in hrs")
            .Orientation = xlDataField
            .Function = xlMax
            .Caption = "ETC in hrs"
        End With

        With .PivotFields("EV in hrs")
            .Orientation = xlDataField
            .Function = xlMax
            .Caption = "EV in hrs"
        End With

        With .PivotFields("Closed Tickets")
            .Orientation = xlDataField
            .Function = xlMax
            .Caption = "Closed Tickets"
        End With

        With .PivotFields("EV Closed")
            .Orientation = xlDataField
            .Function = xlMax
            .Caption = "EV Closed"
        End With

        With .PivotFields("AC Sum Week 1&2")
            .Orientation = xlDataField
            .Function = xlMax
            .Caption = "AC Sum Week 1&2"
        End With

        With .PivotFields("AC Sum Week 3&4")
            .Orientation = xlDataField
            .Function = xlMax
            .Caption = "AC Sum Week 3&4"
        End With

        ' Add Calculated Fields
        On Error Resume Next ' Prevent errors if the calculated field already exists
        .CalculatedFields.Add Name:="EV,% Week", Formula:="=IF([EV Closed]=0, 0, [EV Closed]/[ETC in hrs])"
        .CalculatedFields.Add Name:="EV,%", Formula:="=IF([EV in hrs]=0, 0, [EV in hrs]/[ETC in hrs])"
        .CalculatedFields.Add Name:="AC/ETC, Week 1&2", Formula:="=IFERROR(IF([AC Sum Week 1&2]=0, 0, [AC Sum Week 1&2]/[ETC in hrs]), 0)"
        .CalculatedFields.Add Name:="AC/ETC, Week 3&4", Formula:="=IFERROR(IF([AC Sum Week 3&4]=0, 0, [AC Sum Week 3&4]/[ETC in hrs]), 0)"
        On Error GoTo 0

        ' Configure Calculated Fields
        With .PivotFields("EV,% Week")
            .Orientation = xlDataField
            .NumberFormat = "0.00%"
            .Function = xlSum
            .Caption = "EV,% Week 1&2"
        End With

        With .PivotFields("EV,%")
            .Orientation = xlDataField
            .NumberFormat = "0.00%"
            .Function = xlSum
            .Caption = "EV,% Week 3&4"
        End With

        With .PivotFields("AC/ETC, Week 1&2")
            .Orientation = xlDataField
            .NumberFormat = "0.00%"
            .Function = xlSum
            .Caption = "AC/ETC, Week 1&2"
        End With

        With .PivotFields("AC/ETC, Week 3&4")
            .Orientation = xlDataField
            .NumberFormat = "0.00%"
            .Function = xlSum
            .Caption = "AC/ETC, Week 3&4"
        End With

        ' Apply Pivot Table Style
        .TableStyle2 = "PivotStyleDark8"

        ' Remove Grand Totals
        .ColumnGrand = False
        .RowGrand = False

        ' Set Pivot Table to Tabular Form and Remove Subtotals
        .RowAxisLayout xlTabularRow
        Dim pf As PivotField
        For Each pf In .RowFields
            pf.Subtotals(1) = False
        Next pf

        ' Apply Filter to Show Only "Hardware", "Software", and "System"
        With .PivotFields("Components")
            .ClearAllFilters
            Dim pi As PivotItem
            On Error Resume Next ' Prevent errors if an item does not exist
            .PivotItems("Hardware").Visible = True
            .PivotItems("Software").Visible = True
            .PivotItems("System").Visible = True
            On Error GoTo 0

            ' Hide Other Items
            For Each pi In .PivotItems
                If pi.Name <> "Hardware" And pi.Name <> "Software" And pi.Name <> "System" Then
                    On Error Resume Next
                    pi.Visible = False
                    On Error GoTo 0
                End If
            Next pi
        End With
    End With

    ' Autofit Columns
    wsPivot.Columns.AutoFit
End Sub
```