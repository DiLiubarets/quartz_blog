```vb
Sub CombineSheets_withCol()

    Dim ws As Worksheet
    Dim combinedWs As Worksheet
    Dim rng As Range
    Dim deleteRange As Range
    Dim sprintCol As Range
    Dim statusCol As Range
    Dim lastRow As Long
    Dim cell As Range
    Dim sprintColLetter As String
    Dim newColLetter As String
    Dim columnsToDelete As Variant
    Dim colName As Variant
    Dim found As Range
    Dim nextRow As Long
    Dim spName As String
    Dim tbl As ListObject
    Dim firstSheet As Boolean
    Dim statusColNum As Integer
    Dim originalName As String
    Dim counter As Integer
    Dim newHeaders As Variant
    Dim lastCol As Integer
    Dim i As Integer
    Dim origEstCol As Range, origEstHrsCol As Range
    Dim remEstCol As Range, remHrsCol As Range
    Dim timeSpentCol As Range, timeSpentHrsCol As Range
    Dim origEstHrsColNum As Integer, remHrsColNum As Integer, timeSpentHrsColNum As Integer

    ' Define the columns to delete
    columnsToDelete = Array("Issue Links", "Fix Version/s", "ROI($)", "Updated", "Sprint History", "Sprint commitment", _
                            "Project Status (Date / Comments)", "Last Issue Comment", "Description", "Solution")

    ' Define new columns to add in CombinedData
    newHeaders = Array("Original Estima hrs", "Remaining hrs", "Time Spent hrs", "Original Estima SP", _
                       "Remaining SP", "Time Spent SP", "EV, SP", "WP")

    ' Create or clear the combined sheet
    On Error Resume Next
    Set combinedWs = ThisWorkbook.Sheets("CombinedData")
    On Error GoTo 0

    If combinedWs Is Nothing Then
        Set combinedWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        combinedWs.Name = "CombinedData"
    Else
        combinedWs.Cells.Clear
    End If

    nextRow = 1
    firstSheet = True ' Flag to track the first sheet

    ' Loop through all sheets except "Instructions" and "CombinedData"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Instructions" And ws.Name <> "CombinedData" Then
            With ws
                On Error Resume Next
                .Shapes.Range(Array("Picture 1")).Delete
                On Error GoTo 0
                .Cells.UnMerge
                .Cells.Borders.LineStyle = xlNone ' Clear all borders
                .Rows("1:3").Delete ' Delete the first three rows
                
                ' Format the first row
                .Rows(1).Interior.Color = RGB(64, 64, 64)
                .Rows(1).Font.Color = RGB(255, 255, 255)

                ' Find and delete rows containing "Generated at"
                Set rng = ws.UsedRange.Find(What:="Generated at", LookIn:=xlValues, LookAt:=xlPart)
                If Not rng Is Nothing Then
                    Do
                        rng.EntireRow.Delete
                        Set rng = ws.UsedRange.FindNext(rng)
                    Loop While Not rng Is Nothing
                End If
            End With

            ' Delete specified columns
            For Each colName In columnsToDelete
                Set found = ws.Rows(1).Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)
                If Not found Is Nothing Then ws.Columns(found.Column).Delete
            Next colName

            ' AutoFit columns and rows
            ws.Cells.EntireColumn.AutoFit
            ws.Cells.EntireRow.AutoFit

            ' Copy data to CombinedData sheet
            If firstSheet Then
                ws.UsedRange.Copy Destination:=combinedWs.Cells(nextRow, 1)
                firstSheet = False
            Else
                ws.UsedRange.Offset(1, 0).Resize(ws.UsedRange.Rows.Count - 1, ws.UsedRange.Columns.Count).Copy _
                    Destination:=combinedWs.Cells(nextRow, 1)
            End If
            
            nextRow = combinedWs.Cells(combinedWs.Rows.Count, 1).End(xlUp).Row + 1
        End If
    Next ws

    ' Add new columns in CombinedData
    lastCol = combinedWs.Cells(1, combinedWs.Columns.Count).End(xlToLeft).Column
    For i = 0 To UBound(newHeaders)
        combinedWs.Cells(1, lastCol + i + 1).Value = newHeaders(i)
    Next i

    ' Find relevant columns
    Set origEstCol = combinedWs.Rows(1).Find(What:="Original Estimate", LookIn:=xlValues, LookAt:=xlWhole)
    Set origEstHrsCol = combinedWs.Rows(1).Find(What:="Original Estima hrs", LookIn:=xlValues, LookAt:=xlWhole)
    Set remEstCol = combinedWs.Rows(1).Find(What:="Remaining Estimate", LookIn:=xlValues, LookAt:=xlWhole)
    Set remHrsCol = combinedWs.Rows(1).Find(What:="Remaining hrs", LookIn:=xlValues, LookAt:=xlWhole)
    Set timeSpentCol = combinedWs.Rows(1).Find(What:="Time Spent", LookIn:=xlValues, LookAt:=xlWhole)
    Set timeSpentHrsCol = combinedWs.Rows(1).Find(What:="Time Spent hrs", LookIn:=xlValues, LookAt:=xlWhole)

    ' Apply formulas if columns are found
    lastRow = combinedWs.Cells(combinedWs.Rows.Count, 1).End(xlUp).Row

    If Not origEstCol Is Nothing And Not origEstHrsCol Is Nothing Then
        origEstHrsColNum = origEstHrsCol.Column
        For i = 2 To lastRow
            combinedWs.Cells(i, origEstHrsColNum).Formula = "=" & combinedWs.Cells(i, origEstCol.Column).Address(False, False) & "/3600"
        Next i
    End If

    If Not remEstCol Is Nothing And Not remHrsCol Is Nothing Then
        remHrsColNum = remHrsCol.Column
        For i = 2 To lastRow
            combinedWs.Cells(i, remHrsColNum).Formula = "=" & combinedWs.Cells(i, remEstCol.Column).Address(False, False) & "/3600"
        Next i
    End If

    If Not timeSpentCol Is Nothing And Not timeSpentHrsCol Is Nothing Then
        timeSpentHrsColNum = timeSpentHrsCol.Column
        For i = 2 To lastRow
            combinedWs.Cells(i, timeSpentHrsColNum).Formula = "=" & combinedWs.Cells(i, timeSpentCol.Column).Address(False, False) & "/3600"
        Next i
    End If

    ' Convert formulas to values
    For i = 2 To lastRow
        If Not origEstHrsCol Is Nothing Then combinedWs.Cells(i, origEstHrsColNum).Value = combinedWs.Cells(i, origEstHrsColNum).Value
        If Not remHrsCol Is Nothing Then combinedWs.Cells(i, remHrsColNum).Value = combinedWs.Cells(i, remHrsColNum).Value
        If Not timeSpentHrsCol Is Nothing Then combinedWs.Cells(i, timeSpentHrsColNum).Value = combinedWs.Cells(i, timeSpentHrsColNum).Value
    Next i

    ' AutoFit all columns and rows in the combined sheet
    combinedWs.Cells.EntireColumn.AutoFit
    combinedWs.Cells.EntireRow.AutoFit

    MsgBox "All sheets adjusted, renamed, cleaned, and combined without 'Open' rows. Formulas applied and converted to values!"

End Sub

' Function to check if a sheet with a given name exists
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
```