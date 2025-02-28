```vb
Sub AdjustAndCombineSheets()

    Dim ws As Worksheet
    Dim combinedWs As Worksheet
    Dim rng As Range
    Dim deleteRange As Range
    Dim sprintCol As Range
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

    ' Define the columns to delete
    columnsToDelete = Array("Issue Links", "Fix Version/s", "ROI($)", "Updated", "Sprint History", "Sprint commitment", _
                            "Project Status (Date / Comments)", "Last Issue Comment", "Description", "Solution")

    ' Create a new sheet for the combined data
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

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Instructions" And ws.Name <> "CombinedData" Then
            ws.Columns.AutoFit

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
                Set rng = ws.UsedRange.Find(What:="Generated at", LookIn:=xlValues, LookAt:=xlPart, _
                                            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
                If Not rng Is Nothing Then
                    Set deleteRange = rng
                    Do
                        If deleteRange Is Nothing Then
                            Set deleteRange = rng
                        Else
                            Set deleteRange = Union(deleteRange, rng)
                        End If
                        Set rng = ws.UsedRange.FindNext(rng)
                    Loop While Not rng Is Nothing And rng.Address <> deleteRange.Cells(1, 1).Address
                    deleteRange.EntireRow.Delete
                End If
            End With

            ' Find the column with the header "Sprint"
            Set sprintCol = ws.Rows(1).Find(What:="Sprint", LookIn:=xlValues, LookAt:=xlWhole)

            If Not sprintCol Is Nothing Then
                sprintColLetter = Split(sprintCol.Address, "$")(1)

                ' Insert a new column after the Sprint column
                sprintCol.Offset(0, 1).EntireColumn.Insert Shift:=xlToRight
                newColLetter = Split(sprintCol.Offset(0, 1).Address, "$")(1)
                ws.Cells(1, sprintCol.Column + 1).Value = "SP#"

                ' Find the last row in the Sprint column
                lastRow = ws.Cells(ws.Rows.Count, sprintCol.Column).End(xlUp).Row

                ' Apply the formula to extract sprint number
                For Each cell In ws.Range(newColLetter & "2:" & newColLetter & lastRow)
                    cell.Formula = "=MID(" & sprintColLetter & cell.Row & ", FIND(""_"", " & sprintColLetter & cell.Row & ") + 1, " & _
                                   "FIND(""_"", " & sprintColLetter & cell.Row & ", FIND(""_"", " & sprintColLetter & cell.Row & ") + 1) - " & _
                                   "FIND(""_"", " & sprintColLetter & cell.Row & ") - 1)"
                Next cell

                ' Convert formulas to values
                For Each cell In ws.Range(newColLetter & "2:" & newColLetter & lastRow)
                    cell.Value = cell.Value
                Next cell

                ' Rename the sheet based on the value in the "SP#" column
                Dim originalName As String
                Dim counter As Integer

                ' Ensure the value is treated as a string and handle errors
                On Error Resume Next
                spName = Trim(CStr(ws.Cells(2, sprintCol.Column + 1).Value))
                If Err.Number <> 0 Then spName = "Unknown" ' Assign a default name if an error occurs
                On Error GoTo 0

                ' Check if the value is not empty and ensure unique sheet name
                If spName <> "" Then
                    originalName = spName
                    counter = 1

                    ' Check if a sheet with the same name already exists
                    Do While SheetExists(spName)
                        spName = originalName & "_" & counter
                        counter = counter + 1
                    Loop

                    ' Rename the worksheet
                    ws.Name = spName
                Else
                    ws.Name = "Not Found"
                End If
            End If

            ' Create a table
            Set rng = ws.UsedRange
            Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
            tbl.Name = "JiraData_Table"
            tbl.TableStyle = "TableStyleLight8"

            ' Apply formatting to the first row
            With ws.Rows(1)
                .Interior.Color = RGB(0, 0, 0)
                .Font.Color = RGB(255, 255, 255)
                .Font.Size = 11
                .Font.Name = "Arial"
            End With

            ' Delete specified columns
            For Each colName In columnsToDelete
                ' Find the column with the specified name
                Set found = ws.Rows(1).Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)
                ' If the column is found, delete it
                If Not found Is Nothing Then
                    ws.Columns(found.Column).Delete
                End If
            Next colName

            ' Copy data to the combined sheet
            ws.UsedRange.Copy Destination:=combinedWs.Cells(nextRow, 1)
            nextRow = combinedWs.Cells(combinedWs.Rows.Count, 1).End(xlUp).Row + 1
        End If
    Next ws

    MsgBox "All sheets adjusted, renamed, and combined except 'Instructions'"

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