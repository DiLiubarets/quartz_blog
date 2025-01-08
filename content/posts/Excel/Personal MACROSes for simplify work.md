```vb
Sub JiraRaw_Cleaning()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim deleteRange As Range

    ' Set the active worksheet to a variable for better readability
    Set ws = ActiveSheet
    ws.Columns.AutoFit

    With ws
        On Error Resume Next
        .Shapes.Range(Array("Picture 1")).Delete
        On Error GoTo 0
        .Cells.UnMerge

        ' Clear all borders
        .Cells.Borders.LineStyle = xlLineStyleNone

        ' Delete the first three rows
        .Rows("1:3").Delete

        ' Format the first row
        .Rows(1).Interior.Color = RGB(64, 64, 64)
        .Rows(1).Font.Color = RGB(255, 255, 255)

        ' Find and delete rows containing "Generated at"
        Set rng = ws.UsedRange.Find(What:="Generated at", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
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

            ' Delete the rows
            deleteRange.EntireRow.Delete
        End If

        ' AutoFit columns and rows
        .Cells(1, 1).Select
        .UsedRange.Columns.AutoFit
        .UsedRange.Rows.AutoFit
    End With
End Sub

```