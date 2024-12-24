```vb
Sub JiraRaw_Test()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim lastRow3 As Long
    Dim lastRow4 As Long

    ' Set the active worksheet to a variable for better readability
    Set ws = ActiveSheet

    With ws
        ' Handle potential errors when deleting shapes
        On Error Resume Next
        .Shapes.Range(Array("Picture 1")).Delete
        On Error GoTo 0

        ' Clear all borders
        .Cells.Borders.LineStyle = xlLineStyleNone

        ' Delete the first three rows
        .Rows("1:3").Delete

        ' Format the first row
        .Rows(1).Interior.Color = RGB(64, 64, 64)
        .Rows(1).Font.Color = RGB(255, 255, 255)

        ' Find the last rows for various columns
        lastRow = .Cells(.Rows.Count, "K").End(xlUp).Row
        lastRow1 = .Cells(.Rows.Count, "L").End(xlUp).Row
        lastRow2 = .Cells(.Rows.Count, "Q").End(xlUp).Row
        lastRow3 = .Cells(.Rows.Count, "B").End(xlUp).Row
        lastRow4 = .Cells(.Rows.Count, "J").End(xlUp).Row

        ' Insert new columns and add headers
        .Columns("G:H").Insert Shift:=xlToRight
        .Cells(1, "G").Value = "WP short"
        .Cells(1, "H").Value = "Type of work"

        .Columns("N:O").Insert Shift:=xlToRight
        .Cells(1, "N").Value = "Original Estimate, H"
        .Cells(1, "O").Value = "Remaining Estimate, H"

        .Columns("R").Insert Shift:=xlToRight
        .Cells(1, "R").Value = "Sprint#"

        .Columns("T").Insert Shift:=xlToRight
        .Cells(1, "T").Value = "Status2"

        .Columns("U").Insert Shift:=xlToRight
        .Cells(1, "U").Value = "Project"

        .Columns("V").Insert Shift:=xlToRight
        .Cells(1, "V").Value = "EV"

        ' Apply formulas to the respective columns
        If lastRow >= 2 Then .Range("M2:M" & lastRow).Formula = "=K2/60/60"
        If lastRow1 >= 2 Then .Range("N2:N" & lastRow1).Formula = "=L2/60/60"
        If lastRow2 >= 2 Then .Range("R2:R" & lastRow2).Formula = "=MID(Q2, FIND(""SP"", Q2), 4)"
        If lastRow3 >= 2 Then .Range("U2:U" & lastRow3).Formula = "=LEFT(B2,3)"
        If lastRow4 >= 2 Then .Range("V2:V" & lastRow4).Formula = "=J2*4"

        ' Autofit the columns and rows
        .Columns("A:Z").AutoFit ' Adjust column range if necessary
        .Rows.AutoFit

    End With

End Sub
```