Sub TableColourCopier()

If MsgBox("Make sure that you have selected two tables. The first table you select is a DONOR TABLE. The second table receives the colour formatting of the first one. Continue?", vbYesNo) _
= vbNo Then Exit Sub

Dim ncol, nrow, ncells As Long

ncol = ActiveWindow.Selection.ShapeRange(1).Table.Columns.Count
nrow = ActiveWindow.Selection.ShapeRange(1).Table.Rows.Count

'Get total number of cells
ncells = ncol * nrow

Dim c() As Long
'c = ActiveWindow.Selection.ShapeRange(1).Fill.ForeColor.RGB

ReDim c(1 To ncells)

Dim i, j, k As Long

j = k = 1
i = 1

'Loop through every cell in a table
For j = 1 To nrow
    For k = 1 To ncol
        'ActiveWindow.Selection.ShapeRange(1).Table.Cell(j, k).Shape.Fill.ForeColor.RGB = RGB(255, 0, 0)
        c(i) = ActiveWindow.Selection.ShapeRange(1).Table.Cell(j, k).Shape.Fill.ForeColor.RGB
        'ActiveWindow.Selection.ShapeRange(1).Table.Cell(j, k).Shape.TextFrame.TextRange.Text = i
        i = i + 1
    Next
Next

'Check that the colour has been correctly memorised
Dim colorcode As Integer
For colorcode = 1 To ncells
    Debug.Print c(colorcode)
Next

j = k = 1 'reset values to 1
i = 1

Debug.Print i

'Loop through every cell in table 2
For j = 1 To nrow
    For k = 1 To ncol
        'ActiveWindow.Selection.ShapeRange(1).Table.Cell(j, k).Shape.Fill.ForeColor.RGB = RGB(255, 0, 0)
        ActiveWindow.Selection.ShapeRange(2).Table.Cell(j, k).Shape.Fill.ForeColor.RGB = c(i)
        'ActiveWindow.Selection.ShapeRange(2).Table.Cell(j, k).Shape.TextFrame.TextRange.Text = i
        i = i + 1
    Next
Next

End Sub


