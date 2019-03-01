Sub TableColorDetector()

' This macro takes a table in PowerPoint and prints the corresponding colour code to the Debug Window.
' The loop works by column first, from left to right. 

Dim ncol, nrow, ncells As Long

ncol = ActiveWindow.Selection.ShapeRange.Table.Columns.Count
nrow = ActiveWindow.Selection.ShapeRange.Table.Rows.Count

'Get total number of cells
ncells = ncol * nrow

Dim c() As Long
'c = ActiveWindow.Selection.ShapeRange.Fill.ForeColor.RGB

ReDim c(1 To ncells)

Dim i, j, k As Long

j = k = 1
i = 1

'Loop through every cell in a table
For j = 1 To nrow
    For k = 1 To ncol
        'ActiveWindow.Selection.ShapeRange.Table.Cell(j, k).Shape.Fill.ForeColor.RGB = RGB(255, 0, 0)
        c(i) = ActiveWindow.Selection.ShapeRange.Table.Cell(j, k).Shape.Fill.ForeColor.RGB
        ActiveWindow.Selection.ShapeRange.Table.Cell(j, k).Shape.TextFrame.TextRange.Text = i
        i = i + 1
    Next
Next

Dim colorcode As Integer
For colorcode = 1 To ncells
    Debug.Print c(colorcode)
Next

End Sub

