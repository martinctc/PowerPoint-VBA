Sub TablePositionandSizeTinkerer()

'Alt+F11 to start up code module
'Insert New Module >>
'Select Table >> Run

'r converts cm to points
'this is a constant - DO NOT CHANGE
r = 28.3464567

'Make general specifications on table position and size

With ActiveWindow.Selection.ShapeRange
    .Left = 5.72 * r 'change the number for desired x position
    .Top = 4.75 * r 'change the number for desired y position
    .Height = 11.4 * r
    .Width = 22.42 * r
    .Table.Columns(1).Width = 2.77 * r
    .Table.Columns(2).Width = 2 * r
    .Table.Columns(3).Width = 2 * r
    .Table.Columns(4).Width = 3.8 * r
    .Table.Columns(5).Width = 3.8 * r
    .Table.Columns(6).Width = 3.8 * r
    .Table.Columns(7).Width = 3.8 * r
    .Table.Rows(1).Height = 1.59 * r
    .Table.Rows(2).Height = 0.99 * r
    .Table.Rows(3).Height = 1.61 * r
    .Table.Rows(4).Height = 1.61 * r
    .Table.Rows(5).Height = 1.61 * r
    .Table.Rows(6).Height = 1.61 * r
    .Table.Rows(7).Height = 1.61 * r
    .Height = 11.4 * r
    .Width = 22.42 * r    
  
End With

End Sub

