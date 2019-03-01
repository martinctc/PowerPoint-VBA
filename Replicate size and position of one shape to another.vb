Sub SimpleShapeSizePositionCopier()
' Copies the property of the first selected chart / shape to the second
' No variables are saved into the memory
' IMPORTANT: YOU NEED TO SELECT THE SHAPES IN ORDER

Dim shp1, shp2 As Shape
Dim i, j, k As Long
Set shp1 = ActiveWindow.Selection.ShapeRange(1)
Set shp2 = ActiveWindow.Selection.ShapeRange(2)
shp2.Height = shp1.Height
shp2.Width = shp1.Width
shp2.Top = shp1.Top
shp2.Left = shp1.Left
End Sub
