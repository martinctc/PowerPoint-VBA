Sub PPT_Reverse_Axis()

Dim sld As Slide
Dim shp As Shape
Dim sr As Series
Dim chrt As Chart
Dim i, j, k, m As Long


Set sld = Application.ActiveWindow.View.Slide ' Only applies to chart objects on the active slide
        For Each shp In sld.Shapes
            If shp.HasChart Then
      			With shp.Chart.Axes(xlCategory) 'X-axis - change to xlValue as required
        			.ReversePlotOrder = False 'Flip between True/False as required
                End With
            End If
    Next shp
End Sub