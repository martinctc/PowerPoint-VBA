Sub ChangeAxisTitle()

' This applies to ALL the chart objects for ALL slides in a PowerPoint file
' If you wish to apply this only to a select number of slides, copy them out first to separate PPT file

' Sets Axis Title text for both Vertical and Horizontal Axes
' A "Importance-Performance" grid is used as an example.

Dim sld As Slide
Dim shp As Shape
Dim sr As Series
Dim chrt As Chart
Dim i, j, k, m As Long

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasChart Then
                shp.Chart.Axes(xlValue).AxisTitle.Text = "Importance" 'Y-axis
                shp.Chart.Axes(xlCategory).AxisTitle.Text = "Performance" 'X-axis
            End If
    Next shp
    Next sld

End Sub

