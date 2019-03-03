Sub AllChartsResize()

' This applies to ALL the chart objects for ALL slides in a PowerPoint file
' If you wish to apply this only to a select number of slides, copy them out first to separate PPT file
' Change values below to adjust Size and Position.
' Leave "* r" if you are using values that correspond to units shown in PowerPoint
' Optional chunk included below for displaying Axes - commented out.

Dim sld As Slide
Dim shp As Shape
Dim sr As Series
Dim chrt As Chart
Dim i, j, k, m As Long

'r converts cm to points
r = 28.3464567

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasChart Then
                shp.Height = 10.8 * r
                shp.Width = 22.86 * r
                shp.Left = 1.27 * r
                shp.Top = 5.45 * r
			'   shp.Chart.HasAxis(xlValue) = False
            '   shp.Chart.HasAxis(xlCategory) = False
            End If
            
    Next shp
    Next sld
End Sub