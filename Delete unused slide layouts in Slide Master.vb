Sub SlideMasterCleanup()

' This macro cleans up all the unused slide layouts in your Slide Master. 
' Primarily used for reducing file size, especially if slide layouts make use of large HD images.
' Should ideally be used once the contents are near-complete, to ensure that you do not need those additional slide layouts. 

Dim i As Integer
Dim j As Integer
Dim oPres As Presentation
Set oPres = ActivePresentation
On Error Resume Next
With oPres
    For i = 1 To .Designs.Count
        For j = .Designs(i).SlideMaster.CustomLayouts.Count To 1 Step -1
            .Designs(i).SlideMaster.CustomLayouts(j).Delete
        Next
    Next i
End With
End Sub