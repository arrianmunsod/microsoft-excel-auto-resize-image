Attribute VB_Name = "Module1"
Sub FitImageWithMargin()
    Dim shp As Shape
    Dim targetRange As Range
    Dim margin As Double
    Dim availWidth As Double, availHeight As Double
    
    ' SET YOUR MARGIN HERE (in pixels)
    margin = 2
    
    If TypeName(Selection) = "DrawingObjects" Or TypeName(Selection) = "Picture" Then
        
        For Each shp In Selection.ShapeRange
            ' 1. Identify the target (Merged or Single Cell)
            Set targetRange = shp.TopLeftCell
            If targetRange.MergeCells Then Set targetRange = targetRange.MergeArea
            
            ' 2. Calculate the "Inner Box" (Cell size minus margins)
            availWidth = targetRange.Width - (margin * 2)
            availHeight = targetRange.Height - (margin * 2)
            
            With shp
                .LockAspectRatio = msoTrue
                
                ' 3. Fit to the "Inner Box"
                If (.Width / .Height) > (availWidth / availHeight) Then
                    .Width = availWidth
                Else
                    .Height = availHeight
                End If
                
                ' 4. Center within the cell (including the margin offset)
                .Top = targetRange.Top + (targetRange.Height - .Height) / 2
                .Left = targetRange.Left + (targetRange.Width - .Width) / 2
                
                .Placement = xlMoveAndSize
            End With
        Next shp
        
    Else
        MsgBox "Please select the image(s) first!", vbExclamation
    End If
End Sub

