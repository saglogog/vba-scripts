Sub threeImgRowCrop()
'
' threeImgRow Macro
'
'
Dim img As InlineShape
Set img = Selection.InlineShapes(1)
With img
.LockAspectRatio = False
.Height = CentimetersToPoints(5.44)
.Width = CentimetersToPoints(4.54)
End With
Exit Sub
End Sub
Sub oneImgRowCrop()
'
' oneImgRowCrop Macro
'
'
Dim img As InlineShape
Set img = Selection.InlineShapes(1)
With img
.LockAspectRatio = False
.Height = CentimetersToPoints(5.44)
.Width = CentimetersToPoints(14.34)
End With
Exit Sub
End Sub
Sub twoImgRowCrop()
'
' twoImgRowCrop Macro
'
'
Dim img As InlineShape
Set img = Selection.InlineShapes(1)
With img
.LockAspectRatio = False
.Height = CentimetersToPoints(5.44)
.Width = CentimetersToPoints(7.02)
End With
Exit Sub
End Sub

Sub fourImgRowCrop()
'
' fourImgRowCrop Macro
'
'
Dim img As InlineShape
Set img = Selection.InlineShapes(1)
With img
.LockAspectRatio = False
.Height = CentimetersToPoints(5.44)
.Width = CentimetersToPoints(3.3)
End With
Exit Sub
End Sub
