' These are methods that crop images, when many images are supposed to be presented together in word document so that their sizes are uniform throughout the document. 
' The code is run in word (each method represents a macro).
' Each method is supposed to be attached to a shortcut button on word. 
' Then the image is selected and the corresponding button is pressed, depending on the # of imgs () that you want to present side to side. Apply for each image separately.
' The sizeof the images is hardcoded, but you can change it inside each method.


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
