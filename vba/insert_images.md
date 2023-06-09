# Insert Images

```vba
Sub 매크로_사진넣기()
    Dim img As Variant
    img = Application.GetOpenFilename _
                    (filefilter:="Picture Files,*.jpg;*.bmp;*.tif;*.gif;*.png")
    If img = False Then
        Exit Sub
    End If
  
    With ActiveSheet.Pictures.Insert(img).ShapeRange
        .LockAspectRatio = msoFalse
        .Height = Selection.Height '선택한 영역의 높이
        .Width = Selection.Width
        .Left = Selection.Left
        .Top = Selection.Top
    End With
End Sub
```
