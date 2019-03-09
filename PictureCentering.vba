Sub PictureCentering()

  Dim slide_width As Single ''Get Slide Width
  Dim slide_high As Single ''Get Slide High
  Dim picture_width As Single ''Get Picture Width
  Dim picture_high As Single ''Get Picture High
  Dim msg As String

  With ActiveWindow.Selection

    ''If not select picture, finish Macro
    If .Type = ppSelectionNone _
    Or .Type = ppSelectionSlides Then
      msg = "Select picture!"
      MsgBox msg
      Exit Sub
    End If

    ''Get Slide Sizes
    slide_width = .SlideRange.Master.Width
    slide_high = .SlideRange.Master.Height

    ''Get Picture Sizes
    picture_width = .ShapeRange.Width
    picture_high = .ShapeRange.Height

    ''Adjust Picture
    .ShapeRange.Left = (slide_width - picture_width) / 2
    .ShapeRange.Top = (slide_high - picture_high) / 2

  End With

End Sub
