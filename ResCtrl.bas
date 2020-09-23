Sub ResizeControl (ctrl As Control, frm As Form, SpaceDown, SpaceRight)

ctrl.Width = frm.ScaleWidth - ctrl.Left - SpaceRight
ctrl.Height = frm.ScaleHeight - ctrl.Top - SpaceDown

End Sub

