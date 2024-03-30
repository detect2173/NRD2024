Imports System.Windows.Forms

Public Class TransparentTabControl
    Inherits TabControl

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        ' Intentionally blank to suppress background painting
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        ' Perform the base OnPaint to ensure tab headers and borders are still rendered
        MyBase.OnPaint(e)

        ' Custom painting can be done here if needed
    End Sub
End Class
