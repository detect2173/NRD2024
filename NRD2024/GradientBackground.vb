Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports System.Windows.Forms


Public Class GradientBackground

    Public Shared Sub ApplyGradient(form As Form, hexColor1 As String, hexColor2 As String)
        ' Convert hex strings to Color objects
        Dim color1 As Color = ColorTranslator.FromHtml(hexColor1)
        Dim color2 As Color = ColorTranslator.FromHtml(hexColor2)
        ' Remove any previous handler to prevent multiple handler assignments
        RemoveHandler form.Paint, AddressOf Form_Paint

        ' Add a new Paint event handler to the form
        AddHandler form.Paint, Sub(sender As Object, e As PaintEventArgs)
                                   Using lgb As New LinearGradientBrush(form.ClientRectangle, color1, color2, LinearGradientMode.Vertical)
                                       e.Graphics.FillRectangle(lgb, form.ClientRectangle)
                                   End Using
                               End Sub

        ' Enable double buffering using the extension method
        form.EnableDoubleBuffering(True)

        ' Invalidate the form to trigger a repaint
        form.Invalidate()
    End Sub

    Private Shared Sub Form_Paint(sender As Object, e As PaintEventArgs)
        ' The actual painting occurs in the ApplyGradient method
    End Sub

End Class


