Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

Module FormExtensions

    <Extension()>
    Public Sub EnableDoubleBuffering(ByVal form As Form, ByVal enable As Boolean)
        Dim doubleBufferPropertyInfo As PropertyInfo = GetType(Control).GetProperty("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance)
        doubleBufferPropertyInfo.SetValue(form, enable, Nothing)
    End Sub

End Module
