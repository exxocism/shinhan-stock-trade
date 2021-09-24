Public Class DataGridViewDoubleBuffer
    Inherits DataGridView
    Public _dgv As DataGridView
    Public Sub New(ByRef dgv As DataGridView)
        _dgv = dgv
    End Sub
    Public Sub EnableDoubleBuffered()
        Dim dgvType As Type = _dgv.[GetType]()
        Dim pi As Reflection.PropertyInfo = dgvType.GetProperty("DoubleBuffered",
                                                      Reflection.BindingFlags.Instance Or
                                                      Reflection.BindingFlags.NonPublic)
        pi.SetValue(_dgv, True, Nothing)
    End Sub
End Class
