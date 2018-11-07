Public Class clock
    Public Property Hours As Integer
    Public Property minutes As Integer
    Public Property second As Integer

    Public Overrides Function ToString() As String
        Return String.Format("{0}:{1}:{2}", Hours.ToString, minutes.ToString, second.ToString)
    End Function
End Class
