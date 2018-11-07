Imports hebrew_year = System.Int32
Imports hebrew_month = System.Int32
Imports hebrew_day = System.Int32

Public Class hebrew_date

    Public Property year As hebrew_year
    Private Property _month As hebrew_month
    Public Property month As hebrew_month
        Get
            Return _month
        End Get
        Set(value As hebrew_month)
            _month = value
        End Set
    End Property
    Public Property day As hebrew_year

    Sub New()
        MyClass.New(0, 0, 0)
    End Sub

    Sub New(h_year As hebrew_year, h_month As hebrew_month, h_day As hebrew_day)
        year = h_year
        month = h_month
        day = h_day
    End Sub

    Public Overrides Function ToString() As String
        Return String.Format("{0} {1} {2}", year, month, day)
    End Function

End Class
