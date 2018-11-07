Imports julian_year = System.Int32
Imports julian_month = System.Int32
Imports julian_day = System.Int32

Public Class julian_date

    Public Property year As julian_year
    Public Property month As julian_month
    Public Property day As julian_year

    Sub New()
        MyClass.New(0, 0, 0)
    End Sub

    Sub New(j_year As julian_year, j_month As julian_month, j_day As julian_day)
        year = j_year
        month = j_month
        day = j_day
    End Sub

    Public Overrides Function ToString() As String
        Return String.Format("{0} {1} {2}", year, month, day)
    End Function

End Class
