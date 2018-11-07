Imports gregorian_year = System.Int32
Imports gregorian_month = System.Int32
Imports gregorian_day = System.Int32

Public Class gregorian_date

    Public Property year As gregorian_year
    Public Property month As gregorian_month
    Public Property day As gregorian_year

    Sub New()
        MyClass.New(0, 0, 0)
    End Sub

    Sub New(g_year As gregorian_year, g_month As gregorian_month, g_day As gregorian_day)
        year = g_year
        month = g_month
        day = g_day
    End Sub

    Public Overrides Function ToString() As String
        Return String.Format("{0:0000}-{1:00}-{2}", year, month, day)
    End Function

End Class
