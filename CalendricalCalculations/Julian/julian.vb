Imports julian_year = System.Int32
Imports julian_month = System.Int32
Imports julian_day = System.Int32
Imports CC = CalendricalCalculations.Basic
Imports fixed_date = System.Int32

Public Class julian

    Public Const january = 1
    Public Const february = 2
    Public Const march = 3
    Public Const april = 4
    Public Const may = 5
    Public Const june = 6
    Public Const july = 7
    Public Const august = 8
    Public Const september = 9
    Public Const october = 10
    Public Const november = 11
    Public Const december = 12

    Public Function julian_leap_year(j_year As julian_year) As Boolean              ' 3.1
        Return CC.mod(j_year, 4) = If(j_year > 0, 0, 3)
    End Function

    Public Function julian_epoch() As fixed_date                                    ' 3.2
        Return Basic.rd((New gregorian()).fixed_from_gregorian(New gregorian_date(0, gregorian.december, 30)))
    End Function

    Public Function fixed_from_julian(j_date As julian_date) As fixed_date          ' 3.3
        Dim y As Integer = If(j_date.year < 0, j_date.year + 1, j_date.year)
        Dim result As Integer = julian_epoch()
        result -= 1
        result += 365 * Math.Floor(y - 1)
        result += Math.Floor((y - 1) / 4)
        result += Math.Floor((1 / 12) * (367 * j_date.month - 362))
        If j_date.month <= 2 Then
            result -= 0
        ElseIf julian_leap_year(j_date.year) Then
            result -= 1
        Else
            result -= 2
        End If
        result += j_date.day
        Return result
    End Function

    Public Function julian_from_fixed([date] As fixed_date) As julian_date          ' 3.4
        Dim approx As Integer = Math.Floor((1 / 1461) * (4 * ([date] - julian_epoch()) + 1464))
        Dim year As Integer = If(approx <= 0, approx - 1, approx)
        Dim prior_days As Integer = [date] - fixed_from_julian(New julian_date(year, january, 1))
        Dim correction As Integer = 0
        If [date] < fixed_from_julian(New julian_date(year, march, 1)) Then
            correction = 0
        ElseIf julian_leap_year(year) Then
            correction = 1
        Else
            correction = 2
        End If
        Dim month As Integer = Math.Floor((1 / 367) * (12 * (prior_days + correction) + 373))
        Dim day As Integer = [date] - fixed_from_julian(New julian_date(year, month, 1)) + 1
        Return New julian_date(year, month, day)
    End Function

End Class
