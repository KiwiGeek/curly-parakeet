Imports gregorian_year = System.Int32
Imports gregorian_month = System.Int32
Imports gregorian_day = System.Int32
Imports CC = CalendricalCalculations.Basic
Imports fixed_date = System.Int32

Public Class gregorian

    Public Function gregorian_epoch() As fixed_date                                 ' 2.3
        Return Basic.rd(1)
    End Function

    Public Const january = 1                                                        ' 2.4
    Public Const february = 2                                                       ' 2.5
    Public Const march = 3                                                          ' 2.6
    Public Const april = 4                                                          ' 2.7
    Public Const may = 5                                                            ' 2.8
    Public Const june = 6                                                           ' 2.9
    Public Const july = 7                                                           ' 2.10
    Public Const august = 8                                                         ' 2.11
    Public Const september = 9                                                      ' 2.12
    Public Const october = 10                                                       ' 2.13
    Public Const november = 11                                                      ' 2.14
    Public Const december = 12                                                      ' 2.15

    Public Function gregorian_leap_year(g_year As gregorian_year) As Boolean        ' 2.16
        Return CC.mod(g_year, 4) = 0 And Not {100, 200, 300}.Contains(CC.mod(g_year, 400))
    End Function

    Public Function fixed_from_gregorian(g_date As gregorian_date) As fixed_date      ' 2.17
        Dim result As fixed_date = gregorian_epoch()
        result -= 1
        result += 365 * (g_date.year - 1)
        result += Math.Floor((g_date.year - 1) / 4)
        result -= Math.Floor((g_date.year - 1) / 100)
        result += Math.Floor((g_date.year - 1) / 400)
        result += Math.Floor((1 / 12) * (367 * g_date.month - 362))
        If g_date.month <= 2 Then
            result += 0
        ElseIf gregorian_leap_year(g_date.year) Then
            result += -1
        Else
            result += -2
        End If
        result += g_date.day
        Return result
    End Function

    Public Function gregorian_new_year(g_year As gregorian_year) As fixed_date      ' 2.18
        Return fixed_from_gregorian(New gregorian_date(g_year, january, 1))
    End Function

    Public Function gregorian_year_end(g_year As gregorian_year) As fixed_date      ' 2.19
        Return fixed_from_gregorian(New gregorian_date(g_year, december, 31))
    End Function

    Public Function gregorian_year_range(g_year As gregorian_year) As range         ' 2.20
        Return New range(gregorian_new_year(g_year), gregorian_year_end(g_year))
    End Function

    Public Function gregorian_year_from_fixed([date] As fixed_date) As gregorian_year    ' 2.21
        Dim d0 As Decimal = [date] - gregorian_epoch()
        Dim n400 As Decimal = Math.Floor(d0 / 146097)
        Dim d1 As Decimal = CC.mod(d0, 146097)
        Dim n100 As Decimal = Math.Floor(d1 / 36524)
        Dim d2 As Decimal = CC.mod(d1, 36524)
        Dim n4 As Decimal = Math.Floor(d2 / 1461)
        Dim d3 As Decimal = CC.mod(d2, 1461)
        Dim n1 As Decimal = Math.Floor(d3 / 365)
        Dim year As gregorian_year = 400 * n400 + 100 * n100 + 4 * n4 + n1
        If n100 = 4 Or n1 = 4 Then
            Return year
        Else
            Return year + 1
        End If
    End Function

    Public Function gregorian_from_fixed([date] As fixed_date) As gregorian_date        ' 2.23
        Dim year As gregorian_year = gregorian_year_from_fixed([date])
        Dim prior_days As Integer = [date] - gregorian_new_year(year)
        Dim correction As Integer = 0
        If [date] < fixed_from_gregorian(New gregorian_date(year, march, 1)) Then
            correction = 0
        ElseIf gregorian_leap_year(year) Then
            correction = 1
        Else
            correction = 2
        End If
        Dim month = Math.Floor((1 / 367) * (12 * (prior_days + correction) + 373))
        Dim day = 1 + [date] - fixed_from_gregorian(New gregorian_date(year, month, 1))
        Return New gregorian_date(year, month, day)
    End Function

    Public Function gregorian_date_difference(g_date_1 As gregorian_date, g_date_2 As gregorian_date) As Integer    ' 2.24
        Return fixed_from_gregorian(g_date_2) - fixed_from_gregorian(g_date_1)
    End Function

    Public Function day_number(g_date As gregorian_date) As Integer                     ' 2.25
        Return gregorian_date_difference(gregorian_from_fixed(gregorian_year_end(g_date.year - 1)), g_date)
    End Function

    Public Function days_remaining(g_date As gregorian_date) As Integer                 ' 2.26
        Return gregorian_date_difference(g_date, gregorian_from_fixed(gregorian_year_end(g_date.year)))
    End Function

End Class
