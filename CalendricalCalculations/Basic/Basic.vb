Imports moment = System.Decimal
Imports julian_day_number = System.Decimal
Imports fixed_date = System.Int32
Imports time = System.Decimal

Public Module Basic

    Public Function rd(t As moment) As moment                                   ' 1.1
        Dim epoch As Integer = 0
        Return t - epoch
    End Function

    Public Function jd_epoch() As moment                                        ' 1.3
        Return rd(-1721424.5)
    End Function

    Public Function moment_from_jd(jd As julian_day_number) As moment           ' 1.4
        Return jd + jd_epoch()
    End Function

    Public Function jd_from_moment(t As moment) As julian_day_number            ' 1.5
        Return t - jd_epoch()
    End Function

    Public Function mjd_epoch() As moment                                       ' 1.6
        Return rd(678576)
    End Function

    Public Function fixed_from_mjd(mjd As julian_day_number) As fixed_date      ' 1.7
        Return mjd + mjd_epoch()
    End Function

    Public Function mjd_from_fixed([date] As fixed_date) As julian_day_number   ' 1.8
        Return [date] - mjd_epoch()
    End Function

    Public Function fixed_from_moment(t As moment) As fixed_date                ' 1.9
        Return Math.Floor(t)
    End Function

    Public Function fixed_from_jd(jd As julian_day_number) As fixed_date        ' 1.10
        Return Math.Floor(moment_from_jd(jd))
    End Function

    Public Function jd_from_fixed([date] As fixed_date) As julian_day_number    ' 1.11
        Return jd_from_moment([date])
    End Function

    Public Function [mod](x As Decimal, y As Decimal) As Integer                ' 1.15
        Return x - y * Math.Floor(x / y)
    End Function

    Public Function time_from_moment(t As moment) As Decimal                    ' 1.16
        Return t Mod 1
    End Function

    Public Function gcd(x As Integer, y As Integer) As Integer                  ' 1.20
        If y = 0 Then Return x
        Return gcd(y, [mod](x, y))
    End Function

    Public Function lcm(x As Integer, y As Integer) As Integer                  ' 1.21
        Return (x * y) / (gcd(x, y))
    End Function

    Public Function amod(x As Decimal, y As Decimal) As Decimal                 ' 1.22
        If y = 0 Then Throw New DivideByZeroException
        Return y + [mod](x, -y)
    End Function

    Public Function time_from_clock(h As Decimal, m As Decimal, s As Decimal) As time       ' 1.34
        Return (1 / 24) * (h + ((m + (s / 60)) / 60))
    End Function

    Public Function clock_from_moment(t As moment) As clock                      ' 1.35
        Dim time As time = time_from_moment(t)
        Dim hour As Integer = Math.Floor(time * 24)
        Dim minute As Integer = Math.Floor((time * 24 * 60) Mod 60)
        Dim second As Integer = (time * 24 * 60 ^ 2) Mod 60
        Return New clock With {.Hours = hour, .minutes = minute, .second = second}
    End Function

    Public Function day_of_week_from_fixed([date] As fixed_date) As day_of_week ' 1.51
        Return [mod]([date] - rd(0) - day_of_week.Sunday, 7)
    End Function

    Public Function kday_on_or_before(k As day_of_week, [date] As fixed_date) As fixed_date ' 1.53
        Return [date] - day_of_week_from_fixed([date] - k)
    End Function

    Public Function kday_on_or_after(k As day_of_week, [date] As fixed_date) As fixed_date  ' 1.58
        Return kday_on_or_after(k, [date] + 6)
    End Function

    Public Function kday_nearest(k As day_of_week, [date] As fixed_date) As fixed_date      ' 1.59
        Return kday_on_or_before(k, [date] + 3)
    End Function

    Public Function kday_before(k As day_of_week, [date] As fixed_date) As fixed_date       ' 1.60
        Return kday_on_or_before(k, [date] - 1)
    End Function

    Public Function kday_after(k As day_of_week, [date] As fixed_date) As fixed_date        ' 1.61
        Return kday_on_or_before(k, [date] + 7)
    End Function

End Module
