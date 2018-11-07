Imports hebrew_year = System.Int32
Imports hebrew_month = System.Int32
Imports hebrew_day = System.Int32
Imports moment = System.Decimal
Imports CC = CalendricalCalculations.Basic
Imports fixed_date = System.Int32
Public Class hebrew

    Public Const nisan = 1                                                                      ' 7.1
    Public Const iyyar = 2                                                                      ' 7.2
    Public Const sivan = 3                                                                      ' 7.3
    Public Const tammuz = 4                                                                     ' 7.4
    Public Const av = 5                                                                         ' 7.5
    Public Const elul = 6                                                                       ' 7.6
    Public Const tishri = 7                                                                     ' 7.7
    Public Const marheshvan = 8                                                                 ' 7.8
    Public Const kislev = 9                                                                     ' 7.9
    Public Const tevet = 10                                                                     ' 7.10
    Public Const shevat = 11                                                                    ' 7.11
    Public Const adar = 12                                                                      ' 7.12
    Public Const adarii = 13                                                                    ' 7.13

    Public Function hebrew_leap_year(h_year As hebrew_year) As Boolean                          ' 7.14
        Return CC.mod((7 * h_year + 1), 19) < 7
    End Function

    Public Function last_month_of_hebrew_year(h_year As hebrew_year) As hebrew_month            ' 7.15
        If hebrew_leap_year(h_year) Then Return adarii
        Return adar
    End Function

    Public Function hebrew_epoch() As fixed_date                                                ' 7.17
        Return (New julian).fixed_from_julian(New julian_date(-3761, julian.october, 7))
    End Function

    Public Function hebrew_calendar_elapsed_days(h_year As hebrew_year) As Integer              ' 7.20
        Dim months_elapsed = Math.Floor((1 / 19) * (235 * h_year - 234))
        Dim parts_elapsed = 12084 + 13753 * months_elapsed
        Dim days = 29 * months_elapsed + Math.Floor(parts_elapsed / 25920)
        If CC.mod(3 * (days + 1), 7) < 3 Then Return days + 1
        Return days
    End Function

    Public Function hebrew_year_length_correction(h_year As hebrew_year) As Integer             ' 7.21
        Dim ny0 = hebrew_calendar_elapsed_days(h_year - 1)
        Dim ny1 = hebrew_calendar_elapsed_days(h_year)
        Dim ny2 = hebrew_calendar_elapsed_days(h_year + 1)
        If ny2 - ny1 = 356 Then Return 2
        If ny1 - ny0 = 382 Then Return 1
        Return 0
    End Function

    Public Function hebrew_new_year(h_year) As fixed_date                                       ' 7.22
        Return hebrew_epoch() + hebrew_calendar_elapsed_days(h_year) + hebrew_year_length_correction(h_year)
    End Function

    Public Function last_day_of_hebrew_month(h_month As hebrew_month, h_year As hebrew_year) As Integer     ' 7.23
        If {iyyar, tammuz, elul, tevet, adarii}.Contains(h_month) Or
            (h_month = adar And Not hebrew_leap_year(h_year)) Or
            (h_month = marheshvan And Not long_marhashvan(h_year)) Or
            (h_month = kislev And short_kislev(h_year)) Then Return 29
        Return 30
    End Function

    Public Function long_marhashvan(h_year As hebrew_year) As Boolean                           ' 7.24
        Return {355, 385}.Contains(days_in_hebrew_year(h_year))
    End Function

    Public Function short_kislev(h_year As hebrew_year) As Boolean                              ' 7.25
        Return {353, 383}.Contains(days_in_hebrew_year(h_year))
    End Function

    Public Function days_in_hebrew_year(h_year As hebrew_year) As Integer                       ' 7.26
        Return hebrew_new_year(h_year + 1) - hebrew_new_year(h_year)
    End Function

    Public Function fixed_from_hebrew(h_date As hebrew_date) As fixed_date                      ' 7.27
        Dim result As fixed_date = hebrew_new_year(h_date.year) + h_date.day - 1
        Dim days_to_add As Integer = 0
        Dim m As hebrew_month
        If h_date.month < tishri Then
            m = tishri
            While m <= last_month_of_hebrew_year(h_date.year)
                days_to_add += last_day_of_hebrew_month(m, h_date.year)
                m = m + 1
            End While
            m = nisan
            While m < h_date.month
                days_to_add += last_day_of_hebrew_month(m, h_date.year)
                m = m + 1
            End While
        Else
            m = tishri
            While m < h_date.month
                days_to_add += last_day_of_hebrew_month(m, h_date.year)
                m = m + 1
            End While
        End If
        Return result + days_to_add
    End Function

    Public Function hebrew_from_fixed([date] As fixed_date) As hebrew_date                      ' 7.28
        Dim approx As Integer = Math.Floor((98496 / 35975351) * ([date] - hebrew_epoch())) + 1
        Dim year As hebrew_year = approx - 1
        While hebrew_new_year(year) <= [date]
            year += 1
        End While
        year = year - 1

        Dim start As hebrew_month
        If [date] < fixed_from_hebrew(New hebrew_date(year, nisan, 1)) Then
            start = tishri
        Else
            start = nisan
        End If

        Dim month As hebrew_month = start
        While [date] > fixed_from_hebrew(New hebrew_date(year, month, last_day_of_hebrew_month(month, year)))
            month = month + 1
        End While

        Dim day As hebrew_day = 1 + [date] - fixed_from_hebrew(New hebrew_date(year, month, 1))

        Return New hebrew_date(year, month, day)

    End Function


#Region "Holydays"

    Public Function trumpets(h_year As hebrew_year) As fixed_date
        Return fixed_from_hebrew(New hebrew_date(h_year, tishri, 1))
    End Function

    Public Function atonement(h_year As hebrew_year) As fixed_date
        Return fixed_from_hebrew(New hebrew_date(h_year, tishri, 10))
    End Function

    Public Function feast_of_tabernacles(h_year As hebrew_year) As fixed_date
        Return fixed_from_hebrew(New hebrew_date(h_year, tishri, 15))
    End Function

    Public Function last_great_day(h_year As hebrew_year) As fixed_date
        Return fixed_from_hebrew(New hebrew_date(h_year, tishri, 22))
    End Function

    Public Function passover(h_year As hebrew_year) As fixed_date
        Return fixed_from_hebrew(New hebrew_date(h_year - 1, nisan, 14))
    End Function

    Public Function ulb1(h_year As hebrew_year) As fixed_date
        Return fixed_from_hebrew(New hebrew_date(h_year - 1, nisan, 15))
    End Function

    Public Function ulb2(h_year As hebrew_year) As fixed_date
        Return fixed_from_hebrew(New hebrew_date(h_year - 1, nisan, 21))
    End Function

    Public Function pentecost(h_Year As hebrew_year) As fixed_date
        Dim day_of_week = fixed_from_hebrew(New hebrew_date(h_Year - 1, nisan, 15)) Mod 7
        Dim days_to_add As Integer = 0
        Select Case day_of_week
            Case 0
                days_to_add = -1
            Case 1
                days_to_add = 5
            Case 2
                days_to_add = 4
            Case 3
                days_to_add = 3
            Case 4
                days_to_add = 2
            Case 5
                days_to_add = 1
        End Select
        Return ulb1(h_Year) + days_to_add + 50
    End Function

#End Region


End Class
