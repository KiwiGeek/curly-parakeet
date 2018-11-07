Public Module astronomical

    Public Function degrees(theta As Decimal) As Decimal
        While theta >= 360
            theta = theta - 360
        End While
        While theta < 0
            theta = theta + 360
        End While
        Return theta
    End Function

    Public Function radians_from_degrees(theta As Decimal) As Decimal
        Return degrees(theta) * (Math.PI * (1 / 180))
    End Function

    Public Function degrees_from_radians(theta As Decimal) As Decimal
        Return degrees(theta / Math.PI / (1 / 180))
    End Function

    Public Function hr(x As Decimal) As Decimal
        Return x / 24
    End Function

    Public Function sec(x As Decimal) As Decimal
        Return x / 24 * 60 * 60
    End Function

    Public Function mt(x As Decimal) As Decimal
        Return x
    End Function

    Public Function deg(x) As Decimal
        Return x
    End Function

    Public Function deg(x As List(Of Decimal)) As List(Of Decimal)
        Dim result = New List(Of Decimal)
        For Each bop In x
            result.Add(degrees(bop))
        Next
        Return result
    End Function

    Public Function secs(x As Decimal) As Decimal
        Return x / 3600
    End Function

    Public Function angle(d As Decimal, m As Decimal, s As Decimal) As Decimal
        Return d + (m / 60) + (s / 3600)
    End Function

    Public Function sin_degrees(theta) As Decimal
        Return Math.Sin(radians_from_degrees(theta))
    End Function

    Public Function cosine_degrees(theta) As Decimal
        Return Math.Cos(radians_from_degrees(theta))
    End Function

    Public Function tanent_degrees(theta) As Decimal
        Return Math.Tan(radians_from_degrees(theta))
    End Function

    Public Function arcsin_degrees(x As Decimal) As Decimal
        Return Math.Asin(degrees_from_radians(x))
    End Function

    Public Function direction(locale As location, focus As location) As Decimal             ' 13.5
        Dim y As Decimal = sin_degrees(focus.longitude - locale.longitude)
        Dim x As Decimal = cosine_degrees(locale.latitude) * tanent_degrees(focus.latitude) - sin_degrees(locale.latitude) * cosine_degrees(locale.longitude - focus.longitude)
        If (x = 0 And y = 0) Or focus.latitude = 90 Then Return deg(0)
        If focus.latitude = -90 Then Return deg(180)
        Return arctan_degrees(y, x)

    End Function

    Public Function arctan_degrees(y As Decimal, x As Decimal) As Decimal
        Dim a = degrees_from_radians(Math.Atan(y / x))
        If x = 0 And y <> 0 Then
            Return (signum(y) * deg(90)) Mod 360
        ElseIf x >= 0 Then
            Return a Mod 360
        Else
            Return (a + deg(180)) Mod 360
        End If
    End Function

    Public Function signum(y As Decimal) As Integer
        If y < 0 Then Return -1
        If y = 0 Then Return 0
        Return 1
    End Function

    Public Function mecca() As location
        Return New location(angle(21, 25, 24), angle(39, 49, 24), mt(298), hr(3))
    End Function

    Public Function jerusalem() As location
        Return New location(deg(31.8), deg(35.2), mt(800), hr(2))
    End Function

    Public Function zone_from_longitude(phi As Decimal) As Decimal
        Return phi / deg(360)
    End Function

    Public Function universal_from_local(tee_ell As Decimal, locale As location) As Decimal
        Return tee_ell - zone_from_longitude(locale.longitude)
    End Function

    Public Function local_from_universal(tee_rom_u As Decimal, locale As location) As Decimal
        Return tee_rom_u + zone_from_longitude(locale.longitude)
    End Function

    Public Function standard_from_universal(tee_rom_u As Decimal, locale As location) As Decimal
        Return tee_rom_u + locale.zone
    End Function

    Public Function universal_from_standard(tee_rom_s As Decimal, locale As location) As Decimal
        Return tee_rom_s - locale.zone
    End Function

    Public Function standard_from_local(tee_ell As Decimal, locale As location) As Decimal
        Return standard_from_universal(universal_from_local(tee_ell, locale), locale)
    End Function

    Public Function local_from_standard(teel_rom_s As Decimal, locale As location) As Decimal
        Return local_from_universal(universal_from_standard(teel_rom_s, locale), locale)
    End Function

    Public Function ephemeris_correction(tee As Decimal) As Decimal
        Dim greg As New gregorian
        Dim year As Decimal = greg.gregorian_year_from_fixed(Math.Floor(tee))
        Dim c As Decimal = (1 / 26525) * greg.gregorian_date_difference(New gregorian_date(1900, gregorian.january, 1), New gregorian_date(year, gregorian.january, 1))
        Dim x = hr(12) + greg.gregorian_date_difference(New gregorian_date(1810, gregorian.january, 1), New gregorian_date(year, gregorian.january, 1))
        If 1988 <= year And year <= 2019 Then
            Return (1 / 86400) * (year - 1933)
        ElseIf 1900 <= year And year <= 1987 Then
            Return -0.0002 + 0.000297 * c + 0.025184 * c ^ 2 -
                0.181133 * c ^ 3 + 0.55304 * c ^ 4 -
                0.861938 * c ^ 5 + 0.677066 * c ^ 6 -
                0.212591 * c ^ 7
        ElseIf 1800 <= year And year <= 1899 Then
            Return -0.000009 + 0.003844 * c + 0.083563 * c ^ 2 +
                0.865736 * c ^ 3 + 4.867575 * c ^ 4 + 15.845535 * c ^ 5 +
                31.332267 * c ^ 6 + 38.291999 * c ^ 7 + 28.316289 * c ^ 8 +
                11.636204 * c ^ 9 + 2.043794 * c ^ 10
        ElseIf 1700 <= year And year <= 1799 Then
            Return (1 / 86400) * (8.118780842 -
                                0.005092142 * (year - 1700) +
                                0.003336121 * (year - 1700) ^ 2 -
                                0.0000266484 * (year - 1700) ^ 3)
        ElseIf 1620 <= year And year <= 1699 Then
            Return (1 / 86400) * (196.58333 -
                                4.0675 * (year - 1600) +
                                0.0219167 * (year - 1600) ^ 2)
        Else
            Return (1 / 86400) * ((x ^ 2 / 41048480) - 15)
        End If
    End Function

    Public Function dynamical_from_universal(t As Decimal) As Decimal
        Return t + ephemeris_correction(t)
    End Function

    Public Function universal_from_dynamical(t As Decimal) As Decimal
        Return t - ephemeris_correction(t)
    End Function

    Public Function j2000() As Decimal
        Return hr(12) + (New gregorian).gregorian_new_year(2000)
    End Function

    Public Function julian_centuries(t As Decimal) As Decimal
        Return (1 / 36525) * (dynamical_from_universal(t) - j2000)
    End Function

    Public Function equation_of_time(t As Decimal) As Decimal
        Dim c As Decimal = julian_centuries(t)
        Dim λ As Decimal = deg(280.46645) + deg(36000.46983) * c + 0.003032 + c ^ 2
        Dim anomaly As Decimal = 357.5291 + 35999.0503 * c - 0.0001559 * c ^ 2 - 0.00000048 * c ^ 3
        Dim eccentricity As Decimal = 0.016708617 - 0.000042037 * c - 0.0000001236 * c ^ 2
        Dim ε As Decimal = obliquity(t)
        Dim y As Decimal = (tanent_degrees(ε / 2)) ^ 2
        Dim equation = (1 / 2 * Math.PI) *
            (y * sin_degrees(2 * λ) -
             2 * eccentricity * sin_degrees(anomaly) +
             4 * eccentricity * y * sin_degrees(anomaly) * cosine_degrees(2 * λ) -
             0.5 * y ^ 2 * sin_degrees(4 * λ) -
             1.25 * eccentricity ^ 2 *
             sin_degrees(2 * anomaly))
        Return signum(equation) * If(Math.Abs(equation) < 0.5, Math.Abs(equation), hr(12))
    End Function

    Public Function apparent_from_local(t As Decimal, locale As location) As Decimal
        Return t + equation_of_time(universal_from_local(t, locale))
    End Function

    Public Function local_from_apparent(t As Decimal, locale As location) As Decimal
        Return t - equation_of_time(universal_from_local(t, locale))
    End Function

    Public Function midnight([date] As Long, locale As location) As Decimal
        Return standard_from_local(local_from_apparent([date], locale), locale)
    End Function

    Public Function midday([date] As Long, locale As location) As Decimal
        Return standard_from_local(local_from_apparent([date] + hr(12), locale), locale)
    End Function

    Public Function sidereal_from_moment(t As Decimal) As Decimal
        Dim c As Decimal = (t - j2000()) / 36525
        Return [mod](280.46061837 + 36525 * 360.98564736629 * c + 0.000387933 * c ^ 2 - (1 / 38710000) * c ^ 3, 360)
    End Function

    Private Function obliquity(t As Decimal) As Decimal
        Dim c As Decimal = julian_centuries(t)
        Return angle(23, 26, 21.448) + (angle(0, 0, -46.815) * c + angle(0, 0, -0.00059) * c ^ 2 + angle(0, 0, 0.001813) * c ^ 3)
    End Function

    Private Function declination(t As Decimal, β As Decimal, λ As Decimal) As Decimal
        Dim ε As Decimal = obliquity(t)
        Return arcsin_degrees(sin_degrees(β) * cosine_degrees(ε) + cosine_degrees(β) * sin_degrees(ε) * sin_degrees(λ))
    End Function

    Public Function right_ascension(t As Decimal, β As Decimal, λ As Decimal) As Decimal
        Dim ε As Decimal = obliquity(t)
        Return arctan_degrees(sin_degrees(λ) * cosine_degrees(ε) - tanent_degrees(β) * sin_degrees(ε), cosine_degrees(λ))
    End Function

    Public Function mean_tropical_year() As Decimal
        Return 365.242189
    End Function

    Public Function mean_sidereal_year() As Decimal
        Return 365.25636
    End Function

    Public Function solar_logitude(t As Decimal) As Decimal
        Dim c As Decimal = julian_centuries(t)
        Dim vectorX = {403406, 195207, 119433, 112392, 3891, 2819, 1721, 660, 350, 334, 314, 268, 242, 234, 158, 132, 129, 114, 99, 93, 86, 78, 72, 68, 64, 46, 28, 37, 32, 29, 28, 27, 27, 25, 24, 21, 21, 20, 18, 17, 14, 13, 13, 13, 12, 10, 10, 10, 10}
        Dim vectorY = {270.54861, 340.19128, 63.91854, 331.2622, 317.843, 86.631, 240.052, 310.26, 247.23, 260.87, 297.82, 343.14, 166.79, 81.53, 3.5, 132.75, 182.95, 162.03, 29.8, 266.4, 249.2, 157.6, 257.8, 185.1, 69.9, 8, 197.1, 250.4, 65.3, 162.7, 341.5, 291.6, 98.5, 146.7, 110, 5.2, 243.6, 230.9, 256.1, 45.3, 242.9, 115.2, 151.8, 285.3, 53.3, 126.6, 205.7, 85.9, 146.1}
        Dim vectorZ = {0.9287892, 35999.1376958, 35999.4089666, 35998.7287385, 71998.20261, 71998.4403, 36000.35726, 71997.4814, 32964.4678, -19.441, 445267.1117, 45036.884, 3.1008, 22518.4434, -19.9739, 65928.9345, 9038.0293, 3034.7684, 33718.148, 3034.448, -2280.773, 29929.992, 31556.493, 149.588, 9037.75, 107997.405, -4444.176, 151.771, 67555.316, 31556.08, -4561.54, 107996.706, 1221.655, 62894.167, 31437.369, 14578.298, -31931.757, 34777.243, 1221.999, 62894.511, -4442.039, 107997.909, 119.066, 16859.071, -4.578, 26895.292, -39.127, 12297.536, 90073.778}
        Dim sum As Decimal = 0
        For i As Integer = 0 To vectorX.Count - 1
            sum += (vectorX(i) * sin_degrees(vectorY(i) + vectorZ(i) * c))
        Next
        Dim λ As Decimal = 282.7771834 + 36000.76953744 * c + 0.0000057295779513082322 * sum
        Return [mod](λ + aberration(t) + nutation(t), 360)
    End Function

    Public Function nutation(t As Decimal) As Decimal
        Dim c As Decimal = julian_centuries(t)
        Dim A As Decimal = 129.9 - 1934.134 * c + 0.002063 * c ^ 2
        Dim B As Decimal = 201.11 + 72001.5377 * c + 0.00057 * c ^ 2
        Return -0.004778 * sin_degrees(A) - 0.0003667 * sin_degrees(B)
    End Function

    Public Function aberration(t As Decimal) As Decimal
        Dim c As Decimal = julian_centuries(t)
        Return 0.0000974 * cosine_degrees(177.63 + 35999.01848 * c) - 0.005575
    End Function

    Public Function solar_logitude_after(λ As Decimal, t As Decimal) As Decimal
        Dim rate As Decimal = mean_tropical_year() / 360
        Dim τ = t + rate * [mod](λ - solar_logitude(t), 360)
        Dim a = If(t > τ - 5, t, τ - 5)
        Dim b = τ + 5
        ' Return solar_logitude(λ, )
    End Function
End Module
