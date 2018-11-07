Public Class location
    Public Property latitude As Decimal
    Public Property longitude As Decimal
    Public Property elevation As Decimal
    Public Property zone As Decimal

    Sub New(_lat As Decimal, _long As Decimal, _elev As Decimal, _zone As Decimal)
        Me.latitude = _lat
        Me.longitude = _long
        Me.elevation = _elev
        Me.zone = _zone
    End Sub


    Public Overrides Function ToString() As String
        Return String.Format("{0} {1} {2} {3}", latitude, longitude, elevation, zone)
    End Function
End Class
