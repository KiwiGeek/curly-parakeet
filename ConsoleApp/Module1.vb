
Imports CC = CalendricalCalculations
Imports CalendricalCalculations.location
Imports CalendricalCalculations.Basic
Imports CalendricalCalculations.astronomical

Module Module1

    Sub Main()

        'Dim list_of_rds = {-214193, -61387, 25469, 49217, 171307, 210155, 253427, 369740, 400085, 434355, 452605, 470160, 473837,
        '                   507850, 524156, 544676, 567118, 569477, 601716, 613424, 626596, 645554, 664224, 671401, 694799, 704424,
        '                   708842, 709409, 709580, 727274, 728714, 744313, 764652}

        Dim urbana As CC.location = New CC.location(deg(40.1), deg(-88.2), mt(225), hr(-6))
        Dim greenwich As CC.location = New CC.location(deg(51.4777815), deg(0), mt(46.9), hr(0))


        Console.ReadKey()
    End Sub


End Module
