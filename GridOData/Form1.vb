Imports System.IO

Public Class Form1

    Dim data As New List(Of YearInfo)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim greg As New gregorian
        Dim juli As New julian
        Dim heb As New my_hebrew
        Dim new_heb As New hebrew

        For i As Integer = 1 To 50
            Dim newRow As New YearInfo

            ' important holy day stuff
            newRow.GregYear = If(i > 0, i, i - 1).ToString + " " + If(i > 0, "AD", "BC")
            newRow.JuliYear = If(i > 0, i, i - 1).ToString + " " + If(i > 0, "AD", "BC")
            newRow.HebYear = i + 1 - greg.gregorian_year_from_fixed(heb.hebrew_epoch)
            newRow.PentecostGregOldStyle = greg.gregorian_from_fixed(heb.pentecost(newRow.HebYear)).ToString
            newRow.Sivan6GregOldStyle = greg.gregorian_from_fixed(heb.passover(newRow.HebYear) + 50).ToString
            newRow.Tishri1GregOldStyle = greg.gregorian_from_fixed(heb.fixed_from_hebrew(New hebrew_date(newRow.HebYear, 7, 1))).ToString
            newRow.Tishri1JuliOldStyle = juli.julian_from_fixed(heb.fixed_from_hebrew(New hebrew_date(newRow.HebYear, 7, 1))).ToString
            newRow.Tishri1GregNewStyle = greg.gregorian_from_fixed(new_heb.fixed_from_hebrew(New hebrew_date(newRow.HebYear, 7, 1))).ToString
            newRow.Tishri1JuliNewStyle = juli.julian_from_fixed(new_heb.fixed_from_hebrew(New hebrew_date(newRow.HebYear, 7, 1))).ToString
            newRow.PentecostGregNewStyle = greg.gregorian_from_fixed(new_heb.pentecost(newRow.HebYear)).ToString
            newRow.Sivan6GregNewStyle = greg.gregorian_from_fixed(new_heb.passover(newRow.HebYear) + 50).ToString
            newRow.SummerStartsFixed = SummerStartGreg(i)
            newRow.SummerStartsDate = greg.gregorian_from_fixed(newRow.SummerStartsFixed).ToString
            newRow.SummerStartsTime = Basic.clock_from_moment(newRow.SummerStartsFixed).ToString


            newRow.PentecostGregOldStyleSunsetStartFixed = SunsetOnDate((heb.pentecost(newRow.HebYear) - 1))
            newRow.PentecostGregOldStyleSunsetEndfixed = SunsetOnDate((heb.pentecost(newRow.HebYear)))
            newRow.PentecostGregOldStyleSunsetStartDate = greg.gregorian_from_fixed(newRow.PentecostGregOldStyleSunsetStartFixed).ToString
            newRow.PentecostGregOldStyleSunsetEndDate = greg.gregorian_from_fixed(newRow.PentecostGregOldStyleSunsetEndfixed).ToString
            newRow.PentecostGregOldStyleSunsetStartTime = Basic.clock_from_moment(newRow.PentecostGregOldStyleSunsetStartFixed).ToString
            newRow.PentecostGregOldStyleSunsetEndTime = Basic.clock_from_moment(newRow.PentecostGregOldStyleSunsetEndfixed).ToString

            newRow.Sivan6GregOldStyleSunsetStartFixed = SunsetOnDate(heb.passover(newRow.HebYear) + 48)
            newRow.Sivan6GregOldStyleSunsetEndfixed = SunsetOnDate(heb.passover(newRow.HebYear) + 49)
            newRow.Sivan6GregOldStyleSunsetStartDate = greg.gregorian_from_fixed(newRow.Sivan6GregOldStyleSunsetStartFixed).ToString
            newRow.Sivan6GregOldStyleSunsetEndDate = greg.gregorian_from_fixed(newRow.Sivan6GregOldStyleSunsetEndfixed).ToString
            newRow.Sivan6GregOldStyleSunsetStartTime = Basic.clock_from_moment(newRow.Sivan6GregOldStyleSunsetStartFixed).ToString
            newRow.Sivan6GregOldStyleSunsetEndTime = Basic.clock_from_moment(newRow.Sivan6GregOldStyleSunsetEndfixed).ToString

            newRow.PentecostGregNewStyleSunsetStartFixed = SunsetOnDate((heb.pentecost(newRow.HebYear) - 1))
            newRow.PentecostGregNewStyleSunsetEndfixed = SunsetOnDate((heb.pentecost(newRow.HebYear)))
            newRow.PentecostGregNewStyleSunsetStartDate = greg.gregorian_from_fixed(newRow.PentecostGregNewStyleSunsetStartFixed).ToString
            newRow.PentecostGregNewStyleSunsetEndDate = greg.gregorian_from_fixed(newRow.PentecostGregNewStyleSunsetEndfixed).ToString
            newRow.PentecostGregNewStyleSunsetStartTime = Basic.clock_from_moment(newRow.PentecostGregNewStyleSunsetStartFixed).ToString
            newRow.PentecostGregNewStyleSunsetEndTime = Basic.clock_from_moment(newRow.PentecostGregNewStyleSunsetEndfixed).ToString

            newRow.Sivan6GregNewStyleSunsetStartFixed = SunsetOnDate(heb.passover(newRow.HebYear) + 48)
            newRow.Sivan6GregNewStyleSunsetEndfixed = SunsetOnDate(heb.passover(newRow.HebYear) + 49)
            newRow.Sivan6GregNewStyleSunsetStartDate = greg.gregorian_from_fixed(newRow.Sivan6GregNewStyleSunsetStartFixed).ToString
            newRow.Sivan6GregNewStyleSunsetEndDate = greg.gregorian_from_fixed(newRow.Sivan6GregNewStyleSunsetEndfixed).ToString
            newRow.Sivan6GregNewStyleSunsetStartTime = Basic.clock_from_moment(newRow.Sivan6GregNewStyleSunsetStartFixed).ToString
            newRow.Sivan6GregNewStyleSunsetEndTime = Basic.clock_from_moment(newRow.Sivan6GregNewStyleSunsetEndfixed).ToString



            data.Add(newRow)
        Next
        DataGridView1.DataSource = data

    
    End Sub

    Public Function SummerStartGreg(g_year As Integer) As Decimal

        Dim strPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
        strPath = strPath.Substring(6, strPath.Length - 6)
        
        Dim start_info As New ProcessStartInfo(strPath + "\GCL-2.6.1\lib\gcl-2.6.1\unixport\saved_gcl")
        start_info.Arguments = String.Format("{0} -load ""{1}\summer_start.cl"" -eval ""(use-package 'CC3)"" -batch", g_year, strPath)

        start_info.UseShellExecute = False
        start_info.CreateNoWindow = True
        start_info.RedirectStandardOutput = True
        start_info.RedirectStandardError = True

        ' Make the process and set its start information.
        Dim proc As New Process()
        proc.StartInfo = start_info

        ' Start the process.
        proc.Start()

        ' Attach to stdout and stderr.
        Dim std_out As StreamReader = proc.StandardOutput()
        Dim std_err As StreamReader = proc.StandardError()

        ' Display the results.
        Dim result As String = std_out.ReadToEnd()
        
        ' Clean up.
        std_out.Close()
        std_err.Close()
        proc.Close()
        Return Decimal.Parse(result)
    End Function

    Public Function SunsetOnDate(f_date As Integer) As Decimal
        Dim strPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
        strPath = strPath.Substring(6, strPath.Length - 6)

        Dim start_info As New ProcessStartInfo(strPath + "\GCL-2.6.1\lib\gcl-2.6.1\unixport\saved_gcl")
        start_info.Arguments = String.Format("{0} -load ""{1}\sunset.cl"" -eval ""(use-package 'CC3)"" -batch", f_date, strPath)

        start_info.UseShellExecute = False
        start_info.CreateNoWindow = True
        start_info.RedirectStandardOutput = True
        start_info.RedirectStandardError = True

        ' Make the process and set its start information.
        Dim proc As New Process()
        proc.StartInfo = start_info

        ' Start the process.
        proc.Start()

        ' Attach to stdout and stderr.
        Dim std_out As StreamReader = proc.StandardOutput()
        Dim std_err As StreamReader = proc.StandardError()

        ' Display the results.
        Dim result As String = std_out.ReadToEnd()

        ' Clean up.
        std_out.Close()
        std_err.Close()
        proc.Close()
        Return Decimal.Parse(result)
    End Function

End Class
