Namespace Features.Formatting.NumberFormat
    Public Class CustomNumberFormat
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set range values.
            worksheet.Range("A2:B2").Value = -15.5
            worksheet.Range("A3:B3").Value = 555
            worksheet.Range("A4:B4").Value = 0
            worksheet.Range("A5:B5").Value = "Name"

            'Apply custom number format.
            worksheet.Range("B2:B5").NumberFormat = "[Green]#.00;[Red]#.00;[Blue]0.00;[Cyan]""product: ""@"
        End Sub
    End Class
End Namespace
