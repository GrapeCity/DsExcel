Namespace Features.Grouping
    Public Class CreateRangeGroup
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            '1:20 rows' outline level will be 2.
            worksheet.Range("1:20").Group()

            '1:10 rows' outline level will be 3.
            worksheet.Range("1:10").Group()

            'A:N columns' outline level will be 2.
            worksheet.Range("A:N").Group()

            'A:E columns' outline level will be 3.
            worksheet.Range("A:E").Group()
        End Sub
    End Class
End Namespace
