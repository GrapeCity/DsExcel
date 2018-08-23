Namespace Features.Grouping
    Public Class ClearOutline
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            '1:20 rows' outline level will be 2.
            worksheet.Range("1:20").Group()

            '1:10 rows' outline level will be 3.
            worksheet.Range("1:10").Group()

            '1:20 rows' outline level will be 1.
            worksheet.Range("1:20").ClearOutline()
        End Sub
    End Class
End Namespace
