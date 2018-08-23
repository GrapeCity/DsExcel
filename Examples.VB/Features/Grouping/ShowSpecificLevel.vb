Namespace Features.Grouping
    Public Class ShowSpecificLevel
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A:N").Group()
            worksheet.Range("A:F").Group()
            worksheet.Range("A:C").Group()
            worksheet.Range("Q:Z").Group()
            worksheet.Range("Q:T").Group()

            'level 3 and level 4 will be collapsed. level 2 and level 1 expand.
            worksheet.Outline.ShowLevels(columnLevels:=2)
        End Sub
    End Class
End Namespace
