Namespace Features.Worksheets
    Public Class DeleteWorksheet
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets.Add()
            'workbook must contain one visible worksheet at least, if delete the one visible worksheet, it will throw exception.
            worksheet.Delete()
        End Sub
    End Class
End Namespace
