Namespace Features.Worksheets
    Public Class ActivateWorksheet
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets.Add()
            'Activate new created worksheet.
            worksheet.Activate()
        End Sub
    End Class
End Namespace
