Namespace Features.RangeOperations
    Public Class HiddenRowColumn
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range!E1.Value = 1

            'Hidden row 2:6.
            worksheet.Range("2:6").Hidden = True

            'Hidden column A:D.
            worksheet.Range("A:D").Hidden = True
        End Sub
    End Class
End Namespace
