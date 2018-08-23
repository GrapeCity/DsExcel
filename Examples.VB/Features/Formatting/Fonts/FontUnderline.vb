Namespace Features.Formatting.Fonts
    Public Class FontUnderline
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = "Single Underline"
            worksheet.Range!A1.Font.Underline = UnderlineType.Single
        End Sub
    End Class
End Namespace
