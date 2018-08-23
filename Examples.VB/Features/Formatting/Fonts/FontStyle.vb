Namespace Features.Formatting.Fonts
    Public Class FontStyle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = "Bold"
            worksheet.Range!A1.Font.Bold = True
            worksheet.Range!A2.Value = "Italic"
            worksheet.Range!A2.Font.Italic = True
            worksheet.Range!A3.Value = "Bold Italic"
            worksheet.Range!A3.Font.Bold = True
            worksheet.Range!A3.Font.Italic = True
        End Sub
    End Class
End Namespace
