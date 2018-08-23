Namespace Features.Formatting.Fonts
    Public Class FontSize
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = "Font size is 15"
            worksheet.Range!A1.Font.Size = 15
        End Sub
    End Class
End Namespace
