Namespace Features.Formatting.Fonts
    Public Class FontName
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'use Name property to set font name.
            worksheet.Range!A1.Value = "Calibri"
            worksheet.Range!A1.Font.Name = "Calibri"

            'use ThemeFont property to set font name.
            worksheet.Range!A2.Value = "Major theme font"
            worksheet.Range!A2.Font.ThemeFont = ThemeFont.Major
        End Sub
    End Class
End Namespace
