Namespace Features.PageSetup
    Public Class ConfigPageMargins
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetTemplateStream()
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set margins, in points.
            worksheet.PageSetup.TopMargin = 36
            worksheet.PageSetup.BottomMargin = 36
            worksheet.PageSetup.RightMargin = 72
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Demo.xlsx"
            End Get
        End Property
    End Class
End Namespace
