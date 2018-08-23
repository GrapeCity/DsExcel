Namespace Features.PageSetup
    Public Class ConfigSheetPrintSettings
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetTemplateStream()
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set sheet
            worksheet.PageSetup.PrintGridlines = True
            worksheet.PageSetup.PrintHeadings = True
            worksheet.PageSetup.BlackAndWhite = True
            worksheet.PageSetup.PrintComments = PrintLocation.InPlace
            worksheet.PageSetup.PrintErrors = PrintErrors.Dash
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Demo.xlsx"
            End Get
        End Property
    End Class
End Namespace
