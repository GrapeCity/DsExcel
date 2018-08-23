Namespace Features.PageSetup
    Public Class ConfigPrintArea
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetTemplateStream()
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set print area & print titles
            worksheet.PageSetup.PrintArea = "$D$5:$G$10"
            worksheet.PageSetup.PrintTitleRows = "$5:$10"
            worksheet.PageSetup.PrintTitleColumns = "$D:$G"
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Print Area.xlsx"
            End Get
        End Property
    End Class
End Namespace
