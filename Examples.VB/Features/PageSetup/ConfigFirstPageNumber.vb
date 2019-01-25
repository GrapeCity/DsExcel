Namespace Features.PageSetup
    Public Class ConfigFirstPageNumber
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetResourceStream("PageSetup Demo.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.PageSetup.CenterFooter = "&P"

            'Set first page number, default is p1.
            worksheet.PageSetup.FirstPageNumber = 3
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Demo.xlsx"
            End Get
        End Property
    End Class
End Namespace
