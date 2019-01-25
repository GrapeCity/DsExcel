Namespace Features.PageSetup
    Public Class ConfigPageCenter
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetResourceStream("PageSetup Demo.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set center, default is False.
            worksheet.PageSetup.CenterHorizontally = True
            worksheet.PageSetup.CenterVertically = True
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Demo.xlsx"
            End Get
        End Property
    End Class
End Namespace
