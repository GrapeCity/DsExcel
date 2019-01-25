Namespace Features.PageSetup
    Public Class ConfigPrintTitleRows
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetResourceStream("PageSetup Print Title Rows.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set rows to repeat at top
            worksheet.PageSetup.PrintTitleRows = "$5:$10"
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Print Title Rows.xlsx"
            End Get
        End Property
    End Class
End Namespace
