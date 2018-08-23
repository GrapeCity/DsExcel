Namespace Features.PageSetup
    Public Class ConfigPrintTitleColumns
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetTemplateStream()
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set columns to repeat at left
            worksheet.PageSetup.PrintTitleColumns = "$D:$G"
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Print Title Columns.xlsx"
            End Get
        End Property
    End Class
End Namespace
