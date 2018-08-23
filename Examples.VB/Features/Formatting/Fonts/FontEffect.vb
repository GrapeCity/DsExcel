Namespace Features.Formatting.Fonts
    Public Class FontEffect
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = "Strikethrough"
            worksheet.Range!A1.Font.Strikethrough = True
            worksheet.Range!A2.Value = "Superscript"
            worksheet.Range!A2.Font.Superscript = True
            worksheet.Range!A3.Value = "Subscript"
            worksheet.Range!A3.Font.Subscript = True
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowScreenshot As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
