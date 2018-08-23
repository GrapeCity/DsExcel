Namespace Features.PDFExporting.Text
    Public Class Overflow
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            sheet.Range("F2, F4").Value = "This is a test string of overflow"
            sheet.Range("F6, F8").Value = "This is a test string of overflow with right alignment"
            sheet.Range("F6, F8").HorizontalAlignment = HorizontalAlignment.Right
            sheet.Range("D8, H4").Value = 123

            'Other settings
            sheet.Range("A1:J10").Borders.LineStyle = BorderLineStyle.Thin
        End Sub
        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
