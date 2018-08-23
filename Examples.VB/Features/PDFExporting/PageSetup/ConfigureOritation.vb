Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigureOritation
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)
            sheet.Range("A1:G10").Value = "Text"

            'Set page orientation.
            sheet.PageSetup.Orientation = PageOrientation.Landscape
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
