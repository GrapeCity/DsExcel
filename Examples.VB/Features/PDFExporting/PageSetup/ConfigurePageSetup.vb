Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigurePageSetup
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            'Set data.
            sheet.Range("A1:G10").Value = "Text"

            'Print rowheader and columnheader.
            sheet.PageSetup.PrintHeadings = True

            'Print gridlines.
            sheet.PageSetup.PrintGridlines = True
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
    End Class
End Namespace
