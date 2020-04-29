Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigurePageBreaks
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            sheet.Range("A1:B5").Value = New Object(,)  {
                {1, 2},
                {3, 4},
                {5, 6},
                {7, 8},
                {9, 10}
            }

            'Add page break
            sheet.HPageBreaks.Add(sheet.Range!B3)
            sheet.VPageBreaks.Add(sheet.Range!B3)
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
