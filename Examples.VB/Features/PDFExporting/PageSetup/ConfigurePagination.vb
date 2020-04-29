Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigurePagination
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            Dim row As Integer = 50
            Dim column As Integer = 14
            Dim data(row - 1, column - 1) As Object
            For i As Integer = 0 To row - 1
                For j As Integer = 0 To column - 1
                    data(i, j) = "R" & i & "C" & j
                Next j
            Next i

            'Set data
            sheet.Range("A1:N50").Value = data

            'Set paper size
            sheet.PageSetup.PaperSize = PaperSize.A5

            'With API the margin's unit is pound, with Excel the margins display unit is inch.
            'One inch eaquals to 72 pounds. If the top margin is 0.8 inch, then please set PageSetup.TopMargin = 0.8*72(57.6)
            sheet.PageSetup.TopMargin = 57.6

            'Top margin in excel is 0.8 inch
            sheet.PageSetup.BottomMargin = 21.6
            sheet.PageSetup.LeftMargin = 28.8
            sheet.PageSetup.RightMargin = 0
            sheet.PageSetup.HeaderMargin = 0
            sheet.PageSetup.FooterMargin = 93.6
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
