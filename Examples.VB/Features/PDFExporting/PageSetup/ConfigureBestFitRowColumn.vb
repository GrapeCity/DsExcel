Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigureBestFitRowColumn
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            With sheet.Range
                ' Set text for some cells.
                !A1.Value = "Grapecity"
                !A2.Value = "Document For .NET"
                !B1.Value = "Grapecity"
                !B2.Value = "Excel for .NET"
                'Set font size of cell "A2"
                !A2.Font.Size = 20
            End With

            With sheet.PageSetup
                ' Set bestFitColumns/bestFitRows as true.
                .BestFitColumns = True
                .BestFitRows = True

                ' Set print gridline and heading.
                .PrintGridlines = True
                .PrintHeadings = True
            End With
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
