Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigureScaling
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)
            Dim stream As IO.Stream = GetResourceStream("logo.png")

            Dim picture As IShape = sheet.Shapes.AddPicture(stream, ImageType.PNG, 20, 20, 395, 60)

            sheet.Range("B2:D4").Value = "Text"
            sheet.PageSetup.PrintGridlines = True

            'Set scaling.
            sheet.PageSetup.Zoom = 200
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
