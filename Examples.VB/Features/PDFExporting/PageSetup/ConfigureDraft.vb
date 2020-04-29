Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigureDraft
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            'Set text.
            sheet.Range("A1:G10").Value = "Text"

            'Add picture in sheet.
            Dim stream As IO.Stream = GetResourceStream("logo.png")
            Dim picture As IShape = sheet.Shapes.AddPicture(stream, ImageType.PNG, 20, 20, 395, 60)

            'Add header graphic.
            Dim stream1 As IO.Stream = GetResourceStream("logo.png")
            sheet.PageSetup.CenterHeader = "&G"
            sheet.PageSetup.CenterHeaderPicture.SetGraphicStream(stream1, ImageType.PNG)
            sheet.PageSetup.CenterHeaderPicture.Width = 100
            sheet.PageSetup.CenterHeaderPicture.Height = 13
            'Set print without graphics in page content area.
            sheet.PageSetup.Draft = True
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
