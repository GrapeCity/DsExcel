Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigureHeaderFooter
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            'Set data.
            sheet.Range("A1:G10").Value = "Text"

            'Set page header.
            sheet.PageSetup.LeftHeader = "&""Arial,Italic""LeftHeader"
            sheet.PageSetup.RightHeader = "&KFF0000GrapeCity"
            sheet.PageSetup.CenterHeader = "&P"

            'Set page footer picture.
            Dim stream As IO.Stream = GetResourceStream("logo.png")
            sheet.PageSetup.CenterFooter = "&G"
            sheet.PageSetup.CenterFooterPicture.SetGraphicStream(stream, ImageType.PNG)
            sheet.PageSetup.CenterFooterPicture.Width = 100
            sheet.PageSetup.CenterFooterPicture.Height = 13
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
