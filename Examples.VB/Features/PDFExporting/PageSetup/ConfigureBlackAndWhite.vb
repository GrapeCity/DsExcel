Imports System.Drawing

Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigureBlackAndWhite
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)
            Dim stream As IO.Stream = GetResourceStream("logo.png")
            Dim picture As IShape = sheet.Shapes.AddPicture(stream, ImageType.PNG, 20, 20, 395, 60)

            'Set text font color.
            sheet.Range("A1:D4").Value = "Font"
            sheet.Range("A1:D4").Font.Color = Color.Red

            'Set cell color
            sheet.Range("A7:D10").Value = "Green"
            sheet.Range("A7:D10").Interior.Color = Color.Green

            'Set print black and white.
            sheet.PageSetup.BlackAndWhite = True
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
