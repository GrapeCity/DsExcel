Imports System.Drawing
Imports System.IO

Namespace Features.PDFExporting
    Public Class ConvertShapeToImage
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Add a rectangle
            Dim rectangle As IShape = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 0, 0, 191, 194)

            ' Set soild fill for rectangle
            rectangle.Fill.Color.RGB = Color.OrangeRed

            'NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            'Create a pdf document.
            Dim doc As New Pdf.GcPdfDocument
            Dim page As Pdf.Page = doc.NewPage()

            'Create a PrintManager.
            Dim printManager As New Excel.PrintManager

            'Get the size of the shape.
            Dim topLeftCell As IRange = rectangle.TopLeftCell
            Dim bottomRightCell As IRange = rectangle.BottomRightCell
            Dim shapeRange As IRange = worksheet.Range(topLeftCell.Row,
                                                       topLeftCell.Column,
                                                       bottomRightCell.Row - topLeftCell.Row + 1,
                                                       bottomRightCell.Column - topLeftCell.Column + 1)
            Dim size As SizeF = printManager.GetSize(shapeRange)

            'Modify the size of the page.
            page.Size = size

            ' Draw the Range"A1:E5" to the specified location on the page. 
            printManager.Draw(page, New PointF(0, 0), shapeRange)

            ' Saves the page as an image to a stream.
            page.SaveAsPng(outputStream, New Pdf.SaveAsImageOptions() With {.Resolution = 72})
        End Sub

        Public Overrides ReadOnly Property SaveAsImages As Boolean
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
