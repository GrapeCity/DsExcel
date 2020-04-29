Imports System.Drawing
Imports System.IO

Namespace Features.ImageExporting
    Public Class ConvertShapeToImage
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Add a oval
            Dim oval As IShape = worksheet.Shapes.AddShape(AutoShapeType.Oval, 0, 0, 191, 194)

            ' Set soild fill for rectangle
            oval.Fill.Color.RGB = Color.OrangeRed

            ' Save the shape as image to a stream.
            oval.ToImage(outputStream, ImageType.PNG)
        End Sub

        Public Overrides ReadOnly Property SaveAsImages As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
