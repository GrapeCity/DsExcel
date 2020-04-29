Namespace Features.PDFExporting.ExportShape
    Public Class ShapeWithPictureOrTextureFill
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            ' Get stream of picture
            Dim stream As IO.Stream = GetResourceStream("logo.png")

            ' Add a rectangle
            Dim rectangle As IShape = sheet.Shapes.AddShape(AutoShapeType.Rectangle, 20, 20, 250, 50)
            ' Set picture fill
            rectangle.Fill.UserPicture(stream, ImageType.PNG)
            rectangle.Line.Transparency = 1

            ' Add a oval
            Dim oval As IShape = sheet.Shapes.AddShape(AutoShapeType.Oval, 20, 90, 250, 50)
            ' Set picture fill
            oval.Fill.UserPicture(stream, ImageType.PNG)
            oval.Line.Color.RGB = System.Drawing.Color.FromArgb(&H49129E)

            ' Add a five point star
            Dim star As IShape = sheet.Shapes.AddShape(AutoShapeType.Shape5pointStar, 300, 20, 100, 100)
            ' Set picture fill
            star.Fill.PresetTextured(PresetTexture.WaterDroplets)
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

        Public Overrides ReadOnly Property ShowScreenshot As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"logo.png"}
            End Get
        End Property
    End Class
End Namespace
