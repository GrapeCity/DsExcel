Namespace Features.Shape
    Public Class ConfigShapeWithPictureFill
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.Parallelogram, 20, 20, 200, 100)
            Dim stream As IO.Stream = GetResourceStream("logo.png")

            shape.Fill.UserPicture(stream, ImageType.JPG)
            shape.PictureFormat.ColorType = PictureColorType.Grayscale
            shape.PictureFormat.Brightness = 0.6
            shape.PictureFormat.Contrast = 0.3
            shape.PictureFormat.Crop.PictureOffsetX = 10
            shape.PictureFormat.Crop.PictureOffsetY = -5
            shape.PictureFormat.Crop.PictureWidth = 120
            shape.PictureFormat.Crop.PictureHeight = 80
        End Sub
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
    End Class
End Namespace
