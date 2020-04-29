Namespace Features.PDFExporting.ExportShape
    Public Class ShapeWithGradientAndPatternFill
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            ' Add a sun
            Dim sun As IShape = sheet.Shapes.AddShape(AutoShapeType.Sun, 20, 15, 200, 200)

            ' Set gradient fill for sun
            sun.Fill.OneColorGradient(GradientStyle.Horizontal, 1, 0.8)
            sun.Fill.GradientStops.Insert(&Hff6347, 0)
            sun.Fill.GradientStops.Insert(&HFFA500, 0.2)
            sun.Fill.GradientStops.Insert(&HFFFF00, 0.4)
            sun.Fill.GradientStops.Insert(&HFF0000, 0.6)
            sun.Fill.GradientStops.Insert(&HFF, 0.8)
            sun.Fill.GradientStops.Insert(&H800080, 1)

            ' Add a five point star
            Dim star As IShape = sheet.Shapes.AddShape(AutoShapeType.Shape5pointStar, 250, 15, 200, 200)

            ' Set pattern fill for star
            star.Fill.Patterned(PatternType.DottedDiamond)
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
    End Class
End Namespace
