Namespace Features.PDFExporting.ExportShape
    Public Class BasicShapes
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            ' Add a rectangle
            sheet.Shapes.AddShape(AutoShapeType.Rectangle, 20, 10, 100, 100)

            ' Add an oval
            sheet.Shapes.AddShape(AutoShapeType.Oval, 20, 160, 100, 100)

            ' Add a trangle
            sheet.Shapes.AddShape(AutoShapeType.IsoscelesTriangle, 20, 310, 100, 100)

            ' Add a "Not Allowed" symbol
            sheet.Shapes.AddShape(AutoShapeType.NoSymbol, 20, 460, 100, 100)

            ' Add a "Smile Face" symbol
            sheet.Shapes.AddShape(AutoShapeType.SmileyFace, 20, 600, 100, 100)

            ' Add a "Heart" symbol
            sheet.Shapes.AddShape(AutoShapeType.Heart, 170, 10, 100, 100)

            ' Add a "Sun" symbol
            sheet.Shapes.AddShape(AutoShapeType.Sun, 170, 160, 100, 100)

            ' Add a RightArrow
            sheet.Shapes.AddShape(AutoShapeType.RightArrow, 170, 310, 100, 100)

            ' Add a CurvedRightArrow
            sheet.Shapes.AddShape(AutoShapeType.CurvedRightArrow, 170, 460, 100, 100)

            ' Add a QuadArrow
            sheet.Shapes.AddShape(AutoShapeType.QuadArrow, 170, 600, 100, 100)

            ' Add a MathNotEqual
            sheet.Shapes.AddShape(AutoShapeType.MathNotEqual, 320, 10, 100, 100)

            ' Add a FlowchartMultidocument
            sheet.Shapes.AddShape(AutoShapeType.FlowchartMultidocument, 320, 160, 100, 100)

            ' Add a five points star
            sheet.Shapes.AddShape(AutoShapeType.Shape5pointStar, 320, 310, 100, 100)

            ' Add a CurvedUpRibbon
            sheet.Shapes.AddShape(AutoShapeType.CurvedUpRibbon, 320, 460, 100, 100)

            ' Add a OvalCallout
            sheet.Shapes.AddShape(AutoShapeType.OvalCallout, 320, 580, 100, 100)
        End Sub

        Public Overrides ReadOnly Property SavePdf As Boolean
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
