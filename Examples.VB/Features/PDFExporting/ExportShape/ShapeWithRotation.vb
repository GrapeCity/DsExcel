Namespace Features.PDFExporting.ExportShape
    Public Class ShapeWithRotation
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            sheet.Range!D2.Value = "rectangle with 30 degrees"
            sheet.Range!I2.Value = "right arrow with 40 degrees"

            ' Add a rectangle with rotation
            Dim rectangle As IShape = sheet.Shapes.AddShape(AutoShapeType.Rectangle, 50, 50, 200, 50)
            rectangle.Rotation = 30

            ' Add a right arrow with rotation
            Dim rightArrowWithRotation As IShape = sheet.Shapes.AddShape(AutoShapeType.RightArrow, 270, 50, 200, 100)
            rightArrowWithRotation.Rotation = 40
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
