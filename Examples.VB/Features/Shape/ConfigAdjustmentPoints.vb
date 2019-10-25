Namespace Features.Shape
    Public Class ConfigAdjustmentPoints
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            ' Add a right arrow callout.
            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.RightArrowCallout, 20, 20, 200, 100)

            Dim adjustments As IAdjustments = shape.Adjustments
            adjustments(0) = 0.5 ' arrow neck width
            adjustments(1) = 0.4 ' arrow head width
            adjustments(2) = 0.5 ' arrow head height
            adjustments(3) = 0.6 ' text box width
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

    End Class
End Namespace
