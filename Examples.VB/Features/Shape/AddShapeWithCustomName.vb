Namespace Features.Shape
    Public Class AddShapeWithCustomName
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim shape As IShape = worksheet.Shapes.AddShape("custom parallelogram", AutoShapeType.Parallelogram, 1, 1, 200, 100)

            'Get shape by name
            Dim parallelogram As IShape = worksheet.Shapes("custom parallelogram")
            parallelogram.Fill.Color.RGB = System.Drawing.Color.Red
        End Sub
    End Class
End Namespace
