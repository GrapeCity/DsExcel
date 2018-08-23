Namespace Features.Shape
    Public Class ShapeCopy
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Create a shape in worksheet, shape's range is Range("A7:B7")
            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 1, 1, 100, 100)

            'Range("A1:D10") must contain Range("A7:B7"), copy a new shape to Range("C1:F7")
            worksheet.Range("A1:D10").Copy(worksheet.Range!E1)
            'worksheet.Range("A1:D10").Copy(worksheet.Range("E1:I9"))
            'Cross sheet copy, copy a new shape to worksheet2's Range("C1:F7")
            'Dim worksheet2 As IWorksheet = workbook.Worksheets.Add()
            'worksheet.Range("A1:D10").Copy(worksheet2.Range!E1)
            'worksheet.Range("A1:D10").Copy(worksheet2.Range("E1:I9"))
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
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
