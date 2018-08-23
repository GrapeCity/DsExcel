Namespace Features.Formatting.Alignment
    Public Class RotateCellContents
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim rangeB2 As IRange = worksheet.Range!B2
            rangeB2.Value = "Rotated Cell Contents"
            rangeB2.HorizontalAlignment = HorizontalAlignment.Center
            rangeB2.VerticalAlignment = VerticalAlignment.Center

            'Rotate cell contents.
            rangeB2.Orientation = 15
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
