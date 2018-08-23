Namespace Features.Formatting.Alignment
    Public Class TextOrientation
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim rangeC1 As IRange = worksheet.Range!C1
            rangeC1.Value = "The ReadingOrder property is applied to set text direction."
            rangeC1.ReadingOrder = ReadingOrder.RightToLeft
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
