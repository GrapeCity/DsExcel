Namespace Features.Charts
    Public Class ChartCopy
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Create chart, chart's range is Range("G1:M21")
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 300, 10, 300, 300)
            worksheet.Range("A1:D6").Value = New Object(,)  {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            'Range("G1:M21") must contain chart's range, copy a new shape to Range("N1:T21")
            worksheet.Range("G1:M21").Copy(worksheet.Range!N1)
            'worksheet.Range("G1:M21").Copy(worksheet.Range("N1:T21"))
            'Cross sheet copy, copy a new chart to worksheet2's Range("N1:T21")
            'Dim worksheet2 As IWorksheet = workbook.Worksheets.Add()
            'worksheet.Range("G1:M21").Copy(worksheet2.Range!E1)
            'worksheet.Range("G1:M21").Copy(worksheet2.Range("N1:T21"))
        End Sub
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
