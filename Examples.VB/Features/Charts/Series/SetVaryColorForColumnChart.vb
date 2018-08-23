Namespace Features.Charts.Series
    Public Class SetVaryColorForColumnChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:B6").Value = {
                {Nothing, "S1"},
                {"Item1", 10},
                {"Item2", -51},
                {"Item3", 52},
                {"Item4", 22},
                {"Item5", 23}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B6"), RowCol.Columns, True, True)

            'set vary colors for column chart which only has one series.
            shape.Chart.ColumnGroups(0).VaryByCategories = True
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
