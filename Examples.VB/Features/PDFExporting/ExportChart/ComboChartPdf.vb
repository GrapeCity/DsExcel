Namespace Features.PDFExporting.ExportChart
    Public Class ComboChartPdf
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 20, 20, 360, 230)
            worksheet.Range("A20:C36").Value = New Object(,) {
                {"Area 1", "Column 1", "Column 2"},
                {1350, 120, 75}, {1500, 90, 35}, {1200, 80, 50}, {1300, 80, 80},
                {1750, 90, 100}, {1640, 120, 130}, {1700, 120, 95}, {1100, 90, 80},
                {1350, 120, 75}, {1500, 90, 35}, {1200, 80, 50}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A20:C36"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Combination Chart"

            Dim series1 As ISeries = shape.Chart.SeriesCollection(0)
            Dim series2 As ISeries = shape.Chart.SeriesCollection(1)
            Dim series3 As ISeries = shape.Chart.SeriesCollection(2)

            'change series type
            series1.ChartType = ChartType.Area
            series2.ChartType = ChartType.ColumnStacked
            series3.ChartType = ChartType.ColumnStacked

            'set axis group
            series2.AxisGroup = AxisGroup.Secondary
            series3.AxisGroup = AxisGroup.Secondary

            'config axis sacle and unit
            With shape.Chart.Axes.Item(AxisType.Value)
                .MaximumScale = 1800
                .MajorUnit = 450
            End With

            With shape.Chart.Axes.Item(AxisType.Value, AxisGroup.Secondary)
                .MaximumScale = 300
                .MajorUnit = 75
            End With
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

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
    End Class
End Namespace
