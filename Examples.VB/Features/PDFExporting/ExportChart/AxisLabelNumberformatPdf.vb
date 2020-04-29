Imports System.Drawing

Namespace Features.PDFExporting.ExportChart
    Public Class AxisLabelNumberformatPdf
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 20, 200, 360, 230)

            worksheet.Range("A1:D6").Value = New Object(,) {
                {Nothing, "S1", "S2", "S3"},
                {1, -25, 25, 25},
                {2, 51, 36, 27},
                {3, 52, 80, 30},
                {4, 22, -20, 65},
                {5, 23, 69, 69}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            With shape.Chart.Axes.Item(AxisType.Category)
                'config tick label's format
                .TickLabelPosition = TickLabelPosition.NextToAxis
                .TickLabelSpacing = 2
                .TickLabels.Font.Color.RGB = Color.DarkOrange
                .TickLabels.Font.Size = 12
                .TickLabels.NumberFormat = "#,##0.00"
            End With

            Dim value_axis As IAxis = shape.Chart.Axes.Item(AxisType.Value)

            value_axis.TickLabels.NumberFormat = "#,##0;[Red]#,##0"
        End Sub
    End Class
End Namespace
