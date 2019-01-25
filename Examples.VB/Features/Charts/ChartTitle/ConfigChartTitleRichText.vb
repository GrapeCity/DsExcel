Imports System.Drawing

Namespace Features.Charts.ChartTitle
    Public Class ConfigChartTitleRichText
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D6").Value = New Object(,)  {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -20, 36, 27},
                {"Item3", 62, 70, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 50, 50}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            'config chart title style with rich text
            shape.Chart.HasTitle = True
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs.Add("ChartSubtitle")
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs.Add("ChartTitle", 0)
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs(0).Font.Color.RGB = Color.CornflowerBlue
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs(0).Font.Size = 15
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = Color.Orange
            shape.Chart.ChartTitle.TextFrame.TextRange.Paragraphs(1).Font.Size = 10
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
