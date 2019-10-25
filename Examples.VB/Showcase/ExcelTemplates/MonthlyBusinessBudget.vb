Imports System.Drawing

Namespace Showcase
    Public Class MonthlyBusinessBudget
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Monthly business budget.xlsx from resource
            Dim fileStream = GetResourceStream("Monthly business budget.xlsx")
            workbook.Open(fileStream)

            Dim worksheet = workbook.ActiveSheet

            ' change table style to TableStyleMedium14
            Dim totalsTable = worksheet.Tables!TotalsTable
            totalsTable.TableStyle = workbook.TableStyles("TableStyleMedium14")

            ' change chart type to column stacked
            Dim chart = worksheet.Shapes(0).Chart
            chart.ChartType = ChartType.ColumnStacked
            chart.ColumnGroups(0).Overlap = 100

            ' set a bigger font size for chart title
            chart.ChartTitle.Font.Size = 24
            chart.ChartTitle.Font.Bold = True

            ' give a one color gradient to chart area
            chart.ChartArea.Format.Fill.OneColorGradient(GradientStyle.Horizontal, 1, 0)
            chart.ChartArea.Format.Fill.GradientStops(0).Color.ObjectThemeColor = ThemeColor.Accent6
            chart.ChartArea.Format.Fill.GradientStops(0).Color.Brightness = 0.8

            ' make fill of plot area transparant
            chart.PlotArea.Format.Fill.Transparency = 1

            ' config series1 of chart
            Dim series1 As ISeries = chart.SeriesCollection(0)
            series1.HasDataLabels = True
            series1.Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent1

            ' give some formatting for the first point of series1
            Dim point1 As IPoint = series1.Points(0)
            point1.Format.Line.Color.RGB = Color.Black
            point1.Format.Line.Weight = 2
            point1.Format.Line.DashStyle = LineDashStyle.Dash

            ' config series2 of chart
            Dim series2 As ISeries = chart.SeriesCollection(1)
            series2.HasDataLabels = True
            series2.Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent6
            series2.DataLabels.Font.Color.RGB = Color.Red

            ' get the value axis
            Dim value_axis As IAxis = chart.Axes.Item(AxisType.Value, AxisGroup.Primary)

            ' show the display unit as thousands for value axis
            value_axis.HasDisplayUnitLabel = True
            value_axis.DisplayUnit = DisplayUnit.Thousands

            ' give a color for the major grid line of value axis
            value_axis.MajorGridlines.Format.Line.Color.ObjectThemeColor = ThemeColor.Accent6
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Monthly business budget.xlsx"
            End Get
        End Property

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Monthly business budget.xlsx"}
            End Get
        End Property
    End Class
End Namespace
