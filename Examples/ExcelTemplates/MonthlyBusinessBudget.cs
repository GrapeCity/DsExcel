using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelTemplates
{
    public class MonthlyBusinessBudget : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Monthly business budget.xlsx from resource
            var fileStream = this.GetResourceStream("xlsx\\Monthly business budget.xlsx");
            workbook.Open(fileStream);

            var worksheet = workbook.ActiveSheet;

            // change table style to TableStyleMedium14
            var totalsTable = worksheet.Tables["TotalsTable"];
            totalsTable.TableStyle = workbook.TableStyles["TableStyleMedium14"];

            // change chart type to column stacked
            var chart = worksheet.Shapes[0].Chart;
            chart.ChartType = ChartType.ColumnStacked;
            chart.ColumnGroups[0].Overlap = 100;

            // set a bigger font size for chart title
            chart.ChartTitle.Font.Size = 24;
            chart.ChartTitle.Font.Bold = true;

            // give a one color gradient to chart area
            chart.ChartArea.Format.Fill.OneColorGradient(GradientStyle.Horizontal, 1, 0);
            chart.ChartArea.Format.Fill.GradientStops[0].Color.ObjectThemeColor = ThemeColor.Accent6;
            chart.ChartArea.Format.Fill.GradientStops[0].Color.Brightness = 0.8;

            // make fill of plot area transparant
            chart.PlotArea.Format.Fill.Transparency = 1;

            // config series1 of chart
            ISeries series1 = chart.SeriesCollection[0];
            series1.HasDataLabels = true;
            series1.Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent1;

            // give some formatting for the first point of series1
            IPoint point1 = series1.Points[0];
            point1.Format.Line.Color.RGB = Color.Black;
            point1.Format.Line.Weight = 2;
            point1.Format.Line.DashStyle = LineDashStyle.Dash;

            // config series2 of chart
            ISeries series2 = chart.SeriesCollection[1];
            series2.HasDataLabels = true;
            series2.Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent6;
            series2.DataLabels.Font.Color.RGB = Color.Red;

            // get the value axis
            IAxis value_axis = chart.Axes.Item(AxisType.Value, AxisGroup.Primary);

            // show the display unit as thousands for value axis
            value_axis.HasDisplayUnitLabel = true;
            value_axis.DisplayUnit = DisplayUnit.Thousands;
            
            // give a color for the major grid line of value axis
            value_axis.MajorGridlines.Format.Line.Color.ObjectThemeColor = ThemeColor.Accent6;
        }

        public override string TemplateName
        {
            get
            {
                return "Monthly business budget.xlsx";
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
        
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Monthly business budget.xlsx" };
            }
        }
    }
}
