using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Showcase
{
    public class CostAnalysisWithParetoChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] data = new object[,] {
                { "Cost Center", "Annual Cost", "Percent of Total", "Cumulative Percent" },
                { "Parts and materials", 1325000, null, null },
                { "Manufacturing equipment", 900500, null, null },
                { "Salaries", 575000, null, null },
                { "Maintenance", 395000, null, null },
                { "Office lease", 295000, null, null },
                { "Warehouse lease", 250000, null, null },
                { "Insurance", 180000, null, null },
                { "Benefits and pensions", 130000, null, null },
                { "Vehicles", 125000, null, null },
                { "Research", 75000, null, null },
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "COST DATA and CHART";
            worksheet.TabColor = Color.FromArgb(63, 94, 101);
            worksheet.SheetView.DisplayGridlines = false;

            //Set Value.
            worksheet.Range["B2"].Value = "COST ANALYSIS - PARETO";
            worksheet.Range["B39"].Value = "COST ANALYSIS";
            worksheet.Range["B41:E51"].Value = data;

            //Set NumberFormat.
            worksheet.Range["C42:C52"].NumberFormat = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
            worksheet.Range["D42:E52"].NumberFormat = "0.00%";

            //Change the range's RowHeight and ColumnWidth.
            worksheet.StandardHeight = 15.75;
            worksheet.StandardWidth = 8.43;
            worksheet.Range["1:1"].RowHeight = 12.75;
            worksheet.Range["2:2"].RowHeight = 20.25;
            worksheet.Range["3:3"].RowHeight = 16.5;
            worksheet.Range["39:39"].RowHeight = 20.25;
            worksheet.Range["40:40"].RowHeight = 16.5;
            worksheet.Range["41:41"].RowHeight = 29.25;
            worksheet.Range["42:51"].RowHeight = 20.1;
            worksheet.Range["A:A"].ColumnWidth = 1.44140625;
            worksheet.Range["B:B"].ColumnWidth = 25.88671875;
            worksheet.Range["C:C"].ColumnWidth = 18.5546875;
            worksheet.Range["D:D"].ColumnWidth = 17.77734375;
            worksheet.Range["E:E"].ColumnWidth = 20.77734375;

            //Apply one build in name style on the ranges.
            worksheet.Range["B2:J2"].Style = workbook.Styles["Heading 1"];
            worksheet.Range["B39:E39"].Style = workbook.Styles["Heading 1"];

            //Create a table and apply one build in table style.
            ITable table = worksheet.Tables.Add(worksheet.Range["B41:E51"], true);
            table.Name = "tblData";
            table.TableStyle = workbook.TableStyles["TableStyleLight9"];
            table.ShowTotals = true;
            table.ShowTableStyleRowStripes = true;

            //Use table formula in table range.
            worksheet.Range["D42:D51"].Formula = "=[@[Annual Cost]]/SUM([[Annual Cost]])";
            worksheet.Range["E42:E51"].Formula = "=SUM(INDEX([Percent of Total],1):[@[Percent of Total]])";
            table.Columns[1].TotalsCalculation = TotalsCalculation.Sum;
            table.Columns[2].TotalsCalculation = TotalsCalculation.Sum;
            table.Columns[3].TotalsCalculation = TotalsCalculation.None;

            //Add chart.
            IShape shape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 9.75, 48, 597, 472.5);
            shape.Name = "Pareto Chart";

            //Add Series.
            ISeries series_ColumnClustered = shape.Chart.SeriesCollection.NewSeries();
            series_ColumnClustered.Formula = "=SERIES('COST DATA and CHART'!$C$41,'COST DATA and CHART'!$B$42:$B$51,'COST DATA and CHART'!$C$42:$C$51,1)";
            series_ColumnClustered.Format.Fill.TwoColorGradient(GradientStyle.Horizontal, 1);
            series_ColumnClustered.Format.Fill.GradientAngle = 90;
            series_ColumnClustered.Format.Fill.GradientStops[0].Color.ObjectThemeColor = ThemeColor.Accent1;
            series_ColumnClustered.Format.Fill.GradientStops[0].Color.Brightness = 0;
            series_ColumnClustered.Format.Fill.GradientStops[0].Position = 0;
            series_ColumnClustered.Format.Fill.GradientStops[1].Color.ObjectThemeColor = ThemeColor.Accent1;
            series_ColumnClustered.Format.Fill.GradientStops[1].Color.Brightness = -0.16;
            series_ColumnClustered.Format.Fill.GradientStops[1].Position = 1;
            series_ColumnClustered.Format.Line.Color.ObjectThemeColor = ThemeColor.Light1;

            ISeries series_Line = shape.Chart.SeriesCollection.NewSeries();
            series_Line.ChartType = ChartType.Line;
            series_Line.Formula = "=SERIES('COST DATA and CHART'!$E$41,,'COST DATA and CHART'!$E$42:$E$51,2)";
            series_Line.Format.Line.Weight = 2.25;
            series_Line.AxisGroup = AxisGroup.Secondary;

            //Change the secondary's maxinumscale.
            IAxis secondary_axis = shape.Chart.Axes.Item(AxisType.Value, AxisGroup.Secondary);
            secondary_axis.MaximumScale = 1.0;

            //Set the chart's title format.
            shape.Chart.ChartTitle.Text = "Cost Center";
            shape.Chart.ChartTitle.Font.ThemeFont = ThemeFont.Minor;
            shape.Chart.ChartTitle.Font.Color.RGB = Color.FromArgb(89, 89, 89);
            shape.Chart.ChartTitle.Font.Size = 18;

            //Set the chart has no legend.
            shape.Chart.HasLegend = false;

            //Set the char group's Overlap and GapWidth.
            shape.Chart.ColumnGroups[0].Overlap = 0;
            shape.Chart.ColumnGroups[0].GapWidth = 0;

            //Set chart area's format.
            shape.Chart.ChartArea.Format.Fill.PresetGradient(GradientStyle.Horizontal, 1, PresetGradientType.EarlySunset);
            shape.Chart.ChartArea.Format.Fill.GradientStops.Delete(3);
            shape.Chart.ChartArea.Format.Fill.GradientStops.Delete(3);
            shape.Chart.ChartArea.Format.Fill.GradientAngle = 90;
            shape.Chart.ChartArea.Format.Fill.GradientStops[0].Color.ObjectThemeColor = ThemeColor.Light1;
            shape.Chart.ChartArea.Format.Fill.GradientStops[0].Color.Brightness = 0;
            shape.Chart.ChartArea.Format.Fill.GradientStops[0].Position = 0;
            shape.Chart.ChartArea.Format.Fill.GradientStops[1].Color.ObjectThemeColor = ThemeColor.Light1;
            shape.Chart.ChartArea.Format.Fill.GradientStops[1].Color.Brightness = -0.15;
            shape.Chart.ChartArea.Format.Fill.GradientStops[1].Position = 0.68;
            shape.Chart.ChartArea.Format.Fill.GradientStops[2].Color.ObjectThemeColor = ThemeColor.Light1;
            shape.Chart.ChartArea.Format.Fill.GradientStops[2].Color.Brightness = 0;
            shape.Chart.ChartArea.Format.Fill.GradientStops[2].Position = 1;

            //Create customize theme.
            Themes themes = new Themes();
            ITheme theme = themes.Add("test");
            theme.ThemeColorScheme[ThemeColor.Dark1].RGB = Color.FromArgb(0, 0, 0);
            theme.ThemeColorScheme[ThemeColor.Light1].RGB = Color.FromArgb(255, 255, 255);
            theme.ThemeColorScheme[ThemeColor.Dark2].RGB = Color.FromArgb(96, 89, 88);
            theme.ThemeColorScheme[ThemeColor.Light2].RGB = Color.FromArgb(241, 246, 246);
            theme.ThemeColorScheme[ThemeColor.Accent1].RGB = Color.FromArgb(63, 94, 101);
            theme.ThemeColorScheme[ThemeColor.Accent2].RGB = Color.FromArgb(224, 170, 83);
            theme.ThemeColorScheme[ThemeColor.Accent3].RGB = Color.FromArgb(179, 29, 66);
            theme.ThemeColorScheme[ThemeColor.Accent4].RGB = Color.FromArgb(162, 67, 162);
            theme.ThemeColorScheme[ThemeColor.Accent5].RGB = Color.FromArgb(120, 59, 101);
            theme.ThemeColorScheme[ThemeColor.Accent6].RGB = Color.FromArgb(55, 120, 169);
            theme.ThemeColorScheme[ThemeColor.Hyperlink].RGB = Color.FromArgb(71, 166, 181);
            theme.ThemeColorScheme[ThemeColor.FollowedHyperlink].RGB = Color.FromArgb(120, 59, 101);
            theme.ThemeFontScheme.Major[FontLanguageIndex.Latin].Name = "Constantia";
            theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Helvetica";

            //Apply the above custom theme.
            workbook.Theme = theme;

            //Set active cell.
            worksheet.Range["B43"].Activate();
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
    }

}
