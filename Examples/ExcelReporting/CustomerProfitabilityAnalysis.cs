using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelReporting
{
    public class CustomerProfitabilityAnalysis : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] data = new object[,] {
	            { null, "[Segment Name]", "[Segment Name]", "[Segment Name]", "Overall" },
	            { "Customer Activity:", null, null, null, null },
	            { "Number of active customers—Beginning of period", 5, 8, 8, null },
	            { "Number of customers added", 2, 4, 4, null },
	            { "Number of customers lost/terminated", -1, -2, -2, null },
	            { "Number of active customers—End of period", null, null, null, null },
	            { null, null, null, null, null },
	            { "Profitability Analysis:", null, null, null, null },
	            { "Revenue per segment", 1500000, 1800000, 2500000, null },
	            { "Weighting", null, null, null, null },
	            { null, null, null, null, null },
	            { "Cost of sales:", null, null, null, null },
	            { "Ongoing service and support costs", 1000000, 1400000, 1400000, null },
	            { "Other direct customer costs", 200000, 100000, 100000, null },
	            { "Total cost of sales", null, null, null, null },
	            { "Gross margin", null, null, null, null },
	            { "Weighting", null, null, null, null },
	            { null, null, null, null, null },
	            { "Other costs:", null, null, null, null },
	            { "Customer acquisition", 105000, 120000, 235000, null },
	            { "Customer marketing", 150000, 125000, 275000, null },
	            { "Customer termination", 80000, 190000, 140000, null },
	            { "Total other customer costs", null, null, null, null },
	            { "Customer profit by segment", null, null, null, null },
	            { "Weighting", null, null, null, null },
	            { null, null, null, null, null },
	            { "Summary Metrics:", "[Segment Name]", "[Segment Name]", "[Segment Name]", "Trend" },
	            { "Average cost per acquired customer", null, null, null, null },
	            { "Average cost per terminated customer", null, null, null, null },
	            { "Average marketing cost per active customer", null, null, null, null },
	            { "Average profit (loss) per customer", null, null, null, null },
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Customer Profitability";
            worksheet.TabColor = Color.FromRGB(131, 172, 121);
            worksheet.SheetView.DisplayGridlines = false;

            //Set Value.
            worksheet.Range["B2"].Value = "[Company Name]";
            worksheet.Range["B3"].Value = "Customer Profitability Analysis";
            worksheet.Range["B4"].Value = "[Date]";
            worksheet.Range["B6"].Value = "Gray cells will be calculated for you. You do not need to enter anything into them.";
            worksheet.Range["B7:F37"].Value = data;

            //Set formula.
            worksheet.Range["F9:F11"].Formula = "=SUM(C9:E9)";
            worksheet.Range["C12:F12"].Formula = "=SUM(C9:C11)";

            worksheet.Range["C16:E16"].Formula = "=+C15/$F$15";
            worksheet.Range["F15:F16"].Formula = "=SUM(C15:E15)";

            worksheet.Range["F19:F20"].Formula = "=SUM(C19:E19)";
            worksheet.Range["C21:F21"].Formula = "=SUM(C19:C20)";
            worksheet.Range["C22:F22"].Formula = "=+C15-C21";
            worksheet.Range["C23:E23"].Formula = "=MAX(0, MIN(1,C22/$F$22))";
            worksheet.Range["F23"].Formula = "=SUM(C23:E23)";

            worksheet.Range["F26:F28"].Formula = "=SUM(C26:E26)";
            worksheet.Range["C29:F29"].Formula = "=SUM(C26:C28)";
            worksheet.Range["C30:E30"].Formula = "=+C22-C29";
            worksheet.Range["C31:E31"].Formula = "=MAX(0,MIN(1, C30/$F$30))";
            worksheet.Range["F30:F31"].Formula = "=SUM(C30:E30)";

            worksheet.Range["C34:E34"].Formula = "=+C26/C10";
            worksheet.Range["C35:E35"].Formula = "=-C28/C11";
            worksheet.Range["C36:E36"].Formula = "=+C27/C12";
            worksheet.Range["C37:E37"].Formula = "=+C30/C12";

            //Change the range's RowHeight and ColumnWidth.
            worksheet.StandardHeight = 15;
            worksheet.StandardWidth = 9.140625;
            worksheet.Rows[0].RowHeight = 9.95;
            worksheet.Rows[1].RowHeight = 33;
            worksheet.Rows[2].RowHeight = 27;
            worksheet.Rows[3].RowHeight = 19.5;
            worksheet.Rows[4].RowHeight = 9;
            worksheet.Rows[5].RowHeight = 19.5;
            worksheet.Rows[6].RowHeight = 18;
            worksheet.Rows[12].RowHeight = 9;
            worksheet.Rows[16].RowHeight = 9;
            worksheet.Rows[23].RowHeight = 9;
            worksheet.Rows[31].RowHeight = 9;

            worksheet.Columns[0].ColumnWidth = 1.85546875;
            worksheet.Columns[1].ColumnWidth = 46.7109375;
            worksheet.Columns[2].ColumnWidth = 16.42578125;
            worksheet.Columns[3].ColumnWidth = 16.42578125;
            worksheet.Columns[4].ColumnWidth = 16.42578125;
            worksheet.Columns[5].ColumnWidth = 16.42578125;

            //Modify the build in name styles.
            var nameStyle_Normal = workbook.Styles["Normal"];
            nameStyle_Normal.VerticalAlignment = VerticalAlignment.Center;
            nameStyle_Normal.Font.ThemeColor = ThemeColor.Dark1;
            nameStyle_Normal.Font.TintAndShade = 0.249946592608417;
            nameStyle_Normal.Font.Size = 10;

            var nameStyle_Heading_1 = workbook.Styles["Heading 1"];
            nameStyle_Heading_1.HorizontalAlignment = HorizontalAlignment.Left;
            nameStyle_Heading_1.VerticalAlignment = VerticalAlignment.Center;
            nameStyle_Heading_1.Font.ThemeFont = ThemeFont.Major;
            nameStyle_Heading_1.Font.Bold = false;
            nameStyle_Heading_1.Font.Size = 24;
            nameStyle_Heading_1.Font.ThemeColor = ThemeColor.Dark1;
            nameStyle_Heading_1.Font.TintAndShade = 0.249946592608417;
            nameStyle_Heading_1.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            nameStyle_Heading_1.IncludeAlignment = true;

            var nameStyle_Heading_2 = workbook.Styles["Heading 2"];
            nameStyle_Heading_2.HorizontalAlignment = HorizontalAlignment.Left;
            nameStyle_Heading_2.VerticalAlignment = VerticalAlignment.Center;
            nameStyle_Heading_2.Font.ThemeFont = ThemeFont.Major;
            nameStyle_Heading_2.Font.Bold = false;
            nameStyle_Heading_2.Font.Size = 20;
            nameStyle_Heading_2.Font.ThemeColor = ThemeColor.Dark1;
            nameStyle_Heading_2.Font.TintAndShade = 0.249946592608417;
            nameStyle_Heading_2.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            nameStyle_Heading_2.Interior.ThemeColor = ThemeColor.Accent3;
            nameStyle_Heading_2.Interior.TintAndShade = 0.39994506668294322;
            nameStyle_Heading_2.IncludeNumber = true;
            nameStyle_Heading_2.IncludePatterns = true;

            var nameStyle_Heading_3 = workbook.Styles["Heading 3"];
            nameStyle_Heading_3.HorizontalAlignment = HorizontalAlignment.Left;
            nameStyle_Heading_3.VerticalAlignment = VerticalAlignment.Center;
            nameStyle_Heading_3.Font.ThemeFont = ThemeFont.Major;
            nameStyle_Heading_3.Font.Bold = false;
            nameStyle_Heading_3.Font.Size = 14;
            nameStyle_Heading_3.Font.ThemeColor = ThemeColor.Dark1;
            nameStyle_Heading_3.Font.TintAndShade = 0.249946592608417;
            nameStyle_Heading_3.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            nameStyle_Heading_3.IncludeAlignment = true;
            nameStyle_Heading_3.IncludePatterns = true;

            var nameStyle_Heading_4 = workbook.Styles["Heading 4"];
            nameStyle_Heading_4.HorizontalAlignment = HorizontalAlignment.Left;
            nameStyle_Heading_4.VerticalAlignment = VerticalAlignment.Center;
            nameStyle_Heading_4.Font.ThemeFont = ThemeFont.Major;
            nameStyle_Heading_4.Font.Bold = true;
            nameStyle_Heading_4.Font.Size = 10;
            nameStyle_Heading_4.Font.ThemeColor = ThemeColor.Light1;
            nameStyle_Heading_4.Font.TintAndShade = -0.0499893185216834;
            nameStyle_Heading_4.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            nameStyle_Heading_4.Interior.ThemeColor = ThemeColor.Accent3;
            nameStyle_Heading_4.Interior.TintAndShade = -0.249946592608417;
            nameStyle_Heading_4.IncludeAlignment = true;
            nameStyle_Heading_4.IncludeBorder = true;
            nameStyle_Heading_4.IncludePatterns = true;

            //Apply the above name styles on ranges.
            worksheet.Range["B2:F2"].Style = workbook.Styles["Heading 1"];
            worksheet.Range["B3:F3"].Style = workbook.Styles["Heading 2"];
            worksheet.Range["B4:F4"].Style = workbook.Styles["Heading 3"];
            worksheet.Range["B8:F8"].Style = workbook.Styles["Heading 4"];
            worksheet.Range["B14:F14"].Style = workbook.Styles["Heading 4"];
            worksheet.Range["B18:F18"].Style = workbook.Styles["Heading 4"];
            worksheet.Range["B25:F25"].Style = workbook.Styles["Heading 4"];
            worksheet.Range["B33:F33"].Style = workbook.Styles["Heading 4"];

            //Set NumberFormat.
            worksheet.Range["C9:F12"].NumberFormat = "0_);[Red](0)";
            worksheet.Range["C15:F15"].NumberFormat = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
            worksheet.Range["C16:F16"].NumberFormat = "0%";
            worksheet.Range["C19:F22"].NumberFormat = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
            worksheet.Range["C23:F23"].NumberFormat = "0%";
            worksheet.Range["C26:F30"].NumberFormat = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
            worksheet.Range["C31:F31"].NumberFormat = "0%";
            worksheet.Range["C34:F37"].NumberFormat = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";

            //Set range's font style.
            worksheet.Range["B6"].Font.TintAndShade = 0.34998626667073579;
            worksheet.Range["B6"].Font.Size = 8;
            worksheet.Range["B6"].Font.Italic = true;
            worksheet.Range["C7:F7"].Font.TintAndShade = 0;
            worksheet.Range["B9:F12"].Font.TintAndShade = 0;
            worksheet.Range["B15:F16"].Font.TintAndShade = 0;
            worksheet.Range["B19:F23"].Font.TintAndShade = 0;
            worksheet.Range["B26:F31"].Font.TintAndShade = 0;
            worksheet.Range["B34:F37"].Font.TintAndShade = 0;
            worksheet.Range["C33:F33"].Font.Bold = false;

            //Set range's alignment.
            worksheet.Range["C7:F7"].HorizontalAlignment = HorizontalAlignment.Center;
            worksheet.Range["C33:F33"].HorizontalAlignment = HorizontalAlignment.Center;

            //Set range's border
            worksheet.Range["B9:F12"].Borders.LineStyle = BorderLineStyle.Thin;
            worksheet.Range["B9:F12"].Borders.ThemeColor = ThemeColor.Accent3;
            worksheet.Range["B9:F12"].Borders.TintAndShade = 0.39994506668294322;

            worksheet.Range["B15:F16"].Borders.LineStyle = BorderLineStyle.Thin;
            worksheet.Range["B15:F16"].Borders.ThemeColor = ThemeColor.Accent3;
            worksheet.Range["B15:F16"].Borders.TintAndShade = 0.39994506668294322;

            worksheet.Range["B19:F23"].Borders.LineStyle = BorderLineStyle.Thin;
            worksheet.Range["B19:F23"].Borders.ThemeColor = ThemeColor.Accent3;
            worksheet.Range["B19:F23"].Borders.TintAndShade = 0.39994506668294322;

            worksheet.Range["B26:F31"].Borders.LineStyle = BorderLineStyle.Thin;
            worksheet.Range["B26:F31"].Borders.ThemeColor = ThemeColor.Accent3;
            worksheet.Range["B26:F31"].Borders.TintAndShade = 0.39994506668294322;

            worksheet.Range["B34:F37"].Borders.LineStyle = BorderLineStyle.Thin;
            worksheet.Range["B34:F37"].Borders.ThemeColor = ThemeColor.Accent3;
            worksheet.Range["B34:F37"].Borders.TintAndShade = 0.39994506668294322;

            //Set range's fill.
            worksheet.Range["F9:F12"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["F9:F12"].Interior.TintAndShade = -0.0499893185216834;
            worksheet.Range["C12:E12"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["C12:E12"].Interior.TintAndShade = -0.0499893185216834;
            worksheet.Range["F15:F16"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["F15:F16"].Interior.TintAndShade = -0.0499893185216834;
            worksheet.Range["C16:E16"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["C16:E16"].Interior.TintAndShade = -0.0499893185216834;
            worksheet.Range["F19:F23"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["F19:F23"].Interior.TintAndShade = -0.0499893185216834;
            worksheet.Range["C21:E23"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["C21:E23"].Interior.TintAndShade = -0.0499893185216834;
            worksheet.Range["F26:F31"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["F26:F31"].Interior.TintAndShade = -0.0499893185216834;
            worksheet.Range["C29:E31"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["C29:E31"].Interior.TintAndShade = -0.0499893185216834;
            worksheet.Range["C34:E37"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["C34:E37"].Interior.TintAndShade = -0.0499893185216834;

            //create a new group of sparklines.
            ISparklineGroup sparklineGroup = worksheet.Range["F34:F37"].SparklineGroups.Add(SparkType.Line, "C34:E37");
            sparklineGroup.SeriesColor.ThemeColor = ThemeColor.Accent3;
            sparklineGroup.SeriesColor.TintAndShade = -0.249977111117893;
            sparklineGroup.Points.Negative.Color.ThemeColor = ThemeColor.Accent4;
            sparklineGroup.Points.Markers.Color.ThemeColor = ThemeColor.Accent4;
            sparklineGroup.Points.Markers.Color.TintAndShade = -0.249977111117893;
            sparklineGroup.Points.Highpoint.Color.ThemeColor = ThemeColor.Accent4;
            sparklineGroup.Points.Highpoint.Color.TintAndShade = -0.249977111117893;
            sparklineGroup.Points.Lowpoint.Color.ThemeColor = ThemeColor.Accent4;
            sparklineGroup.Points.Lowpoint.Color.TintAndShade = -0.249977111117893;
            sparklineGroup.Points.Firstpoint.Color.ThemeColor = ThemeColor.Accent4;
            sparklineGroup.Points.Firstpoint.Color.TintAndShade = -0.249977111117893;
            sparklineGroup.Points.Lastpoint.Color.ThemeColor = ThemeColor.Accent4;
            sparklineGroup.Points.Lastpoint.Color.TintAndShade = -0.249977111117893;
            sparklineGroup.Points.Negative.Visible = false;
            sparklineGroup.Points.Firstpoint.Visible = false;
            sparklineGroup.Points.Lastpoint.Visible = false;

            //Add chart.
            IShape shape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 9.75, 576.95, 590.25, 237);
            shape.Name = "Chart 3";

            //Add Series.
            ISeries series1 = shape.Chart.SeriesCollection.NewSeries();
            series1.Formula = "=SERIES('Customer Profitability'!$B$34,'Customer Profitability'!$C$33:$E$33,'Customer Profitability'!$C$34:$E$34,1)";
            series1.Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent2;

            ISeries series2 = shape.Chart.SeriesCollection.NewSeries();
            series2.Formula = "=SERIES('Customer Profitability'!$B$35,'Customer Profitability'!$C$33:$E$33,'Customer Profitability'!$C$35:$E$35,2)";
            series2.Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent4;

            ISeries series3 = shape.Chart.SeriesCollection.NewSeries();
            series3.Formula = "=SERIES('Customer Profitability'!$B$36,'Customer Profitability'!$C$33:$E$33,'Customer Profitability'!$C$36:$E$36,3)";
            series3.Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent3;

            ISeries series4 = shape.Chart.SeriesCollection.NewSeries();
            series4.Formula = "=SERIES('Customer Profitability'!$B$37,'Customer Profitability'!$C$33:$E$33,'Customer Profitability'!$C$37:$E$37,4)";
            series4.Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent5;

            //Set the char group's Overlap and GapWidth.
            shape.Chart.ColumnGroups[0].Overlap = 0;
            shape.Chart.ColumnGroups[0].GapWidth = 199;

            //Set the chart's title format.
            var chartTitle = shape.Chart.ChartTitle;
            chartTitle.Text = "Summary Metrics per Customer Segment";
            chartTitle.Font.ThemeFont = ThemeFont.Major;
            chartTitle.Font.Color.ObjectThemeColor = ThemeColor.Dark1;
            chartTitle.Font.Size = 20;

            //Set the chart legend's position.
            shape.Chart.Legend.Position = LegendPosition.Top;

            //Set category axis format.
            IAxis category_axis = shape.Chart.Axes.Item(AxisType.Category);
            category_axis.HasTitle = true;
            category_axis.AxisTitle.Text = "SEGMENT";
            category_axis.AxisTitle.Font.Size = 9;
            category_axis.AxisTitle.Font.ThemeFont = ThemeFont.Minor;

            //Set value axis format.
            IAxis value_axis = shape.Chart.Axes.Item(AxisType.Value);
            value_axis.CrossesAt = -200000;
            value_axis.HasMinorGridlines = true;
            value_axis.MinorGridlines.Format.Line.Color.ObjectThemeColor = ThemeColor.Dark1;
            value_axis.MinorGridlines.Format.Line.Color.Brightness = 0.95;

            //Create customize theme.
            Themes themes = new Themes();
            ITheme theme = themes.Add("test");
            theme.ThemeColorScheme[ThemeColor.Dark1].RGB = Color.FromRGB(0, 0, 0);
            theme.ThemeColorScheme[ThemeColor.Light1].RGB = Color.FromRGB(255, 255, 255);
            theme.ThemeColorScheme[ThemeColor.Dark2].RGB = Color.FromRGB(77, 70, 70);
            theme.ThemeColorScheme[ThemeColor.Light2].RGB = Color.FromRGB(255, 251, 239);
            theme.ThemeColorScheme[ThemeColor.Accent1].RGB = Color.FromRGB(255, 225, 132);
            theme.ThemeColorScheme[ThemeColor.Accent2].RGB = Color.FromRGB(102, 173, 166);
            theme.ThemeColorScheme[ThemeColor.Accent3].RGB = Color.FromRGB(131, 172, 121);
            theme.ThemeColorScheme[ThemeColor.Accent4].RGB = Color.FromRGB(254, 191, 102);
            theme.ThemeColorScheme[ThemeColor.Accent5].RGB = Color.FromRGB(219, 112, 87);
            theme.ThemeColorScheme[ThemeColor.Accent6].RGB = Color.FromRGB(165, 115, 137);
            theme.ThemeColorScheme[ThemeColor.Hyperlink].RGB = Color.FromRGB(102, 173, 166);
            theme.ThemeColorScheme[ThemeColor.FollowedHyperlink].RGB = Color.FromRGB(165, 115, 137);
            theme.ThemeFontScheme.Major[FontLanguageIndex.Latin].Name = "Marion";
            theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Marion";

            //Apply the above custom theme.
            workbook.Theme = theme;

            //Set active cell.
            worksheet.Range["B7"].Activate();
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
