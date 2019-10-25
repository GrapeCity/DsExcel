using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Showcase
{
    public class BloodPressureTracker : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            #region theme
            //create a custom theme.
            Themes themes = new Themes();
            ITheme theme = themes.Add("testTheme");
            theme.ThemeColorScheme[ThemeColor.Light1].RGB = Color.FromArgb(255, 255, 255);
            theme.ThemeColorScheme[ThemeColor.Dark1].RGB = Color.FromArgb(0, 0, 0);
            theme.ThemeColorScheme[ThemeColor.Light2].RGB = Color.FromArgb(222, 222, 212);
            theme.ThemeColorScheme[ThemeColor.Dark2].RGB = Color.FromArgb(30, 46, 47);
            theme.ThemeColorScheme[ThemeColor.Accent1].RGB = Color.FromArgb(233, 117, 90);
            theme.ThemeColorScheme[ThemeColor.Accent2].RGB = Color.FromArgb(122, 182, 186);
            theme.ThemeColorScheme[ThemeColor.Accent3].RGB = Color.FromArgb(125, 181, 135);
            theme.ThemeColorScheme[ThemeColor.Accent4].RGB = Color.FromArgb(230, 191, 94);
            theme.ThemeColorScheme[ThemeColor.Accent5].RGB = Color.FromArgb(230, 143, 77);
            theme.ThemeColorScheme[ThemeColor.Accent6].RGB = Color.FromArgb(194, 107, 112);
            theme.ThemeColorScheme[ThemeColor.Hyperlink].RGB = Color.FromArgb(122, 182, 186);
            theme.ThemeColorScheme[ThemeColor.FollowedHyperlink].RGB = Color.FromArgb(166, 140, 177);
            theme.ThemeFontScheme.Major[FontLanguageIndex.Latin].Name = "Gill Sans";
            theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Gill Sans";

            //assign the custom theme for workbook.
            workbook.Theme = theme;
            #endregion

            //does not show sheet gridlines.
            worksheet.SheetView.DisplayGridlines = false;

            #region RowHeightColumnWidth
            //set row height and column width.
            worksheet.StandardHeight = 12.75;
            worksheet.StandardWidth = 8.43;
            worksheet.Rows[1].RowHeight = 32.25;
            worksheet.Rows[2].RowHeight = 13.5;
            worksheet.Rows[3].RowHeight = 18.75;
            worksheet.Rows[6].RowHeight = 18.75;
            worksheet.Rows[9].RowHeight = 18.75;
            worksheet.Rows[12].RowHeight = 18.75;
            worksheet.Rows[15].RowHeight = 19.5;
            worksheet.Rows[16].RowHeight = 13.5;
            worksheet.Rows[33].RowHeight = 19.5;
            worksheet.Rows[34].RowHeight = 13.5;

            worksheet.Columns[0].ColumnWidth = 1.7109375;
            worksheet.Columns[1].ColumnWidth = 12.140625;
            worksheet.Columns[2].ColumnWidth = 12.140625;
            worksheet.Columns[3].ColumnWidth = 12.140625;
            worksheet.Columns[4].ColumnWidth = 11.85546875;
            worksheet.Columns[5].ColumnWidth = 12.7109375;
            worksheet.Columns[6].ColumnWidth = 13.85546875;
            worksheet.Columns[7].ColumnWidth = 44.7109375;
            #endregion

            #region Values
            //initialize worksheet's values.
            worksheet.Name = "BLOOD PRESSURE DATA";
            worksheet.Range["B2"].Value = "BLOOD PRESSURE TRACKER";
            worksheet.Range["B4:F13"].Value = new object[,] {
                                { "NAME", null, null, null, null },
                                { null, null, null, null, null },
                                { null, null, null, "Systolic", "Diastolic" },
                                { "TARGET BLOOD PRESSURE", null, null, 120, 80 },
                                { null, null, null, null, null },
                                { null, null, null, "Systolic", "Diastolic" },
                                { "CALL PHYSICIAN IF ABOVE", null, null, 140, 90 },
                                { null, null, null, null, null },
                                { null, null, null, null, null },
                                { "PHYSICIAN PHONE NUMBER", null, null, "[Phone Number]", null}
                            };

            worksheet.Range["B16"].Value = "CHARTED PROGRESS";
            worksheet.Range["B34"].Value = "DATA ENTRY";
            #endregion

            #region Table
            //initialize table data.
            worksheet.Range["B36:H44"].Value = new object[,] {
                                { "TIME", "DATE", "AM/PM", "SYSTOLIC", "DIASTOLIC", "HEART RATE", "NOTES" },
                                { new TimeSpan(10, 00, 00), DateTime.Parse("7/1/2013 12:00:00 AM"), "AM", 129, 99, 72, null },
                                { new TimeSpan(18, 00, 00), DateTime.Parse("7/1/2013 12:00:00 AM"), "PM", 133, 80, 75, null },
                                { new TimeSpan(10, 30, 00), DateTime.Parse("7/2/2012 12:00:00 AM"), "AM", 142, 86, 70, null },
                                { new TimeSpan(19, 00, 00), DateTime.Parse("7/2/2012 12:00:00 AM"), "PM", 141, 84, 68, null },
                                { new TimeSpan(09, 00, 00), DateTime.Parse("7/3/2012 12:00:00 AM"), "AM", 137, 84, 70, null },
                                { new TimeSpan(18, 30, 00), DateTime.Parse("7/3/2012 12:00:00 AM"), "PM", 139, 83, 72, null },
                                { new TimeSpan(10, 00, 00), DateTime.Parse("7/4/2012 12:00:00 AM"), "AM", 140, 85, 78, null },
                                { new TimeSpan(18, 00, 00), DateTime.Parse("7/4/2012 12:00:00 AM"), "PM", 138, 85, 69, null },
                            };
            ITable table = worksheet.Tables.Add(worksheet.Range["B36:H44"], true);
            table.ShowTotals = true;

            //set total row formulas.
            table.Columns[0].Total.Value = "Average";
            table.Columns[3].TotalsCalculation = TotalsCalculation.Average;
            table.Columns[4].TotalsCalculation = TotalsCalculation.Average;
            table.Columns[5].TotalsCalculation = TotalsCalculation.Average;
            table.Columns[6].TotalsCalculation = TotalsCalculation.None;

            //config data body range and total range's number format.
            table.Columns[0].DataBodyRange.NumberFormat = "h:mm;@";
            table.Columns[1].DataBodyRange.NumberFormat = "m/d/yyyy";
            table.Columns[3].DataBodyRange.NumberFormat = "0";
            table.Columns[4].DataBodyRange.NumberFormat = "0";
            table.Columns[5].DataBodyRange.NumberFormat = "0";
            table.Columns[3].Total.NumberFormat = "0";
            table.Columns[4].Total.NumberFormat = "0";
            table.Columns[5].Total.NumberFormat = "0";

            //config table range's alignment.
            table.Range.HorizontalAlignment = HorizontalAlignment.Left;
            table.Range.IndentLevel = 0;
            table.Range.VerticalAlignment = VerticalAlignment.Center;
            #endregion

            #region TableStyle
            //create a custom table style.
            ITableStyle tablestyle = workbook.TableStyles.Add("testStyle");
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Font.ThemeColor = ThemeColor.Dark1;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Font.TintAndShade = 0.25;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thin;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeTop].ThemeColor = ThemeColor.Accent1;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeTop].TintAndShade = 0.4;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].ThemeColor = ThemeColor.Accent1;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].TintAndShade = 0.4;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Accent1;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].TintAndShade = 0.4;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeLeft].ThemeColor = ThemeColor.Accent1;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeLeft].TintAndShade = 0.4;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeRight].ThemeColor = ThemeColor.Accent1;
            tablestyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeRight].TintAndShade = 0.4;

            tablestyle.TableStyleElements[TableStyleElementType.FirstRowStripe].Interior.ThemeColor = ThemeColor.Accent1;
            tablestyle.TableStyleElements[TableStyleElementType.FirstRowStripe].Interior.TintAndShade = 0.8;

            tablestyle.TableStyleElements[TableStyleElementType.HeaderRow].Font.Bold = true;
            tablestyle.TableStyleElements[TableStyleElementType.HeaderRow].Font.ThemeColor = ThemeColor.Dark1;
            tablestyle.TableStyleElements[TableStyleElementType.HeaderRow].Font.TintAndShade = 0.25;
            tablestyle.TableStyleElements[TableStyleElementType.HeaderRow].Interior.ThemeColor = ThemeColor.Accent1;

            tablestyle.TableStyleElements[TableStyleElementType.TotalRow].Font.Bold = true;
            tablestyle.TableStyleElements[TableStyleElementType.TotalRow].Font.ThemeColor = ThemeColor.Dark1;
            tablestyle.TableStyleElements[TableStyleElementType.TotalRow].Font.TintAndShade = 0.25;
            tablestyle.TableStyleElements[TableStyleElementType.TotalRow].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Double;
            tablestyle.TableStyleElements[TableStyleElementType.TotalRow].Borders[BordersIndex.EdgeTop].ThemeColor = ThemeColor.Accent1;

            //assign custom table style for table.
            table.TableStyle = workbook.TableStyles["testStyle"];
            #endregion

            #region Style
            //assign built-in styles for ranges.
            worksheet.Range["B2:H2"].Style = workbook.Styles["Heading 1"];
            worksheet.Range["B4:F4, B7:D7, B10:D10, B13:D13"].Style = workbook.Styles["Heading 2"];
            worksheet.Range["B16:H16, B34:H34"].Style = workbook.Styles["Heading 3"];

            //modify built-in styles.
            IStyle style_Heading1 = workbook.Styles["Heading 1"];
            style_Heading1.HorizontalAlignment = HorizontalAlignment.General;
            style_Heading1.VerticalAlignment = VerticalAlignment.Center;
            style_Heading1.Font.ThemeFont = ThemeFont.Major;
            style_Heading1.Font.Size = 24;
            style_Heading1.Font.Bold = true;
            style_Heading1.Font.ThemeColor = ThemeColor.Accent1;
            style_Heading1.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;
            style_Heading1.Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Accent1;

            style_Heading1.IncludeAlignment = true;
            style_Heading1.IncludeFont = true;
            style_Heading1.IncludeBorder = true;
            style_Heading1.IncludeNumber = false;
            style_Heading1.IncludePatterns = false;
            style_Heading1.IncludeProtection = false;

            IStyle style_Heading2 = workbook.Styles["Heading 2"];
            style_Heading2.HorizontalAlignment = HorizontalAlignment.General;
            style_Heading2.VerticalAlignment = VerticalAlignment.Bottom;
            style_Heading2.Font.ThemeFont = ThemeFont.Minor;
            style_Heading2.Font.Size = 14;
            style_Heading2.Font.ThemeColor = ThemeColor.Dark1;
            style_Heading2.Font.TintAndShade = 0.25;
            style_Heading2.Font.Bold = false;
            style_Heading2.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Dotted;
            style_Heading2.Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Light1;
            style_Heading2.Borders[BordersIndex.EdgeBottom].TintAndShade = -0.5;

            style_Heading2.IncludeAlignment = true;
            style_Heading2.IncludeFont = true;
            style_Heading2.IncludeBorder = true;
            style_Heading2.IncludeNumber = false;
            style_Heading2.IncludePatterns = false;
            style_Heading2.IncludeProtection = false;

            IStyle style_Heading3 = workbook.Styles["Heading 3"];
            style_Heading3.HorizontalAlignment = HorizontalAlignment.General;
            style_Heading3.VerticalAlignment = VerticalAlignment.Center;
            style_Heading3.Font.ThemeFont = ThemeFont.Minor;
            style_Heading3.Font.Size = 14;
            style_Heading3.Font.Bold = true;
            style_Heading3.Font.ThemeColor = ThemeColor.Dark1;
            style_Heading3.Font.TintAndShade = 0.25;
            style_Heading3.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;
            style_Heading3.Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Accent2;

            style_Heading3.IncludeAlignment = true;
            style_Heading3.IncludeFont = true;
            style_Heading3.IncludeBorder = true;
            style_Heading3.IncludeNumber = false;
            style_Heading3.IncludePatterns = false;
            style_Heading3.IncludeProtection = false;

            IStyle style_Normal = workbook.Styles["Normal"];
            style_Normal.NumberFormat = "General";
            style_Normal.HorizontalAlignment = HorizontalAlignment.General;
            style_Normal.VerticalAlignment = VerticalAlignment.Center;
            style_Normal.Font.ThemeFont = ThemeFont.Minor;
            style_Normal.Font.Size = 10;
            style_Normal.Font.ThemeColor = ThemeColor.Dark1;
            style_Normal.Font.TintAndShade = 0.25;

            style_Normal.IncludeAlignment = true;
            style_Normal.IncludeFont = true;
            style_Normal.IncludeBorder = true;
            style_Normal.IncludeNumber = true;
            style_Normal.IncludePatterns = true;
            style_Normal.IncludeProtection = true;

            //modify cell styles.
            worksheet.Range["B4"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            worksheet.Range["C4:F4"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            IRange range1 = worksheet.Range["E7:F7, E10:F10"];
            range1.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Dotted;
            range1.Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Light1;
            range1.Borders[BordersIndex.EdgeBottom].TintAndShade = -0.5;
            range1.Font.Bold = true;
            IRange range2 = worksheet.Range["E7, E10"];
            range2.Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            range2.Borders[BordersIndex.EdgeRight].ThemeColor = ThemeColor.Light1;
            range2.Borders[BordersIndex.EdgeRight].TintAndShade = -0.5;
            #endregion

            #region Chart
            //create a new chart.
            IShape shape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 8.99984251968504, 268.5, 627.750157480315, 184.5);

            //create series for chart.
            ISeries series_systolic = shape.Chart.SeriesCollection.NewSeries();
            ISeries series_diatolic = shape.Chart.SeriesCollection.NewSeries();
            ISeries series_HeartRate = shape.Chart.SeriesCollection.NewSeries();

            //set series formulas.
            series_systolic.Formula = "=SERIES('BLOOD PRESSURE DATA'!$E$36,'BLOOD PRESSURE DATA'!$C$37:$D$44,'BLOOD PRESSURE DATA'!$E$37:$E$44,1)";
            series_diatolic.Formula = "=SERIES('BLOOD PRESSURE DATA'!$F$36,'BLOOD PRESSURE DATA'!$C$37:$D$44,'BLOOD PRESSURE DATA'!$F$37:$F$44,2)";

            //set series plot on secondary axis, and change its chart type.
            series_HeartRate.AxisGroup = AxisGroup.Secondary;
            series_HeartRate.ChartType = ChartType.Line;
            series_HeartRate.Formula = "=SERIES('BLOOD PRESSURE DATA'!$G$36,,'BLOOD PRESSURE DATA'!$G$37:$G$44,3)";

            //set series fill to gradient fill.
            series_systolic.Format.Fill.TwoColorGradient(GradientStyle.Horizontal, 1);
            series_systolic.Format.Fill.GradientAngle = 270;
            series_systolic.Format.Fill.GradientStops[0].Color.RGB = Color.FromArgb(255, 172, 175);
            series_systolic.Format.Fill.GradientStops[1].Color.RGB = Color.FromArgb(255, 227, 228);
            series_systolic.Format.Fill.GradientStops.Insert(0xFEC6C8, 0.35);
            series_systolic.Format.Line.Color.ObjectThemeColor = ThemeColor.Accent6;

            series_diatolic.Format.Fill.TwoColorGradient(GradientStyle.Horizontal, 1);
            series_diatolic.Format.Fill.GradientAngle = 270;
            series_diatolic.Format.Fill.GradientStops[0].Color.RGB = Color.FromArgb(255, 192, 147);
            series_diatolic.Format.Fill.GradientStops[1].Color.RGB = Color.FromArgb(255, 227, 212);
            series_diatolic.Format.Fill.GradientStops.Insert(0xFFCBA9, 0.35);
            series_diatolic.Format.Line.Color.ObjectThemeColor = ThemeColor.Accent5;

            //set series gap width and overlap.
            shape.Chart.ColumnGroups[0].GapWidth = 150;
            shape.Chart.ColumnGroups[0].Overlap = 0;

            //set series line style.
            series_HeartRate.Format.Line.BeginArrowheadLength = ArrowheadLength.Medium;
            series_HeartRate.Format.Line.BeginArrowheadStyle = ArrowheadStyle.None;
            series_HeartRate.Format.Line.BeginArrowheadWidth = ArrowheadWidth.Medium;
            series_HeartRate.Format.Line.Color.ObjectThemeColor = ThemeColor.Accent4;
            series_HeartRate.Format.Line.Color.TintAndShade = 0;
            series_HeartRate.Format.Line.DashStyle = LineDashStyle.Solid;
            series_HeartRate.Format.Line.EndArrowheadLength = ArrowheadLength.Medium;
            series_HeartRate.Format.Line.EndArrowheadStyle = ArrowheadStyle.None;
            series_HeartRate.Format.Line.EndArrowheadWidth = ArrowheadWidth.Medium;
            series_HeartRate.Format.Line.Style = LineStyle.Single;
            series_HeartRate.Format.Line.Weight = 1.25;

            IAxis primary_axis = shape.Chart.Axes.Item(AxisType.Value, AxisGroup.Primary);
            primary_axis.HasTitle = true;
            primary_axis.AxisTitle.Text = "BLOOD PRESSURE";
            primary_axis.AxisTitle.IncludeInLayout = true;

            IAxis secondary_axis = shape.Chart.Axes.Item(AxisType.Value, AxisGroup.Secondary);
            secondary_axis.HasTitle = true;
            secondary_axis.AxisTitle.Text = "HEART RATE";
            secondary_axis.AxisTitle.IncludeInLayout = true;

            IAxis category_axis = shape.Chart.Axes.Item(AxisType.Category, AxisGroup.Primary);
            category_axis.HasTitle = true;
            category_axis.CategoryType = CategoryType.CategoryScale;
            category_axis.Format.Line.Color.ColorType = SolidColorType.None;

            shape.Chart.HasTitle = false;
            //set chart font style.
            shape.Chart.ChartArea.Font.Size = 9;
            shape.Chart.ChartArea.Font.Color.ObjectThemeColor = ThemeColor.Dark1;
            shape.Chart.ChartArea.Font.Color.Brightness = 0.5;
            #endregion


            #region Shape
            IShape shape1 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 402, 77.25, 234, 100);
            shape1.Fill.Solid();
            shape1.Fill.Color.ObjectThemeColor = ThemeColor.Accent1;
            shape1.Fill.Color.Brightness = 0.6;
            //set shape's border to no line.
            shape1.Line.Color.ColorType = SolidColorType.None;

            //set shape rich text.
            ITextRange shape1_p1 = shape1.TextFrame.TextRange.Paragraphs[0];
            shape1_p1.Text = "*";
            shape1_p1.Runs.Add(" Blood pressures may vary dependent on many");
            shape1_p1.Runs.Add(" factors.  Always consult with a physician about what is normal for you.  These numbers may vary slightly.");
            
            ITextRange shape1_p2 = shape1.TextFrame.TextRange.Paragraphs.Add("");
            ITextRange shape1_p3 = shape1.TextFrame.TextRange.Paragraphs.Add("Info from National Institute of Health:");
            ITextRange shape1_p4 = shape1.TextFrame.TextRange.Paragraphs.Add("http://www.nhlbi.nih.gov/health/health-topics/topics/hbp/");

            shape1.TextFrame.TextRange.Font.Size = 10;
            shape1.TextFrame.TextRange.Font.ThemeFont = ThemeFont.Minor;
            shape1.TextFrame.TextRange.Font.Color.ObjectThemeColor = ThemeColor.Dark1;
            shape1.TextFrame.TextRange.Font.Color.Brightness = 0.25;
            shape1_p3.Runs[0].Font.Bold = true;

            IShape shape2 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 421.5, 546.75, 198, 50);
            shape2.Fill.Solid();
            shape2.Fill.Color.ObjectThemeColor = ThemeColor.Accent3;
            shape2.Fill.Color.Brightness = 0.6;
            //set shape's border to no line.
            shape2.Line.Color.ColorType = SolidColorType.None;

            ITextRange shape2_p1 = shape2.TextFrame.TextRange.Paragraphs[0];
            shape2_p1.Text = "NOTE:";
            shape2_p1.Runs.Add(" Any blood pressure readings over the indicated numbers (\"CALL PHYSICIAN IF ABOVE\") will be");
            shape2_p1.Runs.Add(" highlighted.");

            shape2.TextFrame.TextRange.Font.Size = 10;
            shape2.TextFrame.TextRange.Font.ThemeFont = ThemeFont.Minor;
            shape2.TextFrame.TextRange.Font.Color.ObjectThemeColor = ThemeColor.Dark1;
            shape2.TextFrame.TextRange.Font.Color.Brightness = 0.25;
            shape2_p1.Runs[0].Font.Bold = true;
            #endregion

            #region DefinedName
            //create defined names for workbook.
            workbook.Names.Add("MaxDiastolic", "='BLOOD PRESSURE DATA'!$F$10");
            workbook.Names.Add("MaxSystolic", "='BLOOD PRESSURE DATA'!$E$10");
            #endregion

            #region ConditionalFormat
            //create conditional format for ranges.
            IFormatCondition condition1 = worksheet.Range["E37:E44"].FormatConditions.Add(FormatConditionType.Expression, formula1: "=E37>MaxSystolic") as IFormatCondition;
            IFormatCondition condition2 = worksheet.Range["F37:F44"].FormatConditions.Add(FormatConditionType.Expression, formula1: "=F37>MaxDiastolic") as IFormatCondition;
            condition1.Interior.Color = Color.Red;
            condition2.Interior.Color = Color.Red;
            #endregion
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
