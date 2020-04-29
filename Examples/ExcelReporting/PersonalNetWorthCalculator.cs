using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelReporting
{
    public class PersonalNetWorthCalculator : ExampleBase
    {
        protected override void BeforeExecute(Workbook workbook, string[] userAgents)
        {
            if (AgentIsMac(userAgents))
            {
                Themes themes = new Themes();
                ITheme theme = themes.Add("testTheme", Themes.OfficeTheme);
                theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Trebuchet MS";
                workbook.Theme = theme;
                var style_Normal = workbook.Styles["Normal"];
                style_Normal.Font.ThemeFont = ThemeFont.Minor;
            }
        }
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //-------------------------Set RowHeight & Width-----------------------------------
            worksheet.StandardHeight = 30;
            worksheet.StandardWidth = 8.43;

            worksheet.Range["1:1"].RowHeight = 278.25;
            worksheet.Range["2:4"].RowHeight = 30.25;
            worksheet.Range["8:8"].RowHeight = 55.5;
            worksheet.Range["9:30"].RowHeight = 30.25;
            worksheet.Range["33:33"].RowHeight = 55.5;
            worksheet.Range["34:44"].RowHeight = 43.5;
            worksheet.Range["A:A"].ColumnWidth = 2.777;
            worksheet.Range["B:B"].ColumnWidth = 32.887;
            worksheet.Range["C:C"].ColumnWidth = 24.219;
            worksheet.Range["D:D"].ColumnWidth = 10.109;
            worksheet.Range["E:E"].ColumnWidth = 61.332;
            worksheet.Range["F:F"].ColumnWidth = 2.777;


            //-------------------------Set Table Value & Formulas-------------------------------
            ITable assetsTable = worksheet.Tables.Add(worksheet.Range["B9:D30"], true);
            assetsTable.Name = "Assets";
            worksheet.Range["B8"].Value = "Assets";
            worksheet.Range["B9:D30"].Value = new object[,]
            {
                { "Category", "Item", "Value" },
                { "Real Estate", "Home", 560000 },
                { "Real Estate", "Other", 255000 },
                { "Investments", "Retirement accounts", 98000 },
                { "Investments", "Stocks", 53000 },
                { "Investments", "Bonds", 25000 },
                { "Investments", "Mutual funds", 33000 },
                { "Investments", "CDs", 74000 },
                { "Investments", "Bullion", 20000 },
                { "Investments", "Trust funds", 250000 },
                { "Investments", "Health savings account", 18000 },
                { "Investments", "Face value of life insurance policy", 85000 },
                { "Investments", "Other", 20000 },
                { "Cash", "Checking accounts", 14500 },
                { "Cash", "Savings accounts", 5000 },
                { "Cash", "Other", 2000 },
                { "Personal Property", "Cars", 55000 },
                { "Personal Property", "Other vehicles", 85000 },
                { "Personal Property", "Furnishings", 100000 },
                { "Personal Property", "Collectibles", 50000 },
                { "Personal Property", "Jewelry", 60000 },
                { "Personal Property", "Other luxury goods", 40000 },
            };

            ITable debtsTable = worksheet.Tables.Add(worksheet.Range["B34:C44"], true);
            debtsTable.Name = "Debts";
            worksheet.Range["B33"].Value = "Debts";
            worksheet.Range["B34:C44"].Value = new object[,]
            {
                { "Category", "Value" },
                { "Mortgages", 400000 },
                { "Home equity loans", 50000 },
                { "Car loans", 30000 },
                { "Personal loans", 0 },
                { "Credit cards", 0 },
                { "Student loans", 10000 },
                { "Loans against investments", 20000 },
                { "Life insurance loans", 5000 },
                { "Other installment loans", 10000 },
                { "Other debts", 50000 },
            };

            worksheet.Range["B1:C1"].Merge();
            worksheet.Range["B1"].Value = "Personal\r\nNet\r\nWorth";
            worksheet.Range["B2"].Formula = "=\"Total \"&TotalAssetsLabel";
            worksheet.Range["B3"].Formula = "=\"Total \"&TotalDebtsLabel";
            worksheet.Range["B4"].Formula = "=NetWorthLabel";
            worksheet.Range["C2"].Formula = "=TotalAssets";
            worksheet.Range["C3"].Formula = "=TotalDebts";
            worksheet.Range["C4"].Formula = "=NetWorth";

            worksheet.Names.Add("TotalAssets", "=SUM(Assets[Value])");
            worksheet.Names.Add("TotalDebts", "=SUM(Debts[Value])");
            worksheet.Names.Add("NetWorth", "=TotalAssets-TotalDebts");
            worksheet.Names.Add("TotalAssetsLabel", "=Sheet1!$B$8");
            worksheet.Names.Add("TotalDebtsLabel", "=Sheet1!$B$33");
            worksheet.Names.Add("NetWorthLabel", "=\"Net Worth\"");


            //---------------------------Set Table Style---------------------------
            ITableStyle assetsTableStyle = workbook.TableStyles.Add("Assets");
            workbook.DefaultTableStyle = "Assets";
            assetsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Font.Color = Color.FromArgb(64, 64, 64);
            assetsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders.Color = Color.FromArgb(128, 128, 128);
            assetsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Dotted;
            assetsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            assetsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.None;
            assetsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.None;
            assetsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.None;
            assetsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.None;

            assetsTableStyle.TableStyleElements[TableStyleElementType.SecondRowStripe].Interior.Color = Color.White;
            assetsTableStyle.TableStyleElements[TableStyleElementType.SecondRowStripe].StripeSize = 1;

            assetsTableStyle.TableStyleElements[TableStyleElementType.LastColumn].Font.Bold = true;
            assetsTableStyle.TableStyleElements[TableStyleElementType.LastColumn].Font.Color = Color.FromArgb(61, 125, 137);
            assetsTableStyle.TableStyleElements[TableStyleElementType.LastColumn].Interior.Color = Color.White;

            assetsTableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Interior.Color = Color.FromArgb(61, 125, 137);


            ITableStyle debtsTableStyle = workbook.TableStyles.Add("Debts");
            debtsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Font.Color = Color.FromArgb(64, 64, 64);
            debtsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders.Color = Color.FromArgb(128, 128, 128);
            debtsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Dotted;
            debtsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            debtsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.None;
            debtsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.None;
            debtsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.None;
            debtsTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.None;

            debtsTableStyle.TableStyleElements[TableStyleElementType.SecondRowStripe].Interior.Color = Color.White;
            debtsTableStyle.TableStyleElements[TableStyleElementType.SecondRowStripe].StripeSize = 1;

            debtsTableStyle.TableStyleElements[TableStyleElementType.LastColumn].Font.Bold = true;
            debtsTableStyle.TableStyleElements[TableStyleElementType.LastColumn].Font.Color = Color.FromArgb(146, 75, 12);
            debtsTableStyle.TableStyleElements[TableStyleElementType.LastColumn].Interior.Color = Color.White;

            debtsTableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Interior.Color = Color.FromArgb(218, 113, 18);


            //----------------------------Set Named Styles-------------------------
            IStyle normalStyle = workbook.Styles["Normal"];
            normalStyle.Font.Name = "Century Gothic";
            normalStyle.Font.Size = 12;
            normalStyle.Font.Color = Color.FromArgb(64, 64, 64);
            normalStyle.Interior.Color = Color.FromArgb(243, 243, 236);
            normalStyle.Interior.PatternColor = Color.FromArgb(243, 243, 236);
            normalStyle.HorizontalAlignment = HorizontalAlignment.Left;
            normalStyle.IndentLevel = 1;
            normalStyle.VerticalAlignment = VerticalAlignment.Center;
            normalStyle.WrapText = true;

            IStyle titleStyle = workbook.Styles["Title"];
            titleStyle.IncludeAlignment = true;
            titleStyle.VerticalAlignment = VerticalAlignment.Center;
            titleStyle.WrapText = true;
            titleStyle.Font.Name = "Century Gothic";
            titleStyle.Font.Size = 66;
            titleStyle.Font.Color = Color.FromArgb(64, 64, 64);
            titleStyle.IncludePatterns = true;
            titleStyle.Interior.Color = Color.FromArgb(243, 243, 236);

            IStyle heading1Style = workbook.Styles["Heading 1"];
            heading1Style.IncludeAlignment = true;
            heading1Style.HorizontalAlignment = HorizontalAlignment.Left;
            heading1Style.IndentLevel = 4;
            heading1Style.VerticalAlignment = VerticalAlignment.Center;
            heading1Style.Font.Name = "Century Gothic";
            heading1Style.Font.Bold = false;
            heading1Style.Font.Size = 16;
            heading1Style.Font.Color = Color.FromArgb(64, 64, 64);
            heading1Style.IncludeBorder = false;
            heading1Style.IncludePatterns = true;
            heading1Style.Interior.Color = Color.FromArgb(243, 243, 236);

            IStyle heading2Style = workbook.Styles["Heading 2"];
            heading2Style.IncludeNumber = true;
            heading2Style.NumberFormat = "$#,##0";
            heading2Style.IncludeAlignment = true;
            heading2Style.HorizontalAlignment = HorizontalAlignment.Right;
            heading2Style.IndentLevel = 2;
            heading2Style.VerticalAlignment = VerticalAlignment.Center;
            heading2Style.Font.Name = "Century Gothic";
            heading2Style.Font.Size = 16;
            heading2Style.Font.Color = Color.FromArgb(64, 64, 64);
            heading2Style.IncludeBorder = false;
            heading2Style.IncludePatterns = true;
            heading2Style.Interior.Color = Color.FromArgb(243, 243, 236);

            IStyle heading3Style = workbook.Styles["Heading 3"];
            heading3Style.IncludeAlignment = true;
            heading3Style.HorizontalAlignment = HorizontalAlignment.Left;
            heading3Style.VerticalAlignment = VerticalAlignment.Bottom;
            heading3Style.IncludeBorder = false;
            heading3Style.Font.Name = "Century Gothic";
            heading3Style.Font.Bold = false;
            heading3Style.Font.Size = 27;
            heading3Style.Font.Color = Color.FromArgb(64, 64, 64);
            heading3Style.IncludePatterns = true;
            heading3Style.Interior.Color = Color.FromArgb(243, 243, 236);

            IStyle heading4Style = workbook.Styles["Heading 4"];
            heading4Style.Font.Name = "Century Gothic";
            heading4Style.Font.Size = 16;
            heading4Style.Font.Color = Color.White;
            heading4Style.Font.Bold = false;

            IStyle currencyStyle = workbook.Styles["Currency"];
            currencyStyle.NumberFormat = "$#,##0";
            currencyStyle.IncludeAlignment = true;
            currencyStyle.HorizontalAlignment = HorizontalAlignment.Right;
            currencyStyle.IndentLevel = 1;
            currencyStyle.VerticalAlignment = VerticalAlignment.Center;
            currencyStyle.IncludeFont = true;
            currencyStyle.Font.Bold = true;
            currencyStyle.Font.Name = "Century Gothic";
            currencyStyle.Font.Size = 12;


            //----------------------------------Use Style---------------------------
            assetsTable.TableStyle = assetsTableStyle;
            debtsTable.TableStyle = debtsTableStyle;

            worksheet.SheetView.DisplayGridlines = false;
            worksheet.Range["B2:B4"].Style = heading1Style;
            worksheet.Range["C2:C4"].Style = heading2Style;
            worksheet.Range["B9:D9"].Style = heading4Style;
            worksheet.Range["D10:D30"].Style = currencyStyle;
            worksheet.Range["D10:D30"].Font.Color = Color.FromArgb(61, 125, 137);

            worksheet.Range["B34:C34"].Style = heading4Style;
            worksheet.Range["C35:C44"].Style = currencyStyle;
            worksheet.Range["C35:C44"].Font.Color = Color.FromArgb(218, 113, 18);
            worksheet.Range["B1"].Style = titleStyle;
            worksheet.Range["B8"].Style = heading3Style;
            worksheet.Range["B33"].Style = heading3Style;

            worksheet.Range["B3:C3"].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Hair;
            worksheet.Range["B3:C3"].Borders[BordersIndex.EdgeTop].Color = Color.FromArgb(128, 128, 128);
            worksheet.Range["B3:C3"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Hair;
            worksheet.Range["B3:C3"].Borders[BordersIndex.EdgeBottom].Color = Color.FromArgb(128, 128, 128);


            //--------------------------------Add Shape--------------------------------
            IShape recShape1 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 17.81, 282.75, 20.963, 21.75);
            recShape1.Line.Color.ColorType = SolidColorType.None;
            recShape1.Fill.Color.RGB= Color.FromArgb(60, 126, 138);
            IShape recShape2 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 17.81, 312.75, 20.963, 21.75);
            recShape2.Line.Color.ColorType = SolidColorType.None;
            recShape2.Fill.Color.RGB = Color.FromArgb(218, 118, 13);
            IShape recShape3 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 17.81, 342.75, 20.963, 21.75);
            recShape3.Line.Color.ColorType = SolidColorType.None;
            recShape3.Fill.Color.RGB = Color.FromArgb(84, 138, 57);

            IShape pieShape = worksheet.Shapes.AddChart(ChartType.Pie, 442.5, 26.25, 346, 350.25);
            pieShape.Chart.HasLegend = false;
            pieShape.Chart.HasTitle = false;
            pieShape.Chart.ChartGroups[0].FirstSliceAngle = 180;
            pieShape.Placement = Placement.Move;

            IChartArea chartArea = pieShape.Chart.ChartArea;
            chartArea.Format.Fill.Transparency = 1;
            chartArea.Format.Line.Transparency = 1;

            ISeries chartSeries = pieShape.Chart.SeriesCollection.NewSeries();
            chartSeries.Formula = "=SERIES('Sheet1'!$B$2:$B$4,,'Sheet1'!$C$2:$C$4,1)";

            chartSeries.HasDataLabels = true;
            chartSeries.DataLabels.Font.Name = "Century Gothic";
            chartSeries.DataLabels.Font.Size = 20;
            chartSeries.DataLabels.Font.Bold = true;
            chartSeries.DataLabels.Font.Color.RGB = Color.White;
            chartSeries.DataLabels.ShowValue = false;
            chartSeries.DataLabels.ShowPercentage = true;
            chartSeries.DataLabels.Position = DataLabelPosition.Center;

            chartSeries.Points[0].Format.Fill.Color.RGB = Color.FromArgb(60, 126, 138);
            chartSeries.Points[1].Format.Fill.Color.RGB = Color.FromArgb(218, 118, 13);
            chartSeries.Points[2].Format.Fill.Color.RGB = Color.FromArgb(84, 138, 57);
            chartSeries.Explosion = 1;

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
