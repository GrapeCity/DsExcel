using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelReporting
{
    public class BidTracker : ExampleBase
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

            //***********************Set RowHeight & ColumnWidth***************
            worksheet.StandardHeight = 30;
            worksheet.Range["1:1"].RowHeight = 57.75;
            worksheet.Range["2:9"].RowHeight = 30.25;
            worksheet.Range["A:A"].ColumnWidth = 2.71;
            worksheet.Range["B:B"].ColumnWidth = 11.71;
            worksheet.Range["C:C"].ColumnWidth = 28;
            worksheet.Range["D:D"].ColumnWidth = 22.425;
            worksheet.Range["E:E"].ColumnWidth = 16.71;
            worksheet.Range["F:F"].ColumnWidth = 28;
            worksheet.Range["G:H"].ColumnWidth = 16.71;
            worksheet.Range["I:I"].ColumnWidth = 2.71;



            //**************************Set Table Value & Formulas*********************
            ITable table = worksheet.Tables.Add(worksheet.Range["B2:H9"], true);
            worksheet.Range["B2:H9"].Value = new object[,]
            {
                { "BID #", "DESCRIPTION", "DATE RECEIVED", "AMOUNT", "PERCENT COMPLETE", "DEADLINE", "DAYS LEFT" },
                { 1, "Bid number 1", null, 2000, 0.5, null, null },
                { 2, "Bid number 2", null, 3500, 0.25, null, null },
                { 3, "Bid number 3", null, 5000, 0.3, null, null },
                { 4, "Bid number 4", null, 4000, 0.2, null, null },
                { 5, "Bid number 5", null, 4000, 0.75, null , null },
                { 6, "Bid number 6", null, 1500, 0.45, null , null },
                { 7, "Bid number 7", null, 5000, 0.65, null, null },
            };
            worksheet.Range["B1"].Value = "Bid Details";
            worksheet.Range["D3"].Formula = "=TODAY()-10";
            worksheet.Range["D4:D5"].Formula = "=TODAY()-20";
            worksheet.Range["D6"].Formula = "=TODAY()-10";
            worksheet.Range["D7"].Formula = "=TODAY()-28";
            worksheet.Range["D8"].Formula = "=TODAY()-17";
            worksheet.Range["D9"].Formula = "=TODAY()-15";
            worksheet.Range["G3:G9"].Formula = "=[@[DATE RECEIVED]]+30";
            worksheet.Range["H3:H9"].Formula = "=[@DEADLINE]-TODAY()";



            //****************************Set Table Style********************************
            ITableStyle tableStyle = workbook.TableStyles.Add("Bid Tracker");
            workbook.DefaultTableStyle = "Bid Tracker";

            //Set WholeTable element style.
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Font.Color = Color.FromArgb(89, 89, 89);
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders.Color = Color.FromArgb(89, 89, 89);
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;

            //Set HeaderRow element style.
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders.Color = Color.FromArgb(89, 89, 89);
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Interior.Color = Color.FromArgb(131, 95, 1);
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Interior.PatternColor = Color.FromArgb(254, 184, 10);


            //Set TotalRow element style.
            tableStyle.TableStyleElements[TableStyleElementType.TotalRow].Borders.Color = Color.White;
            tableStyle.TableStyleElements[TableStyleElementType.TotalRow].Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.TotalRow].Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.TotalRow].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.TotalRow].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.TotalRow].Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.TotalRow].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.TotalRow].Interior.Color = Color.FromArgb(131, 95, 1);



            //***********************************Set Named Styles*****************************
            IStyle titleStyle = workbook.Styles["Title"];
            titleStyle.Font.Name = "Trebuchet MS";
            titleStyle.Font.Size = 36;
            titleStyle.Font.Color = Color.FromArgb(56, 145, 167);
            titleStyle.IncludeAlignment = true;
            titleStyle.VerticalAlignment = VerticalAlignment.Center;


            IStyle heading1Style = workbook.Styles["Heading 1"];
            heading1Style.IncludeAlignment = true;
            heading1Style.HorizontalAlignment = HorizontalAlignment.Left;
            heading1Style.IndentLevel = 1;
            heading1Style.VerticalAlignment = VerticalAlignment.Bottom;
            heading1Style.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            heading1Style.Font.Size = 14;
            heading1Style.Font.Color = Color.White;
            heading1Style.Font.Bold = false;
            heading1Style.IncludePatterns = true;
            heading1Style.Interior.Color = Color.FromArgb(131, 95, 1);
            heading1Style.Font.Name = "Trebuchet MS";


            IStyle dateStyle = workbook.Styles.Add("Date");
            dateStyle.IncludeNumber = true;
            dateStyle.NumberFormat = "m/d/yyyy";
            dateStyle.IncludeAlignment = true;
            dateStyle.HorizontalAlignment = HorizontalAlignment.Left;
            dateStyle.IndentLevel = 1;
            dateStyle.VerticalAlignment = VerticalAlignment.Center;
            dateStyle.IncludeFont = false;
            dateStyle.IncludeBorder = false;
            dateStyle.IncludePatterns = false;
            dateStyle.Font.Name = "Trebuchet MS";


            IStyle commaStyle = workbook.Styles["Comma"];
            commaStyle.IncludeNumber = true;
            commaStyle.NumberFormat = "#,##0_);(#,##0)";
            commaStyle.IncludeAlignment = true;
            commaStyle.HorizontalAlignment = HorizontalAlignment.Left;
            commaStyle.IndentLevel = 1;
            commaStyle.VerticalAlignment = VerticalAlignment.Center;
            commaStyle.Font.Name = "Trebuchet MS";


            IStyle normalStyle = workbook.Styles["Normal"];
            normalStyle.HorizontalAlignment = HorizontalAlignment.Left;
            normalStyle.IndentLevel = 1;
            normalStyle.VerticalAlignment = VerticalAlignment.Center;
            normalStyle.WrapText = true;
            normalStyle.Font.Color = Color.FromArgb(89, 89, 89);


            IStyle currencyStyle = workbook.Styles["Currency"];
            currencyStyle.NumberFormat = "$#,##0.00";
            currencyStyle.IncludeAlignment = true;
            currencyStyle.HorizontalAlignment = HorizontalAlignment.Left;
            currencyStyle.IndentLevel = 1;
            currencyStyle.VerticalAlignment = VerticalAlignment.Center;
            currencyStyle.Font.Name = "Trebuchet MS";


            IStyle percentStyle = workbook.Styles["Percent"];
            percentStyle.IncludeAlignment = true;
            percentStyle.HorizontalAlignment = HorizontalAlignment.Right;
            percentStyle.VerticalAlignment = VerticalAlignment.Center;
            percentStyle.IncludeFont = true;
            percentStyle.Font.Name = "Trebuchet MS";
            percentStyle.Font.Size = 20;
            percentStyle.Font.Bold = true;
            percentStyle.Font.Color = Color.FromArgb(89, 89, 89);
            percentStyle.Font.Name = "Trebuchet MS";


            IStyle comma0Style = workbook.Styles["Comma [0]"];
            comma0Style.NumberFormat = "#,##0_);(#,##0)";
            comma0Style.IncludeAlignment = true;
            comma0Style.HorizontalAlignment = HorizontalAlignment.Right;
            comma0Style.IndentLevel = 3;
            comma0Style.VerticalAlignment = VerticalAlignment.Center;
            percentStyle.Font.Name = "Trebuchet MS";



            //************************************Add Conditional Formatting****************
            IDataBar dataBar = worksheet.Range["F3:F9"].FormatConditions.AddDatabar();
            dataBar.MinPoint.Type = ConditionValueTypes.Number;
            dataBar.MinPoint.Value = 1;
            dataBar.MaxPoint.Type = ConditionValueTypes.Number;
            dataBar.MaxPoint.Value = 0;

            dataBar.BarFillType = DataBarFillType.Gradient;
            dataBar.BarColor.Color = Color.FromArgb(126, 194, 211);
            dataBar.Direction = DataBarDirection.Context;

            dataBar.AxisColor.Color = Color.Black;
            dataBar.AxisPosition = DataBarAxisPosition.Automatic;

            dataBar.NegativeBarFormat.ColorType = DataBarNegativeColorType.Color;
            dataBar.NegativeBarFormat.Color.Color = Color.Red;
            dataBar.ShowValue = true;



            //****************************************Use NamedStyle**************************
            worksheet.SheetView.DisplayGridlines = false;
            table.TableStyle = tableStyle;
            worksheet.Range["B1"].Style = titleStyle;
            worksheet.Range["B1"].WrapText = false;
            worksheet.Range["B2:H2"].Style = heading1Style;
            worksheet.Range["B3:B9"].Style = commaStyle;
            worksheet.Range["C3:C9"].Style = normalStyle;
            worksheet.Range["D3:D9"].Style = dateStyle;
            worksheet.Range["E3:E9"].Style = currencyStyle;
            worksheet.Range["F3:F9"].Style = percentStyle;
            worksheet.Range["G3:G9"].Style = dateStyle;
            worksheet.Range["H3:H9"].Style = comma0Style;

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
