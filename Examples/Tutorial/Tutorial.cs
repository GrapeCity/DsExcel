using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using GrapeCity.Documents.Excel.Drawing;

namespace GrapeCity.Documents.Excel.Examples
{
    public class Tutorial : ExampleBase
    {
        public override bool CanDownload
        {
            get
            {
                return false;
            }
        }

        public override bool ShowCode
        {
            get
            {
                return false;
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override string GetDescriptionByCulture(string culture)
        {
                string resource = "GrapeCity.Documents.Excel.Examples.Tutorial.Tutorial_" + culture + ".md";
                var assembly = this.GetType().GetTypeInfo().Assembly;
                using (Stream stream = assembly.GetManifestResourceStream(resource))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
        }

        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //-----------------------------Set Value------------------------------
            worksheet.Range["B3:C7"].Value = new object[,]
            {
                { "ITEM", "AMOUNT" },
                { "Income 1", 2500 },
                { "Income 2", 1000 },
                { "Income 3", 250 },
                { "Other", 250 },
            };
            worksheet.Range["B10:C23"].Value = new object[,]
            {
                { "ITEM", "AMOUNT" },
                { "Rent/mortgage", 800 },
                { "Electric", 120 },
                { "Gas", 50 },
                { "Cell phone", 45 },
                { "Groceries", 500 },
                { "Car payment", 273 },
                { "Auto expenses", 120 },
                { "Student loans", 50 },
                { "Credit cards", 100 },
                { "Auto Insurance", 78 },
                { "Personal care", 50 },
                { "Entertainment", 100 },
                { "Miscellaneous", 50 },
            };

            worksheet.Range["B2:C2"].Merge();
            worksheet.Range["B2"].Value = "MONTHLY INCOME";
            worksheet.Range["B9:C9"].Merge();
            worksheet.Range["B9"].Value = "MONTHLY EXPENSES";
            worksheet.Range["E2:G2"].Merge();
            worksheet.Range["E2"].Value = "PERCENTAGE OF INCOME SPENT";
            worksheet.Range["E5:G5"].Merge();
            worksheet.Range["E5"].Value = "SUMMARY";
            worksheet.Range["E3:F3"].Merge();
            worksheet.Range["E9"].Value = "BALANCE";
            worksheet.Range["E6"].Value = "Total Monthly Income";
            worksheet.Range["E7"].Value = "Total Monthly Expenses";


            //--------------------------------Set Height & Width--------------------------------
            worksheet.StandardHeight = 26.25;
            worksheet.StandardWidth = 8.43;

            worksheet.Range["2:24"].RowHeight = 27;
            worksheet.Range["A:A"].ColumnWidth = 2.855;
            worksheet.Range["B:B"].ColumnWidth = 33.285;
            worksheet.Range["C:C"].ColumnWidth = 25.57;
            worksheet.Range["D:D"].ColumnWidth = 1;
            worksheet.Range["E:F"].ColumnWidth = 25.57;
            worksheet.Range["G:G"].ColumnWidth = 14.285;


            //------------------------------Set Table--------------------------------------
            ITable incomeTable = worksheet.Tables.Add(worksheet.Range["B3:C7"], true);
            incomeTable.Name = "tblIncome";
            incomeTable.TableStyle = workbook.TableStyles["TableStyleMedium4"];
            ITable expensesTable = worksheet.Tables.Add(worksheet.Range["B10:C23"], true);
            expensesTable.Name = "tblExpenses";
            expensesTable.TableStyle = workbook.TableStyles["TableStyleMedium4"];

            //------------------------------Set Formulas-----------------------------------
            worksheet.Names.Add("TotalMonthlyIncome", "=SUM(tblIncome[AMOUNT])");
            worksheet.Names.Add("TotalMonthlyExpenses", "=SUM(tblExpenses[AMOUNT])");
            worksheet.Range["E3"].Formula = "=TotalMonthlyExpenses";
            worksheet.Range["G3"].Formula = "=TotalMonthlyExpenses/TotalMonthlyIncome";
            worksheet.Range["G6"].Formula = "=TotalMonthlyIncome";
            worksheet.Range["G7"].Formula = "=TotalMonthlyExpenses";
            worksheet.Range["G9"].Formula = "=TotalMonthlyIncome-TotalMonthlyExpenses";

            //----------------------------Set Styles-------------------------
            IStyle currencyStyle = workbook.Styles["Currency"];
            currencyStyle.IncludeAlignment = true;
            currencyStyle.HorizontalAlignment = HorizontalAlignment.Left;
            currencyStyle.VerticalAlignment = VerticalAlignment.Bottom;
            currencyStyle.NumberFormat = "$#,##0.00";

            IStyle heading1Style = workbook.Styles["Heading 1"];
            heading1Style.IncludeAlignment = true;
            heading1Style.HorizontalAlignment = HorizontalAlignment.Center;
            heading1Style.VerticalAlignment = VerticalAlignment.Center;
            heading1Style.Font.Name = "Century Gothic";
            heading1Style.Font.Bold = true;
            heading1Style.Font.Size = 11;
            heading1Style.Font.Color = Color.White;
            heading1Style.IncludeBorder = false;
            heading1Style.IncludePatterns = true;
            heading1Style.Interior.Color = Color.FromArgb(32, 61, 64);

            IStyle percentStyle = workbook.Styles["Percent"];
            percentStyle.IncludeAlignment = true;
            percentStyle.HorizontalAlignment = HorizontalAlignment.Center;
            percentStyle.IncludeFont = true;
            percentStyle.Font.Color = Color.FromArgb(32, 61, 64);
            percentStyle.Font.Name = "Century Gothic";
            percentStyle.Font.Bold = true;
            percentStyle.Font.Size = 14;

            worksheet.SheetView.DisplayGridlines = false;
            worksheet.Range["C4:C7, C11:C23, G6:G7, G9"].Style = currencyStyle;
            worksheet.Range["B2, B9, E2, E5"].Style = heading1Style;
            worksheet.Range["G3"].Style = percentStyle;

            worksheet.Range["E6:G6"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Medium;
            worksheet.Range["E6:G6"].Borders[BordersIndex.EdgeBottom].Color = Color.FromArgb(32, 61, 64);
            worksheet.Range["E7:G7"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Medium;
            worksheet.Range["E7:G7"].Borders[BordersIndex.EdgeBottom].Color = Color.FromArgb(32, 61, 64);

            worksheet.Range["E9:G9"].Interior.Color = Color.FromArgb(32, 61, 64);
            worksheet.Range["E9:G9"].HorizontalAlignment = HorizontalAlignment.Left;
            worksheet.Range["E9:G9"].VerticalAlignment = VerticalAlignment.Center;
            worksheet.Range["E9:G9"].Font.Name = "Century Gothic";
            worksheet.Range["E9:G9"].Font.Bold = true;
            worksheet.Range["E9:G9"].Font.Size = 11;
            worksheet.Range["E9:G9"].Font.Color = Color.White;
            worksheet.Range["E3:F3"].Borders.Color = Color.FromArgb(32, 61, 64);

            //----------------------------Set Conditional Format-------------------------
            IDataBar dataBar = worksheet.Range["E3"].FormatConditions.AddDatabar();
            dataBar.MinPoint.Type = ConditionValueTypes.Number;
            dataBar.MinPoint.Value = 1;
            dataBar.MaxPoint.Type = ConditionValueTypes.Number;
            dataBar.MaxPoint.Value = "=TotalMonthlyIncome";
            dataBar.BarFillType = DataBarFillType.Gradient;
            dataBar.BarColor.Color = Color.Red;
            dataBar.ShowValue = false;
            
            //--------------------------------Set Shape--------------------------------
            IShape shape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 339, 247, 316.5, 346);
            shape.Chart.ChartArea.Format.Line.Transparency = 1;
            shape.Chart.ColumnGroups[0].Overlap = 0;
            shape.Chart.ColumnGroups[0].GapWidth = 37;

            IAxis category_axis = shape.Chart.Axes.Item(AxisType.Category);
            category_axis.Format.Line.Color.RGB = Color.Black;
            category_axis.TickLabels.Font.Size = 11;
            category_axis.TickLabels.Font.Color.RGB = Color.Black;

            IAxis series_axis = shape.Chart.Axes.Item(AxisType.Value);
            series_axis.Format.Line.Weight = 1;
            series_axis.Format.Line.Color.RGB = Color.Black;
            series_axis.TickLabels.NumberFormat = "$###0";
            series_axis.TickLabels.Font.Size = 11;
            series_axis.TickLabels.Font.Color.RGB = Color.Black;

            ISeries chartSeries = shape.Chart.SeriesCollection.NewSeries();
            chartSeries.Formula = "=SERIES(\"Simple Budget\",{\"Income\",\"Expenses\"},'Sheet1'!$G$6:$G$7,1)";
            chartSeries.Points[0].Format.Fill.Color.RGB = Color.FromArgb(176, 21, 19);
            chartSeries.Points[1].Format.Fill.Color.RGB = Color.FromArgb(234, 99, 18);
            chartSeries.DataLabels.Font.Size = 11;
            chartSeries.DataLabels.Font.Color.RGB = Color.Black;
            chartSeries.DataLabels.ShowValue = true;
            chartSeries.DataLabels.Position = DataLabelPosition.OutsideEnd;
        }

    }
}
