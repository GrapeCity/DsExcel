using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SaveTable : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            //Add Table
            ITable table = sheet.Tables.Add(sheet.Range["B5:G16"], true);
            table.ShowTotals = true;

            //Set values
            int[] data = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 };
            sheet.Range["C6:C16"].Value = data;
            sheet.Range["D6:D16"].Value = data;

            //Set total functions
            table.Columns[1].TotalsCalculation = TotalsCalculation.Average;
            table.Columns[2].TotalsCalculation = TotalsCalculation.Sum;

            //Create custom table style
            ITableStyle customTableStyle = workbook.TableStyles["TableStyleMedium10"].Duplicate();

            var wholeTableStyle = customTableStyle.TableStyleElements[TableStyleElementType.WholeTable];
            wholeTableStyle.Font.Italic = true;
            wholeTableStyle.Borders[BordersIndex.EdgeTop].ThemeColor = ThemeColor.Accent1;
            wholeTableStyle.Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thick;
            wholeTableStyle.Borders[BordersIndex.EdgeRight].ThemeColor = ThemeColor.Accent1;
            wholeTableStyle.Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thick;
            wholeTableStyle.Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Accent1;
            wholeTableStyle.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;
            wholeTableStyle.Borders[BordersIndex.EdgeLeft].ThemeColor = ThemeColor.Accent1;
            wholeTableStyle.Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thick;

            var firstRowStripStyle = customTableStyle.TableStyleElements[TableStyleElementType.FirstRowStripe];
            firstRowStripStyle.Font.Bold = true;

            //Apply custom style to table
            table.TableStyle = customTableStyle;
        }

        public override bool SavePdf
        {
            get
            {
                return true;
            }
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
