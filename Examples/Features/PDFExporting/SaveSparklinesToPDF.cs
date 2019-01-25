using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SaveSparklinesToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] data = new object[,]
             {
                { "Customer", "0-30 Days", "30-60 Days", "60-90 Days", ">90 Days"},
                { "Customer A",1200.15, 1916.18, 1105.23, 1806.53},
                { "Customer B",896.23, 1005.53, 1800.56, 1150.49,},
                { "Customer C", 827.63, 1009.23, 1869.23, 1002.56,}
             };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["B2:E5"].Value = data;
            worksheet.Range["B:F"].ColumnWidth = 15;
            worksheet.Range["B:E"].HorizontalAlignment = HorizontalAlignment.Center;
            ITable table = worksheet.Tables.Add(worksheet.Range["B2:F5"], true);
            table.TableStyle = workbook.TableStyles["TableStyleMedium3"];
            table.Columns[4].Name = "Sparklines";

            //create a new group of sparklines.
            worksheet.Range["F3"].SparklineGroups.Add(SparkType.Line, "C3:E3");
            worksheet.Range["F4"].SparklineGroups.Add(SparkType.Column, "C4:E4");
            worksheet.Range["F5"].SparklineGroups.Add(SparkType.ColumnStacked100, "C5:E5");
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
