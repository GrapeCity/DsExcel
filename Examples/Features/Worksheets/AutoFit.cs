using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class AutoFit : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Auto fit column width of range 'A1'
            worksheet.Range["A1"].Value = "Grapecity Documents for Excel";
            worksheet.Range["A1"].Columns.AutoFit();

            //Auto fit row height of range 'B2'
            worksheet.Range["B2"].Value = "Grapecity";
            worksheet.Range["B2"].Font.Size = 20;
            worksheet.Range["B2"].Rows.AutoFit();

            //Auto fit column width and row height of range 'C3'
            worksheet.Range["C3"].Value = "Grapecity Documents for Excel";
            worksheet.Range["C3"].Font.Size = 32;
            worksheet.Range["C3"].AutoFit();
        }
    }
}
