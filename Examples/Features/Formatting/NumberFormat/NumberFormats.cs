using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.NumberFormat
{
    public class NumberFormats : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A:H"].ColumnWidth = 17;

            // Display 111 as 111.
            worksheet.Range["A1"].Value = 111;
            worksheet.Range["A1"].NumberFormat = "#####";

            // Display 222 as 00222.
            worksheet.Range["B1"].Value = 222;
            worksheet.Range["B1"].NumberFormat = "00000";

            // Display 12345678 as 12,345,678.
            worksheet.Range["C1"].Value = 12345678;
            worksheet.Range["C1"].NumberFormat = "#,#";

            // Display .126 as 0.13.
            worksheet.Range["D1"].Value = .126;
            worksheet.Range["D1"].NumberFormat = "0.##";

            // Display 74.4 as 74.400.
            worksheet.Range["E1"].Value = 74.4;
            worksheet.Range["E1"].NumberFormat = "##.000";

            // Display 1.6 as 160.0%.
            worksheet.Range["F1"].Value = 1.6;
            worksheet.Range["F1"].NumberFormat = "0.0%";

            // Display 4321 as $4,321.00.
            worksheet.Range["G1"].Value = 4321;
            worksheet.Range["G1"].NumberFormat = "$#,##0.00";

            // Display 8.75 as 8 3/4.
            worksheet.Range["H1"].Value = 8.75;
            worksheet.Range["H1"].NumberFormat = "# ?/?";
        }

    }
}
