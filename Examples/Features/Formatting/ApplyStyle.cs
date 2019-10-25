using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting
{
    public class ApplyStyle : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Change to build in name style.
            worksheet.Range["A1"].Value = "Bad";
            worksheet.Range["A1"].Style = workbook.Styles["Bad"];
        }
    }
}
