using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting
{
    public class CreateStyleBasedOn : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Style = workbook.Styles["Good"];
            worksheet.Range["A1"].Value = "Good";

            // Create and modify a style based on current existing style
            IStyle myGood = workbook.Styles.Add("MyGood", workbook.Styles["Good"]);
            myGood.Font.Bold = true;
            myGood.Font.Italic = true;

            worksheet.Range["B1"].Style = workbook.Styles["MyGood"];
            worksheet.Range["B1"].Value = "MyGood";
        }

        public override bool IsNew => true;
    }
}
