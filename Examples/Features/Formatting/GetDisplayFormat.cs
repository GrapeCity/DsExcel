using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting
{
    public class GetDisplayFormat : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Range A1's displaying color is table style
            worksheet.Tables.Add(worksheet.Range["A1:E5"], true);
            var color_A1 = worksheet.Range["A1"].DisplayFormat.Interior.Color;

            //Range A1's displaying color will be cell style, yellow.
            worksheet.Range["A1"].Interior.Color = Color.Yellow;
            var color_A1_1 = worksheet.Range["A1"].DisplayFormat.Interior.Color;

            //Range A1's displaying color will be conditional format style, green.
            IFormatCondition condition = worksheet.Range["A1"].FormatConditions.Add(FormatConditionType.NoBlanksCondition) as IFormatCondition;
            condition.Interior.Color = Color.Green;
            var color_A1_2 = worksheet.Range["A1"].DisplayFormat.Interior.Color;
        }

    }
}
