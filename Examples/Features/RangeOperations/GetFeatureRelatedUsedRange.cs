using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class GetFeatureRelatedUsedRange : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1:B2"].Value = new object[,]
            {
                {1, 2},
                {"aaa", "bbb"}
            };
            worksheet.Range["A2:C3"].Interior.Color = Color.Green;

            //style used range is A2:C3.
            var UsedRange_style = worksheet.GetUsedRange(UsedRangeType.Style);
            UsedRange_style.Interior.Color = Color.LightBlue;
        }
    }
}
