using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Grouping
{
    public class ShowSpecificLevel : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A:N"].Group();
            worksheet.Range["A:F"].Group();
            worksheet.Range["A:C"].Group();

            worksheet.Range["Q:Z"].Group();
            worksheet.Range["Q:T"].Group();

            //level 3 and level 4 will be collapsed. level 2 and level 1 expand.
            worksheet.Outline.ShowLevels(columnLevels: 2);
        }
    }
}
