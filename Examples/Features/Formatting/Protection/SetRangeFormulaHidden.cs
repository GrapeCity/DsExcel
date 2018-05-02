using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Protection
{
    public class SetRangeFormulaHidden : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["B1"].Formula = "=A1";

            //config range B1's FormulaHidden property.
            worksheet.Range["B1"].FormulaHidden = true;
            //protect worksheet, range B1's formula will not show in exported xlsx file.
            worksheet.Protection = true;
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
