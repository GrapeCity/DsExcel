using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formulas
{
    public class ConfigWorkbookReferenceStyle : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //set workbook's reference style to R1C1. exported xlsx file will be R1C1 style.
            workbook.ReferenceStyle = ReferenceStyle.R1C1;
        }
    }
}
