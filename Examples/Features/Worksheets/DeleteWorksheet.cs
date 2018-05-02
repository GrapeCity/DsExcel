using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class DeleteWorksheet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {           
            IWorksheet worksheet = workbook.Worksheets.Add();

            //workbook must contain one visible worksheet at least, if delete the one visible worksheet, it will throw exception.
            worksheet.Delete();
        }
    }
}
