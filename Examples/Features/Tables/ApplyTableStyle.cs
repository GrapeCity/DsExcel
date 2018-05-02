using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Tables
{
    public class ApplyTableStyle : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //add table.
            IWorksheet worksheet = workbook.Worksheets[0];
            ITable table = worksheet.Tables.Add(worksheet.Range["A1:F7"], true);
            worksheet.Range["A:F"].ColumnWidth = 15;

            //Add one custom table style.
            ITableStyle style = workbook.TableStyles.Add("test");
            //set custom table style for table.
            table.TableStyle = style;

            //Use table style name get one build in table style.
            ITableStyle tableStyle = workbook.TableStyles["TableStyleMedium3"];
            //set built-in table style for table.
            table.TableStyle = tableStyle;
        }
    }
}
