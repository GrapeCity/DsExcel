using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataValidation
{
    public class CreateDateValidation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["C2:E4"].Value = new object[,]
            {
                {new DateTime(2015, 12, 1), new DateTime(2015, 12, 14), new DateTime(2015, 12, 15) },
                {new DateTime(2015, 12, 18), new DateTime(2015, 12, 19), new DateTime(2015, 12, 30) },
                {new DateTime(2015, 12, 31), new DateTime(2014, 12, 13), new DateTime(2017, 12, 15) }
            };

            //create date validation.
            worksheet.Range["C2:E4"].Validation.Add(ValidationType.Date, ValidationAlertStyle.Stop, ValidationOperator.Between, new DateTime(2015, 12, 13), new DateTime(2015, 12, 18));

            //set column width just for export shown.
            worksheet.Range["C:E"].EntireColumn.ColumnWidthInPixel = 120;

            //judge if Range["C2:E4"] has validation.
            for (int i = 1; i <= 3; i++)
            {
                for (int j = 2; j <= 4; j++)
                {
                    if (worksheet.Range[i, j].HasValidation)
                    {
                        //set the range[i, j]'s interior color.
                        worksheet.Range[i, j].Interior.Color = Color.LightBlue;
                    }
                }
            }
        }
    }
}
