using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataValidation
{
    public class ModifyValidation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["C2:E4"].Value = new object[,]
            {
                {new TimeSpan(13, 0, 0), new TimeSpan(13, 29, 59), new TimeSpan(13, 30, 0)},
                {new TimeSpan(14, 0, 0), new TimeSpan(15, 0, 0), new TimeSpan(16, 30, 0)},
                {new TimeSpan(19, 0, 0), new TimeSpan(18, 29, 59), new TimeSpan(18, 30, 0)}
            };

            //create time validation.
            worksheet.Range["C2:E4"].Validation.Add(ValidationType.Time, ValidationAlertStyle.Stop, ValidationOperator.Between, new TimeSpan(13, 30, 0), new TimeSpan(18, 30, 0));

            //Modify validation.
            worksheet.Range["C2:E4"].Validation.Formula1 = new TimeSpan(14, 30, 0);
            worksheet.Range["C2:E4"].Validation.Formula2 = new TimeSpan(19, 0, 0);

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
