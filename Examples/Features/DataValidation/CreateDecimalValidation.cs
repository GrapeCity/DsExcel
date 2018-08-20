using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataValidation
{
    public class CreateDecimalValidation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["C2:E4"].Value = new object[,]
            {
                {1, 3.0, 3.4},
                {7, 9, 102.7},
                {102.8, 110, 120}
            };

            //add decimal validation.
            worksheet.Range["C2:E4"].Validation.Add(ValidationType.Decimal, ValidationAlertStyle.Stop, ValidationOperator.Between, 3.4, 102.8);
            IValidation validation = worksheet.Range["C2:E4"].Validation;
            validation.IgnoreBlank = true;
            validation.InputTitle = "Tips";
            validation.InputMessage = "Input a decimal value between 3.4 and 102.8, please.";
            validation.ErrorTitle = "Error";
            validation.ErrorMessage = "input value does not between 3.4 and 102.8.";
            validation.ShowInputMessage = true;
            validation.ShowError = true;

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
