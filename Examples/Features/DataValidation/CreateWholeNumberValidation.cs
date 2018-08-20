using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataValidation
{
    public class CreateWholeNumberValidation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["C2:E4"].Value = new object[,]
            {
                {1, 3, 5},
                {7, 9, 11},
                {13, 15, 17}
            };

            //add whole number validation.
            worksheet.Range["C2:E4"].Validation.Add(ValidationType.Whole, ValidationAlertStyle.Stop, ValidationOperator.Between, 1, 8);
            IValidation validation = worksheet.Range["C2:E4"].Validation;
            validation.IgnoreBlank = true;
            validation.InputTitle = "Tips";
            validation.InputMessage = "Input a value between 1 and 8, please";
            validation.ErrorTitle = "Error";
            validation.ErrorMessage = "input value does not between 1 and 8";
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
