using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataValidation
{
    public class CreateCustomValidation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A2"].Value = 1;
            worksheet.Range["A3"].Value = 2;

            worksheet.Range["C2"].Value = 0;

            //create custom validation, if the expression "=$C$2" result is true, the cell's validation will be true, otherwise, it is false.
            //when use custom validation, validationOperator and formula2 parameters will be ignored even if you have given.
            worksheet.Range["A2:A3"].Validation.Add(ValidationType.Custom, ValidationAlertStyle.Information, formula1:"=$C$2");

            //judge if Range["A2:A3"] has validation.
            for (int i = 1; i <= 2; i++)
            {
                if (worksheet.Range[i, 0].HasValidation)
                {
                    //set the range[i, 0]'s interior color.
                    worksheet.Range[i, 0].Interior.Color = Color.LightBlue;
                }
            }
        }
    }
}
