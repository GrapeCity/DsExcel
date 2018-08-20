using GrapeCity.Documents.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CustomFunctions
{
    public class MyAddFunction : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            GrapeCity.Documents.Excel.Workbook.AddCustomFunction(new MyAddFunctionX());

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Value = 1;
            worksheet.Range["B1"].Value = 2;
            worksheet.Range["C1"].Formula = "=MyAdd(A1, B1)";
            //Range["C1"]'s value is 3.
            var result = worksheet.Range["C1"].Value;

            worksheet.Range["E1:F2"].Value = new object[,] { { 1, 3},
                                                             { 2, 4}};

            //Custom function use in array formula.
            worksheet.Range["G1:G2"].FormulaArray = "=MyAdd(E1:E2, F1:F2)";

            //Range["G1"]'s value is 4, Range["G2"]'s value is 6.
            var resultG1 = worksheet.Range["G1"].Value;
            var resultG2 = worksheet.Range["G2"].Value;

            /* Implementation of MyAddFunctionX

            public class MyAddFunctionX : CustomFunction
            {
                public MyAddFunctionX()
                    : base("MyAdd", FunctionValueType.Number, new Parameter[] { new Parameter(FunctionValueType.Number), new Parameter(FunctionValueType.Number) })
                {

                }

                public override object Evaluate(object[] arguments, ICalcContext context)
                {
                    return (double)arguments[0] + (double)arguments[1];
                }
            }

            */
        }

        public override bool CanDownload
        {
            get
            {
                return false;
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }

        public class MyAddFunctionX : CustomFunction
        {
            public MyAddFunctionX()
                : base("MyAdd", FunctionValueType.Number, new Parameter[] { new Parameter(FunctionValueType.Number), new Parameter(FunctionValueType.Number) })
            {

            }

            public override object Evaluate(object[] arguments, ICalcContext context)
            {
                return (double)arguments[0] + (double)arguments[1];
            }
        }

    }
}
