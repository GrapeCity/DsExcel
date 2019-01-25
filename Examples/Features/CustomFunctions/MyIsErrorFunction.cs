using GrapeCity.Documents.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CustomFunctions
{
    public class MyIsErrorFunction : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            GrapeCity.Documents.Excel.Workbook.AddCustomFunction(new MyIsErrorFunctionX());

            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1"].Value = CalcError.Num;
            worksheet.Range["A2"].Value = 100;

            worksheet.Range["B1"].Formula = "=MyIsError(A1)";
            worksheet.Range["B2"].Formula = "=MyIsError(A2)";

            //Range["B1"]'s value is true.
            var resultB1 = worksheet.Range["B1"].Value;

            //Range["B2"]'s value is false.
            var resultB2 = worksheet.Range["B2"].Value;


            /* Implementation of MyIsErrorFunctionX
  
            public class MyIsErrorFunctionX : CustomFunction
            {
                public MyIsErrorFunctionX()
                    : base("MyIsError", FunctionValueType.Boolean, new Parameter[] { new Parameter(FunctionValueType.Variant) })
                {

                }

                public override object Evaluate(object[] arguments, ICalcContext context)
                {
                    if (arguments[0] is CalcError)
                    {
                        if ((CalcError)arguments[0] != CalcError.None && (CalcError)arguments[0] != CalcError.GettingData)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }

                    return false;
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

        public class MyIsErrorFunctionX : CustomFunction
        {
            public MyIsErrorFunctionX()
                : base("MyIsError", FunctionValueType.Boolean, new Parameter[] { new Parameter(FunctionValueType.Variant) })
            {

            }

            public override object Evaluate(object[] arguments, ICalcContext context)
            {
                if (arguments[0] is CalcError)
                {
                    if ((CalcError)arguments[0] != CalcError.None && (CalcError)arguments[0] != CalcError.GettingData)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                return false;
            }
        }

    }
}
