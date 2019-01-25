using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CustomFunctions
{
    public class MyIsMergedRangeFunction : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            GrapeCity.Documents.Excel.Workbook.AddCustomFunction(new MyIsMergedRangeFunctionX());

            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1:B2"].Merge();

            worksheet.Range["C1"].Formula = "=MyIsMergedRange(A1)";
            worksheet.Range["C2"].Formula = "=MyIsMergedRange(H2)";

            //A1 is a merged cell, Range["C1"]'s value is true.
            var resultC1 = worksheet.Range["C1"].Value;

            //H2 is not a merged cell, Range["C2"]'s value is false.
            var resultC2 = worksheet.Range["C2"].Value;


            /* Implementation of MyIsErrorFunctionX
            
            public class MyIsMergedRangeFunctionX : CustomFunction
            {
                public MyIsMergedRangeFunctionX()
                    : base("MyIsMergedRange", FunctionValueType.Boolean, new Parameter[] { new Parameter(FunctionValueType.Object, true) })
                {

                }

                public override object Evaluate(object[] arguments, ICalcContext context)
                {
                    if (arguments[0] is CalcReference)
                    {
                        IEnumerable<IRange> ranges = (arguments[0] as CalcReference).GetRanges();

                        foreach (var range in ranges)
                        {
                            return range.MergeCells;
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

        public class MyIsMergedRangeFunctionX : CustomFunction
        {
            public MyIsMergedRangeFunctionX()
                : base("MyIsMergedRange", FunctionValueType.Boolean, new Parameter[] { new Parameter(FunctionValueType.Object, true) })
            {

            }

            public override object Evaluate(object[] arguments, ICalcContext context)
            {
                if (arguments[0] is CalcReference)
                {
                    IEnumerable<IRange> ranges = (arguments[0] as CalcReference).GetRanges();

                    foreach (var range in ranges)
                    {
                        return range.MergeCells;
                    }
                }

                return false;
            }
        }

    }
}
