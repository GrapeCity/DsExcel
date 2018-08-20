using GrapeCity.Documents.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CustomFunctions
{
    public class MyConcatenateFunction : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            GrapeCity.Documents.Excel.Workbook.AddCustomFunction(new MyConcatenateFunctionX());

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Formula = "=MyConcatenate(\"I\", \" \", \"live\", \" \", \"in\", \" \", \"Xi'an\", \".\")";
            worksheet.Range["A2"].Formula = "=MyConcatenate(A1, \"haha.\")";

            worksheet.Range["B1"].Value = 12;
            worksheet.Range["B2"].Value = 34;
            worksheet.Range["B3"].Formula = "=MyConcatenate(B1, B2)";

            worksheet.Range["M5:N5"].FormulaArray = "=CONCATENATE({\"aa\",\"bb\"}, 12, 34)";

            //"I live in Xi'an."
            var resultA1 = worksheet.Range["A1"].Value;
            //"I live in Xi'an.haha."
            var resultA2 = worksheet.Range["A2"].Value;
            //"1234"
            var resultB3 = worksheet.Range["B3"].Value;
            //"aa1234"
            var resultM5 = worksheet.Range["M5"].Value;
            //"bb1234"
            var resultN5 = worksheet.Range["N5"].Value;

            /* Implementation of MyAddFunctionX

            public class MyConcatenateFunctionX : CustomFunction
            {
                public MyConcatenateFunctionX()
                    : base("MyConcatenate", FunctionValueType.Text, CreateParameters())
                {

                }

                private static Parameter[] CreateParameters()
                {
                    Parameter[] parameters = new Parameter[254];
                    for (int i = 0; i < 254; i++)
                    {
                        parameters[i] = new Parameter(FunctionValueType.Variant);
                    }

                    return parameters;
                }

                public override object Evaluate(object[] arguments, ICalcContext context)
                {
                    StringBuilder sb = new StringBuilder();

                    string result = string.Empty;
                    foreach (var argument in arguments)
                    {
                        if (argument is CalcError)
                        {
                            return argument;
                        }
                        if (argument is string || argument is double)
                        {
                            sb.Append(argument);
                        }
                    }

                    return sb.ToString();
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



        public class MyConcatenateFunctionX : CustomFunction
        {
            public MyConcatenateFunctionX()
                : base("MyConcatenate", FunctionValueType.Text, CreateParameters())
            {

            }

            private static Parameter[] CreateParameters()
            {
                Parameter[] parameters = new Parameter[254];
                for (int i = 0; i < 254; i++)
                {
                    parameters[i] = new Parameter(FunctionValueType.Variant);
                }

                return parameters;
            }

            public override object Evaluate(object[] arguments, ICalcContext context)
            {
                StringBuilder sb = new StringBuilder();

                string result = string.Empty;
                foreach (var argument in arguments)
                {
                    if (argument is CalcError)
                    {
                        return argument;
                    }
                    if (argument is string || argument is double)
                    {
                        sb.Append(argument);
                    }
                }

                return sb.ToString();
            }

        }


    }
}
