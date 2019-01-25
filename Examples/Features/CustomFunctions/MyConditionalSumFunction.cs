using GrapeCity.Documents.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CustomFunctions
{
    public class MyConditionalSumFunction : ExampleBase
    {

        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            GrapeCity.Documents.Excel.Workbook.AddCustomFunction(new MyConditionalSumFunctionX());

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1:A10"].Value = new object[,] {
                                                                {1 },
                                                                {2 },
                                                                {3 },
                                                                {4 },
                                                                {5 },
                                                                {6 },
                                                                {7 },
                                                                {8 },
                                                                {9 },
                                                                {10 }};
            IFormatCondition cellValueRule = worksheet.Range["A1:A10"].FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Greater, 5) as IFormatCondition;
            cellValueRule.Interior.Color = System.Drawing.Color.Red;

            //Sum cells value which display format interior color are red.
            worksheet.Range["C1"].Formula = "=MyConditionalSum(A1:A10)";

            //Range["C1"]'s value is 40.
            var result = worksheet.Range["C1"].Value;


            /* Implementation of MyConditionalSumFunctionX
            
            public class MyConditionalSumFunctionX : CustomFunction
            {
                public MyConditionalSumFunctionX() : base("MyConditionalSum", FunctionValueType.Number, CreateParameters())
                {

                }

                private static Parameter[] CreateParameters()
                {
                    Parameter[] parameters = new Parameter[254];
                    for (int i = 0; i < 254; i++)
                    {
                        parameters[i] = new Parameter(FunctionValueType.Object, true);
                    }

                    return parameters;
                }

                public override object Evaluate(object[] arguments, ICalcContext context)
                {
                    double sum = 0d;
                    foreach (var argument in arguments)
                    {
                        foreach (var item in Enumerate(argument))
                        {
                            if (item is CalcError)
                            {
                                return item;
                            }
                            if (item is double)
                            {
                                sum += (double)item;
                            }
                        }
                    }

                    return sum;
                }

                private static IEnumerable<object> Enumerate(object obj)
                {
                    if (obj is IEnumerable<object>)
                    {
                        foreach (var item in obj as IEnumerable<object>)
                        {
                            foreach (var item2 in Enumerate(item))
                            {
                                yield return item2;
                            }
                        }
                    }
                    else if (obj is object[,])
                    {
                        var array = obj as object[,];
                        int rowCount = array.GetLength(0);
                        int colCount = array.GetLength(1);
                        for (int i = 0; i < rowCount; i++)
                        {
                            for (int j = 0; j < colCount; j++)
                            {
                                yield return array[i, j];
                            }
                        }
                    }
                    else if (obj is CalcReference)
                    {
                        foreach (var item in Enumerate(obj as CalcReference))
                        {
                            yield return item;
                        }
                    }

                    yield return obj;
                }

                private static IEnumerable<object> Enumerate(CalcReference reference)
                {
                    foreach (var range in reference.GetRanges())
                    {
                        int rowCount = range.Rows.Count;
                        int colCount = range.Columns.Count;
                        for (int i = 0; i < rowCount; i++)
                        {
                            for (int j = 0; j < colCount; j++)
                            {
                                if (range.Cells[i, j].DisplayFormat.Interior.Color == System.Drawing.Color.Red)
                                {
                                    yield return range.Cells[i, j].Value;
                                }
                            }
                        }
                    }
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
        

        public class MyConditionalSumFunctionX : CustomFunction
        {
            public MyConditionalSumFunctionX() : base("MyConditionalSum", FunctionValueType.Number, CreateParameters())
            {

            }

            private static Parameter[] CreateParameters()
            {
                Parameter[] parameters = new Parameter[254];
                for (int i = 0; i < 254; i++)
                {
                    parameters[i] = new Parameter(FunctionValueType.Object, true);
                }

                return parameters;
            }

            public override object Evaluate(object[] arguments, ICalcContext context)
            {
                double sum = 0d;
                foreach (var argument in arguments)
                {
                    foreach (var item in Enumerate(argument))
                    {
                        if (item is CalcError)
                        {
                            return item;
                        }
                        if (item is double)
                        {
                            sum += (double)item;
                        }
                    }
                }

                return sum;
            }

            private static IEnumerable<object> Enumerate(object obj)
            {
                if (obj is IEnumerable<object>)
                {
                    foreach (var item in obj as IEnumerable<object>)
                    {
                        foreach (var item2 in Enumerate(item))
                        {
                            yield return item2;
                        }
                    }
                }
                else if (obj is object[,])
                {
                    var array = obj as object[,];
                    int rowCount = array.GetLength(0);
                    int colCount = array.GetLength(1);
                    for (int i = 0; i < rowCount; i++)
                    {
                        for (int j = 0; j < colCount; j++)
                        {
                            yield return array[i, j];
                        }
                    }
                }
                else if (obj is CalcReference)
                {
                    foreach (var item in Enumerate(obj as CalcReference))
                    {
                        yield return item;
                    }
                }

                yield return obj;
            }

            private static IEnumerable<object> Enumerate(CalcReference reference)
            {
                foreach (var range in reference.GetRanges())
                {
                    int rowCount = range.Rows.Count;
                    int colCount = range.Columns.Count;
                    for (int i = 0; i < rowCount; i++)
                    {
                        for (int j = 0; j < colCount; j++)
                        {
                            if (range.Cells[i, j].DisplayFormat.Interior.Color == System.Drawing.Color.Red)
                            {
                                yield return range.Cells[i, j].Value;
                            }
                        }
                    }
                }
            }
        }

    }
}
