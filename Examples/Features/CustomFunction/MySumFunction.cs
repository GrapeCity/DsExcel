using GrapeCity.Documents.Excel.CustomFunctions;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CustomFunction
{
    public class MySumFunction : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            GrapeCity.Documents.Excel.Workbook.AddCustomFunction(new MySumFunctionX());

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Value = 1000;
            worksheet.Range["B1"].Value = 2000;
            worksheet.Range["C1"].Formula = "=MySum(A1:B1)";
            
            //Range["C1"]'s value is 12.
            var result = worksheet.Range["C1"].Value;


            #region Implementation of MySumFunctionX

            //public class MySumFunctionX : Function
            //{
            //    public MySumFunctionX() : base("MYSUM", ResultType.Number, new Parameter[] { new Parameter(ParameterType.ObjectArray) })
            //    {

            //    }

            //    public override object Evaluate(object[] arguments, ICalcContext context)
            //    {
            //        double sum = 0d;
            //        foreach (var argument in arguments)
            //        {
            //            foreach (var item in Enumerate(argument))
            //            {
            //                if (item is CalcError)
            //                {
            //                    return item;
            //                }
            //                if (item is double)
            //                {
            //                    sum += (double)item;
            //                }
            //            }
            //        }

            //        return sum;
            //    }

            //    private static IEnumerable<object> Enumerate(object obj)
            //    {
            //        if (obj is IEnumerable<object>)
            //        {
            //            foreach (var item in obj as IEnumerable<object>)
            //            {
            //                foreach (var item2 in Enumerate(item))
            //                {
            //                    yield return item2;
            //                }
            //            }
            //        }
            //        else if (obj is object[,])
            //        {
            //            var array = obj as object[,];
            //            int rowCount = array.GetLength(0);
            //            int colCount = array.GetLength(1);
            //            for (int i = 0; i < rowCount; i++)
            //            {
            //                for (int j = 0; j < colCount; j++)
            //                {
            //                    yield return array[i, j];
            //                }
            //            }
            //        }

            //        if (obj != null)
            //        {
            //            yield return obj;
            //        }
            //    }
            //}

            #endregion


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

        public class MySumFunctionX : Function
        {
            public MySumFunctionX() : base("MYSUM", ResultType.Number, new Parameter[] { new Parameter(ParameterType.Object) })
            {

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
                            double value = (double)item;
                            if (value >= 500)
                            {
                                sum += (double)item;
                            }
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

                if (obj != null)
                {
                    yield return obj;
                }
            }
        }


    }

}
