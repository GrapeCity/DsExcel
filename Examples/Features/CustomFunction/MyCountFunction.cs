using GrapeCity.Documents.Excel.CustomFunctions;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.CustomFunction
{
    public class MyCountFunction : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            GrapeCity.Documents.Excel.Workbook.AddCustomFunction(new MyCountX());
            IWorksheet worksheet = workbook.Worksheets[0];

            object[,] data = new object[,]{
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new DateTime(1968, 6, 8), "Blue", 67, 165},
                {"Nia", "New York", new DateTime(1972, 7, 3), "Brown", 62, 134},
                {"Jared", "New York", new DateTime(1964, 3, 2), "Hazel", 72, 180},
                {"Natalie", "Washington", new DateTime(1972, 8, 8), "Blue", 66, 163},
                {"Damon", "Washington", new DateTime(1986, 2, 2), "Hazel", 76, 176},
                {"Angela", "Washington", new DateTime(1993, 2, 15), "Brown", 68, 145}
            };

            worksheet.Range["A1:F7"].Value = data;
            worksheet.Range["A:F"].ColumnWidth = 15;

            worksheet.Range["E6"].Formula = "=MyCount(E2:E7)";

            // the result would be 2
            var value = worksheet.Range["E6"].Value;

            #region Implementation of MyAddFunctionX

            //public class MyCountX : Function
            //{
            //    // Make AcceptReference to true, then you can get the source range when evaluating
            //    public MyCountX() : base("MyCount", ResultType.Number, new Parameter[] { new Parameter(ParameterType.Object, true) })
            //    {

            //    }

            //    public override object Evaluate(object[] arguments, ICalcContext context)
            //    {
            //        int count = 0;
            //        CalcReference calcReference = arguments[0] as CalcReference;
            //        foreach (IRange srcRange in calcReference.GetRanges())
            //        {
            //            object[,] values = srcRange.Value as object[,];
            //            for (int row = 0; row < srcRange.Rows.Count; row++)
            //            {
            //                for (int col = 0; col < srcRange.Columns.Count; col++)
            //                {
            //                    if (values[row, col] is double)
            //                    {
            //                        double num = (double)values[row, col];
            //                        if (num > 70)
            //                        {
            //                            count++;
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //        return count;
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

        public class MyCountX : Function
        {
            // Make AcceptReference to true, then you can get the source range when evaluating
            public MyCountX() : base("MyCount", ResultType.Number, new Parameter[] { new Parameter(ParameterType.Object, true) })
            {

            }

            public override object Evaluate(object[] arguments, ICalcContext context)
            {
                int count = 0;
                CalcReference calcReference = arguments[0] as CalcReference;
                foreach (IRange srcRange in calcReference.GetRanges())
                {
                    object[,] values = srcRange.Value as object[,];
                    for (int row = 0; row < srcRange.Rows.Count; row++)
                    {
                        for (int col = 0; col < srcRange.Columns.Count; col++)
                        {
                            if (values[row, col] is double)
                            {
                                double num = (double)values[row, col];
                                if (num > 70)
                                {
                                    count++;
                                }
                            }
                        }
                    }
                }
                return count;
            }

        }
    }

}
