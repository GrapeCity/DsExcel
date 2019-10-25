using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formulas.Precedentsanddependents
{
    public class GetAllPrecedents : ExampleBase
    {
        public override bool IsNew => true;
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["E2"].Formula = "=sum(C1:C2)";
            worksheet.Range["C1"].Formula = "=B1";
            worksheet.Range["B1"].Formula = "=sum(A1:A2)";
            worksheet.Range["A1"].Value = 1;
            worksheet.Range["A2"].Value = 2;
            worksheet.Range["C2"].Value = 3;

            List<IRange> list = new List<IRange>();
            foreach (var item in worksheet.Range["E2"].GetPrecedents())
            {
                list.Add(item);
            }

            while (list.Count > 0)
            {
                var temp = list;
                list = new List<IRange>();
                foreach (var item in temp)
                {
                    for (int i = 0; i < item.RowCount; i++)
                    {
                        for (int j = 0; j < item.ColumnCount; j++)
                        {
                            var dependents = item.Cells[i, j].GetPrecedents();
                            if (dependents.Count == 0)
                            {
                                item.Cells[i, j].Interior.Color = Color.SkyBlue;
                            }
                            else
                            {
                                item.Cells[i, j].Interior.Color = Color.Gray;
                                list.AddRange(dependents);
                            }
                        }
                    }
                }
            }
        }
    }
}
