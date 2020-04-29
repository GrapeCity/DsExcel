using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class AddShapeWithCustomName : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape("custom parallelogram", GrapeCity.Documents.Excel.Drawing.AutoShapeType.Parallelogram, 1, 1, 200, 100);

            //Get shape by name
            GrapeCity.Documents.Excel.Drawing.IShape parallelogram = worksheet.Shapes["custom parallelogram"];
            parallelogram.Fill.Color.RGB = System.Drawing.Color.Red;
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }

    }
}
