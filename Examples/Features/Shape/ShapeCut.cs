﻿using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ShapeCut : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Create a shape in worksheet, shape's range is Range["A7:B7"]
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Excel.Drawing.AutoShapeType.Rectangle, 1, 1, 100, 100);

            //Range["A1:D10"] must contain Range["A7:B7"], cut a new shape to Range["C1:F7"]
            worksheet.Range["A1:D10"].Cut(worksheet.Range["E1"]);
            //worksheet.Range["A1:D10"].Cut(worksheet.Range["E1:I9"]);

            //Cross sheet cut, cut a new shape to worksheet2's Range["C1:F7"]
            //IWorksheet worksheet2 = workbook.Worksheets.Add()
            //worksheet.Range["A1:D10"].Cut(worksheet2.Range["E1"]);
            //worksheet.Range["A1:D10"].Cut(worksheet2.Range["E1:I9"]);

        }

        public override bool ShowViewer
        {
            get
            {
                return true;
            }
        }
        
    }
}
