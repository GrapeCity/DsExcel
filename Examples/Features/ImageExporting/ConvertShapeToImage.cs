using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ImageExporting
{
    public class ConvertShapeToImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Add a oval
            IShape oval = worksheet.Shapes.AddShape(AutoShapeType.Oval, 0, 0, 191, 194);

            // Set soild fill for rectangle
            oval.Fill.Color.RGB = Color.OrangeRed;

            // Save the shape as image to a stream.
            oval.ToImage(outputStream, ImageType.PNG);
        }

        public override bool SaveAsImages
        {
            get
            {
                return true;
            }
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
