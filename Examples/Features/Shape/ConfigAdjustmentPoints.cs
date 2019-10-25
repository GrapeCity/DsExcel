using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ConfigAdjustmentPoints : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            // Add a right arrow callout.
            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Excel.Drawing.AutoShapeType.RightArrowCallout, 20, 20, 200, 100);

            GrapeCity.Documents.Excel.Drawing.IAdjustments adjustments = shape.Adjustments;
            adjustments[0] = 0.5;// arrow neck width
            adjustments[1] = 0.4;// arrow head width
            adjustments[2] = 0.5;// arrow head height
            adjustments[3] = 0.6;// text box width
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
