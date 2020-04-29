using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class GetRangeBoundary : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\Sport sign-up sheet.xlsx");
            workbook.Open(fileStream);

            IWorksheet worksheet = workbook.Worksheets[0];


            // Get the absolute location and size of the Range["G1"] in the worksheet.
            IRange range = worksheet.Range["G1"];
            Rectangle rect = GrapeCity.Documents.Excel.CellInfo.GetRangeBoundary(range);
            // Add the image to the Range["G1"]
            System.IO.Stream stream = this.GetResourceStream("logo.png");
            worksheet.Shapes.AddPictureInPixel(stream, GrapeCity.Documents.Excel.Drawing.ImageType.PNG, rect.X, rect.Y, rect.Width, rect.Height);
        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Sport sign-up sheet.xlsx", "logo.png" };
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
