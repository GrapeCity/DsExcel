using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Picture
{
    public class PictureCopy : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            System.IO.Stream stream = this.GetResourceStream("logo.png");

            //Create a shape in worksheet, picture's range is Range["A2:I6"]
            GrapeCity.Documents.Excel.Drawing.IShape picture = worksheet.Shapes.AddPicture(stream, GrapeCity.Documents.Excel.Drawing.ImageType.PNG, 20, 20, 395, 60);

            //Range["A2:I6"] must contain picture's range, copy a new picture to Range["J2:R6"]
            worksheet.Range["A2:I6"].Copy(worksheet.Range["J2"]);
            //worksheet.Range["A2:I6"].Copy(worksheet.Range["J2:R6"]);

            //Cross sheet copy, copy a new shape to worksheet2's Range["J2:R6"]
            //IWorksheet worksheet2 = workbook.Worksheets.Add()
            //worksheet.Range["A2:I6"].Copy(worksheet2.Range["J2"]);
            //worksheet.Range["A2:I6"].Copy(worksheet2.Range["J2:R6"]);

        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "logo.png" };
            }
        }
    }
}
