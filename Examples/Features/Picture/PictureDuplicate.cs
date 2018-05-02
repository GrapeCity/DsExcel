using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Picture
{
    public class PictureDuplicate : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            System.IO.Stream stream = this.GetResourceStream("logo.png");

            //Create a shape in worksheet
            GrapeCity.Documents.Excel.Drawing.IShape picture = worksheet.Shapes.AddPicture(stream, GrapeCity.Documents.Excel.Drawing.ImageType.PNG, 20, 20, 395, 60);

            //Duplicate picture
            GrapeCity.Documents.Excel.Drawing.IShape newPicture = picture.Duplicate();

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
