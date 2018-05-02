using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace GrapeCity.Documents.Excel.Examples.Features.Picture
{
    public class DeletePicture : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            System.IO.Stream stream = this.GetResourceStream("logo.png");
            GrapeCity.Documents.Excel.Drawing.IShape picture = worksheet.Shapes.AddPicture(stream, GrapeCity.Documents.Excel.Drawing.ImageType.PNG, 20, 20, 100, 100);
            //set picture size.
            picture.WidthInPixel = 700;
            picture.HeightInPixel = 120;
            //config picture layout.
            picture.Fill.Solid();
            picture.Fill.Color.ObjectThemeColor = ThemeColor.Accent1;

            //delete picture.
            picture.Delete();
        }
    }
}
