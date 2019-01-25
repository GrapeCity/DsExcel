using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace GrapeCity.Documents.Excel.Examples.Features.Picture
{
    public class CreatePicture : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            System.IO.Stream stream = this.GetResourceStream("logo.png");
            GrapeCity.Documents.Excel.Drawing.IShape picture = worksheet.Shapes.AddPicture(stream, GrapeCity.Documents.Excel.Drawing.ImageType.PNG, 20, 20, 790, 120);

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
