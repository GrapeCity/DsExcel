using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class SetBackgroundImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Use sheet index to get worksheet.
            IWorksheet worksheet = workbook.Worksheets[0];

            Stream stream = this.GetResourceStream("logo.png");
            byte[] imageBytes = new byte[stream.Length];
            stream.Read(imageBytes, 0, imageBytes.Length);
            //Set a background image for worksheet
            worksheet.BackgroundPicture = imageBytes;
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool IsNew => true;
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "logo.png" };
            }
        }
    }
}
