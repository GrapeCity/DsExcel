using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ImageExporting
{
    public class ConvertRangeToImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set value
            worksheet.Range["A1:C1"].Value = new string[] { "Device", "Quantity", "Unit Price" };
            worksheet.Range["A2:C5"].Value = new object[,] {
                { "T540p", 12, 9850 },
                { "T570", 5, 7460 },
                { "Y460", 6, 5400 },
                { "Y460F", 8, 6240 }
            };

            //Set style
            worksheet.Range["A1:C1"].Font.Bold = true;
            worksheet.Range["A1:C1"].Font.Color = Color.White;
            worksheet.Range["A1:C1"].Interior.Color = Color.LightBlue;
            worksheet.Range["A2:C5"].Borders[BordersIndex.InsideHorizontal].Color = Color.Orange;
            worksheet.Range["A2:C5"].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.DashDot;

            // Save the range "A1:C5" as an image to a stream.
            worksheet.Range["A1:C5"].ToImage(outputStream, Drawing.ImageType.PNG);
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
