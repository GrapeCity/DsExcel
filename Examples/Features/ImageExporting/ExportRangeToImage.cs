using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ImageExporting
{
    public class ExportRangeToImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\Personal Monthly Budget.xlsx");

            //Open a xlsx file
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Export range "B14:E25" to image
            worksheet.Range["B14:E25"].ToImage(outputStream, ImageType.PNG);
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

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Personal Monthly Budget.xlsx" };
            }
        }
    }
}
