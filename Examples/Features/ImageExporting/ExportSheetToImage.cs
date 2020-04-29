using GrapeCity.Documents.Drawing;
using GrapeCity.Documents.Imaging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ImageExporting
{
    public class ExportSheetToImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\Home inventory.xlsx");
            workbook.Open(fileStream);

            IWorksheet worksheet = workbook.Worksheets[0];

            // Save the worksheet as image to a stream.
            worksheet.ToImage(outputStream, Drawing.ImageType.PNG);
        }

        public override bool SaveAsImages 
        {
            get
            {
                return true;
            }
        }

        public override string TemplateName
        {
            get
            {
                return "Employee absence schedule.xlsx";
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
                return new string[] { "xlsx\\Home inventory.xlsx" };
            }
        }
    }
}
