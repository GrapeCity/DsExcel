using GrapeCity.Documents.Excel.Drawing;
using System.IO;

namespace GrapeCity.Documents.Excel.Examples.Features.ImageExporting
{
    public class ExportShapeToImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\GroupShape.xlsx");
            
            //Open a xlsx file contains a group shape 
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Export the shape to image
            worksheet.Shapes[0].ToImage(outputStream, ImageType.PNG);
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
                return new string[] { "xlsx\\GroupShape.xlsx" };
            }
        }
    }
}
