using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.ImageExporting
{
    public class ExportChartToImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\Template_SalesTracker_report.xlsx");

            //Open a xlsx file contains a chart 
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Export the chart to image
            worksheet.Shapes["ProductIncomeChart"].ToImage(outputStream, ImageType.PNG);
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
                return new string[] { "xlsx\\Template_SalesTracker_report.xlsx" };
            }
        }
    }
}
