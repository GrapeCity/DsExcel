using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportChart
{
    public class ChartGalleryPdf : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Open an excel file
            var fileStream = this.GetResourceStream("xlsx\\Chart Gallery.xlsx");
            workbook.Open(fileStream);
        }

        public override bool SavePdf
        {
            get
            {
                return true;
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Chart Gallery.xlsx" };
            }
        }
    }
}
