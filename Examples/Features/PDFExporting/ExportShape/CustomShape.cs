using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportShape
{
    public class CustomShape : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Open an excel file
            var fileStream = this.GetResourceStream("xlsx\\CustomShapes.xlsx");
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
                return new string[] { "xlsx\\CustomShapes.xlsx" };
            }
        }
    }
}
