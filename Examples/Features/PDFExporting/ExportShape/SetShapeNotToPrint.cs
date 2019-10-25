using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ExportShape
{
    public class SetShapeNotToPrint : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            // Add a rectangle
            IShape rectangle = sheet.Shapes.AddShape(AutoShapeType.Rectangle, 20, 15, 100, 100);
            
            // Add an oval
            IShape oval = sheet.Shapes.AddShape(AutoShapeType.Oval, 160, 15, 100, 100);

            //set oval not to print
            oval.IsPrintable = false;
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
    }
}
