using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class ConvertShapeToImage : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle
            IShape rectangle = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 0, 0, 191, 194);

            // Set soild fill for rectangle
            rectangle.Fill.Color.RGB = Color.OrangeRed;

            //NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            //Create a pdf document.
            GrapeCity.Documents.Pdf.GcPdfDocument doc = new GrapeCity.Documents.Pdf.GcPdfDocument();
            GrapeCity.Documents.Pdf.Page page = doc.NewPage();

            //Create a PrintManager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Get the size of the shape.
            IRange topLeftCell = rectangle.TopLeftCell;
            IRange bottomRightCell = rectangle.BottomRightCell;
            IRange shapeRange = worksheet.Range[topLeftCell.Row, topLeftCell.Column, bottomRightCell.Row - topLeftCell.Row + 1, bottomRightCell.Column - topLeftCell.Column + 1];
            SizeF size = printManager.GetSize(shapeRange);

            //Modify the size of the page.
            page.Size = size;

            // Draw the Range"A1:E5" to the specified location on the page. 
            printManager.Draw(page, new PointF(0, 0), shapeRange);

            // Saves the page as an image to a stream.
            page.SaveAsPng(outputStream, new Pdf.SaveAsImageOptions());
        }

        public override bool SaveAsImages
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
