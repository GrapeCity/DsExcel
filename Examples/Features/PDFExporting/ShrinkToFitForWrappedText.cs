using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class ShrinkToFitForWrappedText : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.PageSetup.PrintGridlines = true;
            worksheet.PageSetup.PrintHeadings = true;

            //"A1" is ordinary wrapped text.
            worksheet.Range["A1"].WrapText = true;
            worksheet.Range["A1"].Value = "GrapeCity Documents for Excel";
            worksheet.Range["A1"].RowHeight = 42;
            worksheet.Range["A1"].ColumnWidth = 9;

            //The wrapped text "A2" will be shrink to fit.
            worksheet.Range["A2"].WrapText = true;
            worksheet.Range["A2"].ShrinkToFit = true;
            worksheet.Range["A2"].Value = "GrapeCity Documents for Excel";
            worksheet.Range["A2"].RowHeight = 32;

            //You must create a pdfSaveOptions object before using ShrinkToFitSettings.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            //Shrink the wrapped text within the cell with existing row height/column width, while exporting to PDF. 
            pdfSaveOptions.ShrinkToFitSettings.CanShrinkToFitWrappedText = true;

            //Set minimum font size with which the text should shrink.
            //pdfSaveOptions.ShrinkToFitSettings.MinimumFont = 10;
            //If after setting the minimum font size, the text is very long not fully visible, the ellipsis string to show for long text.
            //pdfSaveOptions.ShrinkToFitSettings.Ellipsis = "~";

            //Save the workbook into pdf file.
            workbook.Save(outputStream, pdfSaveOptions);
        }

        public override bool SavePageInfos
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
