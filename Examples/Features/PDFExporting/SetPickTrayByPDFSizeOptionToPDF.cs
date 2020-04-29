using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SetPickTrayByPDFSizeOptionToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1"].Value = "GrapeCity Documents for Excel";
            worksheet.Range["A1"].Font.Size = 25;

            // You must create a pdfSaveOptions object before using ViewerPreferences.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // The PDF page size is used to select the input paper tray when printing.
            pdfSaveOptions.ViewerPreferences.PickTrayByPDFSize = true;

            // Save the workbook into pdf file.
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

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
