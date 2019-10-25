using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SaveSheetBackgroundToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Value = "GrapeCity Documents for Excel";
            worksheet.Range["A1"].Font.Size = 25;

            Stream stream = this.GetResourceStream("logo.png");
            byte[] imageBytes = new byte[stream.Length];
            stream.Read(imageBytes, 0, imageBytes.Length);
            //Set a background image for worksheet
            worksheet.BackgroundPicture = imageBytes;

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                //Print the background image when saving pdf.
                //The background image will be centered on every page of the sheet.
                PrintBackgroundPicture = true
            };

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
                return new string[] { "logo.png" };
            }
        }
    }
}
