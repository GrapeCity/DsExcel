using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SetDocumentPropertiesToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Value = "GrapeCity Documents for Excel";
            worksheet.Range["A1"].Font.Size = 25;

            DocumentProperties documentProperties = new DocumentProperties
            {
                //Sets the name of the person that created the PDF document.
                Author = "Jaime Smith",
                //Sets the title of the  PDF document.
                Title = "GcPdf Document Info Sample",
                //Do not embed a font.
                EmbedStandardWindowsFonts = false,
                //Set the PDF version.
                PdfVersion = 1.5f,
                //Set the subject of the PDF document.
                Subject = "GcPdfDocument.DocumentInfo",
                //Set the keyword associated with the PDF document.
                Keywords = "Keyword1",
                //Set the creation date and time of the PDF document.
                CreationDate = DateTime.Now.AddYears(10),
                //Set the date and time the PDF document was most recently modified.
                ModifyDate = DateTime.Now.AddYears(11),
                //Set the name of the application that created the original PDF document.
                Creator = "GcPdfWeb Creator",
                //Set the name of the application that created the PDF document.
                Producer = "GcPdfWeb Producer"
            };


            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                //Sets the document properties of the pdf.
                DocumentProperties = documentProperties
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
    }
}
