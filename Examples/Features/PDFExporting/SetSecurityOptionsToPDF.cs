using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SetSecurityOptionsToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Value = "GrapeCity Documents for Excel";
            worksheet.Range["A1"].Font.Size = 25;

            //The security settings of pdf when converting excel to pdf.
            PdfSecurityOptions securityOptions = new PdfSecurityOptions
            {
                //Sets the user password.
                UserPassword = "user",
                //Sets the owner password.
                OwnerPassword = "owner",
                //Allow to print pdf document.
                PrintPermission = true,
                //Print the pdf document in high quality.
                FullQualityPrintPermission = true,
                //Allow to copy or extract the content of the pdf document.
                ExtractContentPermission = true,
                //Allow to modify the pdf document.
                ModifyDocumentPermission = true,
                //Allow to insert, rotate, or delete pages and create bookmarks or thumbnail images of the pdf document.
                AssembleDocumentPermission = true,
                //Allow to modify text annotations and fill the form fields of the pdf document.
                ModifyAnnotationsPermission = true,
                //Filling the form fields of the pdf document is not allowed.
                FillFormsPermission = false
            };


            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                //Sets the secutity settings of the pdf.
                SecurityOptions = securityOptions
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
    }
}
