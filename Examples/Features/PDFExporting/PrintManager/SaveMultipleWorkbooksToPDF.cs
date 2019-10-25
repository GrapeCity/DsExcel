using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PrintManager
{
    public class SaveMultipleWorkbooksToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            Stream fileStream = this.GetResourceStream("xlsx\\Any year calendar1.xlsx");
            workbook.Open(fileStream);

            GrapeCity.Documents.Excel.Workbook workbook2 = new GrapeCity.Documents.Excel.Workbook();
            Stream fileStream2 = this.GetResourceStream("xlsx\\Any year calendar (Ion theme)1.xlsx");
            workbook2.Open(fileStream2);

            //Create a printmanager.
            GrapeCity.Documents.Excel.PrintManager printManager = new GrapeCity.Documents.Excel.PrintManager();

            //Save the workbook1 and workbook2 into one pdf file.
            printManager.SavePDF(outputStream, workbook, workbook2);
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

        public override string TemplateName
        {
            get
            {
                return "Any year calendar1.xlsx";
            }
        }
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Any year calendar1.xlsx" , "xlsx\\Any year calendar (Ion theme)1.xlsx" };
            }
        }
    }
}
