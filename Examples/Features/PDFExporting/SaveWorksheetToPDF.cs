using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class SaveWorksheetToPDF : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            //set style.
            sheet.Range["A1"].Value = "Sheet1";
            sheet.Range["A1"].Font.Name = "Wide Latin";
            sheet.Range["A1"].Font.Color = Color.Red;
            sheet.Range["A1"].Interior.Color = Color.Green;

            //change the path to real export path when save.
            sheet.Save(this.CurrentDirectory + "dest.pdf", SaveFileFormat.Pdf);
        }

        public override bool CanDownload
        {
            get
            {
                return false;
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
