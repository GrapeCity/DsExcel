using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigureBestFitRowColumn : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            // Set text for some cells.
            sheet.Range["A1"].Value = "Grapecity";
            sheet.Range["A2"].Value = "Document For .NET";
            sheet.Range["B1"].Value = "Grapecity";
            sheet.Range["B2"].Value = "Excel for .NET";
            //Set font size of cell "A2"
            sheet.Range["A2"].Font.Size = 20;

            // Set bestFitColumns/bestFitRows as true.
            sheet.PageSetup.BestFitColumns = true;
            sheet.PageSetup.BestFitRows = true;

            // Set print gridline and heading.
            sheet.PageSetup.PrintGridlines = true;
            sheet.PageSetup.PrintHeadings = true;
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

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
