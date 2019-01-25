using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigurePrintArea : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            int row = 50;
            int column = 14;
            object[,] data = new object[row, column];
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    data[i, j] = "R" + i + "C" + j;
                }
            }

            //Set data.
            sheet.Range["A1:N50"].Value = data;
            sheet.Range["C10:H20"].Value = "PrintArea";
            sheet.Range["C10:H20"].Interior.Color = Color.Yellow;
            sheet.PageSetup.PrintHeadings = true;

            //Set print area.
            sheet.PageSetup.PrintArea = "$C$10:$H$20";
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
