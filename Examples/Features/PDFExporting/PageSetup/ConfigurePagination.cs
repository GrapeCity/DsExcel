using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.PdfPageSetup
{
    public class ConfigurePagination : ExampleBase
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

            //Set data
            sheet.Range["A1:N50"].Value = data;

            //Set paper size
            sheet.PageSetup.PaperSize = PaperSize.A5;

            //With API the margin's unit is pound, with Excel the margins display unit is inch.
            //One inch eaquals to 72 pounds. If the top margin is 0.8 inch, then please set PageSetup.TopMargin = 0.8*72(57.6);
            sheet.PageSetup.TopMargin = 57.6;
            //Top margin in excel is 0.8 inch
            sheet.PageSetup.BottomMargin = 21.6;
            sheet.PageSetup.LeftMargin = 28.8;
            sheet.PageSetup.RightMargin = 0;
            sheet.PageSetup.HeaderMargin = 0;
            sheet.PageSetup.FooterMargin = 93.6;
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
