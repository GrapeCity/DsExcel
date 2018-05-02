using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.Text
{
    public class NumberFormating : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            sheet.Range["B3:B7"].Value = 123456.789;
            sheet.Range["B9:B13"].Value = -123456.789;

            sheet.Range["B4, B10"].NumberFormat = "0.00;[Red]0.00";
            sheet.Range["B5, B11"].NumberFormat = "$#,##0.00;[Red]$#,##0.00";
            sheet.Range["B6, B12"].NumberFormat = "0.00E+00";
            sheet.Range["B7, B13"].NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* \" - \"??_);_(@_)";

            //Other settings
            sheet.Columns[1].ColumnWidthInPixel = 100;
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
