using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.Text
{
    public class VerticalText : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet sheet = workbook.Worksheets[0];

            sheet.Range["A1"].Font.Name = "@Meiryo";
            sheet.Range["A1"].Value = "日本語（にほんご、にっぽんご）は、主として、日本列島で使用されてきた言語である。GrapeCity";
            sheet.Range["A1"].HorizontalAlignment = HorizontalAlignment.Right;
            sheet.Range["A1"].VerticalAlignment = VerticalAlignment.Top;
            sheet.Range["A1"].Orientation = -90;
            sheet.Range["A1"].WrapText = true;

            sheet.Range["A1"].ColumnWidth = 27;
            sheet.Range["A1"].RowHeight = 190;
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
