using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class ConfigureWorksheetView : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Worksheet view settings.
            IWorksheetView sheetView = worksheet.SheetView;
            sheetView.DisplayFormulas = false;
            sheetView.DisplayRightToLeft = true;
            sheetView.GridlineColor = Color.Red;
            sheetView.Zoom = 200;

        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool ShowScreenshot
        {
            get
            {
                return true;
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
