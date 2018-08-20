using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class ConfigWorksheet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            
            //Set worksheet tab color.
            worksheet.TabColor = Color.Green;

            //Set worksheet default row height.
            worksheet.StandardHeight = 20;
            //Set worksheet default column width.
            worksheet.StandardWidth = 50;

            //Split worksheet to panes.
            worksheet.SplitPanes(worksheet.Range["B3"].Row, worksheet.Range["B3"].Column);

            IWorksheet worksheet1 = workbook.Worksheets.Add();
            //Hide worksheet.
            worksheet1.Visible = Visibility.Hidden;
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

        public override bool IsUpdate
        {
            get
            {
                return true;
            }
        }

    }
}
