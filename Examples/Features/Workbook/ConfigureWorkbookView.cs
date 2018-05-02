using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ConfigureWorkbookView : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Workbook view settings.
            IWorkbookView bookView = workbook.BookView;
            bookView.DisplayVerticalScrollBar = false;
            bookView.DisplayWorkbookTabs = true;
            bookView.TabRatio = 0.5;

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
