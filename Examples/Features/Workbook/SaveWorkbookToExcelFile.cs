using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class SaveWorkbookToExcelFile : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //change the path to real export path when save.
            workbook.Save(this.CurrentDirectory + "dest.xlsx", SaveFileFormat.Xlsx);

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

        public override bool IsUpdate
        {
            get
            {
                return true;
            }
        }
    }
}
