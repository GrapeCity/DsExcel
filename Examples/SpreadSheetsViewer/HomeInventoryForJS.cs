using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.SpreadSheetsViewer
{
    public class HomeInventoryForJS : ExampleBase
    {

        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            //Load template file
            var fileStream = this.GetTemplateStream("Home inventory.xlsx");
            workbook.Open(fileStream);
        }

        public override string TemplateName
        {
            get
            {
                return "Home inventory.xlsx";
            }
        }

        public override bool HasTemplate
        {
            get
            {
                return true;
            }
        }

        public override bool IsViewReadOnly
        {
            get
            {
                return false;
            }
        }

        public override bool ShowCode
        {
            get
            {
                return false;
            }
        }
    }
}
