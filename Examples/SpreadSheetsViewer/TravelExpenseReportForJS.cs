using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.SpreadSheetsViewer
{
    public class TravelExpenseReportForJS : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            //Load template file
            var fileStream = this.GetTemplateStream("Travel expense report.xlsx");
            workbook.Open(fileStream);
        }

        public override string TemplateName
        {
            get
            {
                return "Travel expense report.xlsx";
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
