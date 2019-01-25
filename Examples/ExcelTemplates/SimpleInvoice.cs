using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelTemplates
{
    public class SimpleInvoice : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Simple invoice.xlsx from resource
            var fileStream = this.GetResourceStream("xlsx\\Simple invoice.xlsx");

            workbook.Open(fileStream);

            var worksheet = workbook.ActiveSheet;

            // fill some new items
            worksheet.Range["E09:H09"].Value = new object[] { "DD1-001", "Item 3", 5.60, 12 };
            worksheet.Range["E10:H10"].Value = new object[] { "DD2-001", "Item 3", 8.5, 14 };
            worksheet.Range["E11:H11"].Value = new object[] { "DD3-001", "Item 3", 9.6, 16 };
        }

        public override string TemplateName
        {
            get
            {
                return "Simple invoice.xlsx";
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        
        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Simple invoice.xlsx" };
            }
        }
    }
}
