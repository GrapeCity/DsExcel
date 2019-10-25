using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Workbook
{
    public class ToJsonFromJson : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //ToJson&FromJson can be used in combination with spread.sheets product:http://spread.grapecity.com/spreadjs/sheets/

            //GrapeCity Documents for Excel import an excel file.
            //change the path to real source file path.
            string source = System.IO.Path.Combine(this.CurrentDirectory, "source.xlsx");
            workbook.Open(source);
            //GrapeCity Documents for Excel export to a json string.
            var jsonstr = workbook.ToJson();
            //use the json string to initialize spread.sheets product.
            //spread.sheets will show the excel file contents.

            //spread.sheets product export a json string.
            //GrapeCity Documents for Excel use the json string to initialize.
            workbook.FromJson(jsonstr);
            //GrapeCity Documents for Excel export workbook to an excel file.
            //change the path to real export file path.
            string export = System.IO.Path.Combine(this.CurrentDirectory, "export.xlsx");
            workbook.Save(export);

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
    }
}
