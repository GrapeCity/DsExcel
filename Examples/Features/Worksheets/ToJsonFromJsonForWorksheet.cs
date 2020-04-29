using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Worksheets
{
    public class ToJsonFromJsonForWorksheet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //ToJson&FromJson can be used in combination with spread.sheets product:http://spread.grapecity.com/spreadjs/sheets/

            //GrapeCity Documents for Excel import an excel file.
            //Change the path to real source file path.
            string source = System.IO.Path.Combine(this.CurrentDirectory, "source.xlsx");
            workbook.Open(source);

            //Open the same user file
            GrapeCity.Documents.Excel.Workbook new_workbook = new GrapeCity.Documents.Excel.Workbook();
            new_workbook.Open(source);

            foreach (IWorksheet worksheet in workbook.Worksheets)
            {
                //Do any change of worksheet
                //...

                //GrapeCity Documents for Excel export a worksheet to a json string.
                string json = worksheet.ToJson();
                //Use the json string to initialize spread.sheets product.
                //Product spread.sheets will show the excel file contents.

                //Use spread.sheets product export a json string of worksheet.
                //GrapeCity Documents for Excel use the json string to update content of the corresponding worksheet.
                new_workbook.Worksheets[worksheet.Name].FromJson(json);
            }
           
            //GrapeCity Documents for Excel export workbook to an excel file.
            //Change the path to real export file path.
            string export = System.IO.Path.Combine(this.CurrentDirectory, "export.xlsx");
            new_workbook.Save(export);

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

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
