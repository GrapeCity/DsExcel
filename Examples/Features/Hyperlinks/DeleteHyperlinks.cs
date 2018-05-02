using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Hyperlinks
{
    public class DeleteHyperlinks : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A:A"].ColumnWidth = 30;

            //add a hyperlink link to web page.
            worksheet.Range["A1:B2"].Hyperlinks.Add(worksheet.Range["A1"],
                                                "http://www.google.com/",
                                                null,
                                                "open google web site.",
                                                "Google");

            //add a hyperlink link to a range in this document.
            worksheet.Range["A3:B4"].Hyperlinks.Add(worksheet.Range["A3"],
                                            null,
                                            "Sheet1!$C$3:$E$4",
                                            "Go to sheet1 C3:E4");

            //add a hyperlink link to email address.
            worksheet.Range["A5:B6"].Hyperlinks.Add(worksheet.Range["A5"],
                                            "mailto:spread.sales@grapecity.com",
                                            null,
                                            "Send an email to spread sales",
                                            "Send an eamil to spread sales");

            //add a hyperlink link to external file.
            //change the path to real picture file path.
            string path = this.CurrentDirectory + "external.xlsx";
            worksheet.Range["A7:B8"].Hyperlinks.Add(worksheet.Range["A7"],
                                                path,
                                                null,
                                                "link to external.xlsx file.",
                                                "External.xlsx");

            //delete hyperlinks in range A1:A2.
            worksheet.Range["A1:A2"].Hyperlinks.Delete();

            //delete all hyperlinks in this worksheet.
            worksheet.Hyperlinks.Delete();
        }
    }
}
