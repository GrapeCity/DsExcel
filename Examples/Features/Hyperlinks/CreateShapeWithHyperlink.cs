using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Hyperlinks
{
    public class CreateShapeWithHyperlink : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            // Add shapes
            IShape shape1 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 10, 0, 100, 100);
            shape1.TextFrame.TextRange.Add("Go to google web site.");
            IShape shape2 = worksheet.Shapes.AddShape(AutoShapeType.RightArrow, 10, 120, 100, 100);
            shape2.TextFrame.TextRange.Add("Go to sheet1 C3:E4");
            IShape shape3 = worksheet.Shapes.AddShape(AutoShapeType.Oval, 10, 240, 100, 100);
            shape3.TextFrame.TextRange.Add("Send an email to sales");
            IShape shape4 = worksheet.Shapes.AddShape(AutoShapeType.LeftArrow, 10, 360, 100, 100);
            shape4.TextFrame.TextRange.Add("Link to external.xlsx file.");

            //add a hyperlink link to web page.
            worksheet.Hyperlinks.Add(shape1,
                                    "http://www.google.com/",
                                     null,
                                     "open google web site.",
                                     "Google");

            //add a hyperlink link to a range in this document.
            worksheet.Hyperlinks.Add(shape2,
                                    null,
                                    "Sheet1!$C$3:$E$4",
                                    "Go to sheet1 C3:E4");

            //add a hyperlink link to email address.
            worksheet.Hyperlinks.Add(shape3,
                                    "mailto:us.sales@grapecity.com", 
                                    null, 
                                    "Send an email to sales", 
                                    "Send an email to sales");

            //add a hyperlink link to external file.
            //change the path to real picture file path.
            string path = "external.xlsx";
            worksheet.Hyperlinks.Add(shape4, 
                                    path, 
                                    null, 
                                    "link to external.xlsx file.", 
                                    "External.xlsx");
        }

        public override bool IsNew
        {
            get
            {
                return true;
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
