using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Grouping
{
    public class OutlineColumn : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set data.
            object[,] data = new object[,]{
                { "Preface",                                   "1",      1,  0 } ,
                { "Java SE5 and SE6",                          "1.1",    2,  1 },
                { "Java SE6",                                  "1.1.1",  2,  2 },
                { "The 4th edition",                           "1.2",    2,  1 },
                { "Changes",                                   "1.2.1",  3,  2 },
                { "Note on the cover design",                  "1.3",    4,  1 },
                { "Acknowledgements",                          "1.4",    4,  1 },
                { "Introduction",                              "2",      9,  0 },
                { "Prerequisites",                             "2.1",    9,  1 },
                { "Learning Java",                             "2.2",    10, 1 },
                { "Goals",                                     "2.3",    10, 1 },
                { "Teaching from this book",                   "2.4",    11, 1 },
                { "JDK HTML documentation",                    "2.5",    11, 1 },
                { "Exercises",                                 "2.6",    12, 1 },
                { "Foundations for Java",                      "2.7",    12, 1 },
                { "Source code",                               "2.8",    12, 1 },
                { "Coding standards",                          "2.8.1",  14, 2 },
                { "Errors",                                    "2.9",    14, 1 },
                { "Introduction to Objects",                   "3",      15, 0 },
                { "The progress of abstraction",               "3.1",    15, 1 },
                { "An object has an interface",                "3.2",    17, 1 },
                { "An object provides services",               "3.3",    18, 1 },
                { "The hidden implementation",                 "3.4",    19, 1 },
                { "Reusing the implementation",                "3.5",    20, 1 },
                { "Inheritance",                               "3.6",    21, 1 },
                { "Is-a vs. is-like-a relationships",          "3.6.1",  24, 2 },
                { "Interchangeable objects with polymorphism", "3.7",    25, 1 },
                { "The singly rooted hierarchy",               "3.8",    28, 1 },
                { "Containers",                                "3.9",    28, 1 },
                { "Parameterized types (Generics)",            "3.10",   29, 1 },
                { "Object creation & lifetime",                "3.11",   30, 1 },
                { "Exception handling: dealing with errors",   "3.12",   31, 1 },
                { "Concurrent programming",                    "3.13",   32, 1 },
                { "Java and the Internet",                     "3.14",   33, 1 },
                { "What is the Web?",                          "3.14.1", 33, 2 },
                { "Client-side programming",                   "3.14.2", 34, 2 },
                { "Server-side programming",                   "3.14.3", 38, 2 },
                { "Summary",                                   "3.15",   38, 1 }
            };
            worksheet.Range["A1:C38"].Value = data;

            //Set ColumnWidth.
            worksheet.Range["A:A"].ColumnWidthInPixel = 310;
            worksheet.Range["B:C"].ColumnWidthInPixel = 150;

            //Set IndentLevel.
            for (int i = 0; i < data.GetLength(0); i++)
            {
                worksheet.Range[i, 0].IndentLevel = (int)data[i, 3];
            }

            //Show the summary row above the detail rows.
            worksheet.Outline.SummaryRow = SummaryRow.Above;

            //Don't show the row outline when interacting with SJS, the exported excel file still show the row outline.
            worksheet.ShowRowOutline = false;

            //Set outline column, the corresponding row outlines will also be automatically created.
            worksheet.OutlineColumn.ColumnIndex = 0;
            worksheet.OutlineColumn.ShowCheckBox = true;
            worksheet.OutlineColumn.ShowImage = true;
            worksheet.OutlineColumn.MaxLevel = 2;
            worksheet.OutlineColumn.Images.Add(new ImageSource(this.GetResourceStream("archiverFolder.png"), ImageType.PNG));
            worksheet.OutlineColumn.Images.Add(new ImageSource(this.GetResourceStream("newFloder.png"), ImageType.PNG));
            worksheet.OutlineColumn.Images.Add(new ImageSource(this.GetResourceStream("docFile.png"), ImageType.PNG));
            worksheet.OutlineColumn.CollapseIndicator = new ImageSource(this.GetResourceStream("decreaseIndicator.png"), ImageType.PNG);
            worksheet.OutlineColumn.ExpandIndicator = new ImageSource(this.GetResourceStream("increaseIndicator.png"), ImageType.PNG);
            worksheet.OutlineColumn.SetCheckStatus(0, true);
            worksheet.OutlineColumn.SetCollapsed(1, true);

            //MSExcel does not support the outline column, so when exporting to the excel file, 
            //The checkbox, level images, expand&collapse images are not visible.
            //But the data is seen with heirarchical structure. 
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "archiverFolder.png", "newFloder.png", "docFile.png", "increaseIndicator.png", "decreaseIndicator.png" };
            }
        }
    }
}
