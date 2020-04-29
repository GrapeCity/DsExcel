using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting
{
    public class PrintTransparentCell : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook, MemoryStream outputStream)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Initialize worksheet's values.
            worksheet.Range["A1"].Value = "Info from Acme Institute of Health:";
            worksheet.Range["B2"].Value = "BLOOD PRESSURE TRACKER";
            worksheet.Range["B4:F13"].Value = new object[,] {
                { "NAME", null, null, null, "JAMES HILL" },
                { null, null, null, null, null },
                { null, null, null, "Systolic", "Diastolic" },
                { "TARGET BLOOD PRESSURE", null, null, 120, 80 },
                { null, null, null, null, null },
                { null, null, null, "Systolic", "Diastolic" },
                { "CALL PHYSICIAN IF ABOVE", null, null, 140, 90 },
                { null, null, null, null, null },
                { null, null, null, null, null },
                { "PHYSICIAN PHONE NUMBER", null, null, "(001))5104234242", null}
            };
            worksheet.Range["A1"].Font.Size = 25;

            //Set row height.
            worksheet.StandardHeight = 12.75;
            worksheet.StandardWidth = 8.43;
            worksheet.Rows[1].RowHeight = 32.25;
            worksheet.Rows[2].RowHeight = 13.5;
            worksheet.Rows[3].RowHeight = 18.75;
            worksheet.Rows[6].RowHeight = 18.75;
            worksheet.Rows[9].RowHeight = 18.75;
            worksheet.Rows[12].RowHeight = 18.75;
            worksheet.Rows[15].RowHeight = 19.5;
            worksheet.Rows[16].RowHeight = 13.5;
            worksheet.Rows[33].RowHeight = 19.5;
            worksheet.Rows[34].RowHeight = 13.5;

            //Set column width.
            worksheet.Columns[0].ColumnWidth = 1.7109375;
            worksheet.Columns[1].ColumnWidth = 12.140625;
            worksheet.Columns[2].ColumnWidth = 12.140625;
            worksheet.Columns[3].ColumnWidth = 12.140625;
            worksheet.Columns[4].ColumnWidth = 11.85546875;
            worksheet.Columns[5].ColumnWidth = 12.7109375;

            //Set the transparency value of the background color of range["A1:G20"] to 50.
            worksheet.Range["A1:G20"].Interior.Color = Color.FromArgb(50, 255, 0, 0);

            //Add a background picture.
            Stream stream = this.GetResourceStream("AcmeLogo-vertical-250px.png");
            IBackgroundPicture picture = worksheet.BackgroundPictures.AddPictureInPixel(stream, ImageType.JPG, 10, 10, 300, 150);

            //You must create a pdfSaveOptions object before using PrintTransparentCell.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            //Set print the transparency of cell's background color, so the background picture will come out in the back.
            pdfSaveOptions.PrintTransparentCell = true;

            //Save the workbook into pdf file.
            workbook.Save(outputStream, pdfSaveOptions);
        }

        public override bool SavePageInfos
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
                return new string[] { "AcmeLogo-vertical-250px.png" };
            }
        }
    }
}
