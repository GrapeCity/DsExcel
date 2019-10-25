using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.RangeOperations
{
    public class CopyPasteOptions : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set data of PC
            worksheet.Range["A2"].Value = "PC";
            worksheet.Range["A4:C4"].Value = new string[] {"Device", "Quantity", "Unit Price" };
            worksheet.Range["A5:C10"].Value = new object[,] {
                { "T540p", 12, 9850 },
                { "T570", 5, 7460 },
                { "Y460", 6, 5400 },
                { "Y460F", 8, 6240 }
            };

            //Set style
            worksheet.Range["A2"].RowHeight = 30;
            worksheet.Range["A2"].Font.Size = 20;
            worksheet.Range["A2"].Font.Bold = true;
            worksheet.Range["A4:C4"].Font.Bold = true;
            worksheet.Range["A4:C4"].Font.Color = Color.White;
            worksheet.Range["A4:C4"].Interior.Color = Color.LightBlue;
            worksheet.Range["A5:C10"].Borders[BordersIndex.InsideHorizontal].Color = Color.Orange;
            worksheet.Range["A5:C10"].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.DashDot;

            //Copy only style and row height
            worksheet.Range["H1"].Value = "Copy style and row height from previous cells.";
            worksheet.Range["H1"].Font.Color = Color.Red;
            worksheet.Range["H1"].Font.Bold = true;
            worksheet.Range["A2:C10"].Copy(worksheet.Range["H2"], PasteType.Formats);

            //Set data of mobile devices
            worksheet.Range["H2"].Value = "Mobile";
            worksheet.Range["H4:J4"].Value = new string[] { "Device", "Quantity", "Unit Price" };
            worksheet.Range["H5:J10"].Value = new object[,] {
                { "HW-P30", 20, 4200 },
                { "IPhone-X", 5, 9888 },
                { "IPhone-6s plus", 15, 6880 }
            };

            //Add new sheet
            IWorksheet worksheet2 = workbook.Worksheets.Add();

            //Copy only style to new sheet
            worksheet.Range["A2:C10"].Copy(worksheet2.Range["A2"], PasteType.Formats);
            worksheet2.Range["A3"].Value = "Copy style from sheet1.";
            worksheet2.Range["A3"].Font.Color = Color.Red;
            worksheet2.Range["A3"].Font.Bold = true;
        }
    }
}
