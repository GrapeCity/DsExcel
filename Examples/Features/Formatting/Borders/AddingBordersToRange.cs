using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Borders
{
    public class AddingBordersToRange : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            IRange rangeB2_E6 = worksheet.Range["B2:E6"];

            //set left, top, right, bottom borders together.
            rangeB2_E6.Borders.LineStyle = BorderLineStyle.DashDot;
            rangeB2_E6.Borders.Color = Color.Green;

            //set inside horizontal border.
            rangeB2_E6.Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Dashed;
            rangeB2_E6.Borders[BordersIndex.InsideHorizontal].Color = Color.Tomato;

            //set inside vertical border.
            rangeB2_E6.Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.Double;
            rangeB2_E6.Borders[BordersIndex.InsideVertical].Color = Color.Blue;

            //set top border individually.
            rangeB2_E6.Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Medium;
            rangeB2_E6.Borders[BordersIndex.EdgeTop].Color = Color.Red;
        }
    }
}
