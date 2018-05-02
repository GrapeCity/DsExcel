using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Formatting.Borders
{
    public class AddBordersToCell : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            IRange rangeB2 = worksheet.Range["B2"];

            //set left, top, right, bottom borders together.
            rangeB2.Borders.LineStyle = BorderLineStyle.DashDot;
            rangeB2.Borders.Color = Color.Green;

            //set top border individually.
            rangeB2.Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Medium;
            rangeB2.Borders[BordersIndex.EdgeTop].Color = Color.Red;

            //set diagonal down border individually.
            rangeB2.Borders[BordersIndex.DiagonalDown].LineStyle = BorderLineStyle.Hair;
            rangeB2.Borders[BordersIndex.DiagonalDown].Color = Color.Blue;

            //set diagonal up border individually.
            rangeB2.Borders[BordersIndex.DiagonalUp].LineStyle = BorderLineStyle.Dotted;
            rangeB2.Borders[BordersIndex.DiagonalUp].Color = Color.Blue;
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool ShowScreenshot
        {
            get
            {
                return true;
            }
        }
    }
}
