using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.NewCharts
{
    public class AddBoxWhiskerChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1:D16"].Value = new object[,]
            {
                {"Course", "SchoolA", "SchoolB", "SchoolC"},
                {"English", 63, 53, 45},
                {"Physics", 61, 55, 65},
                {"English", 63, 50, 65},
                {"Math", 62, 51, 64},
                {"English", 46, 53, 66},
                {"English", 58, 56, 67},
                {"Math", 60, 51, 67},
                {"Math", 62, 53, 66},
                {"English", 63, 54, 64},
                {"English", 63, 52, 67},
                {"Physics", 60, 56, 64},
                {"English", 60, 56, 67},
                {"Math", 61, 56, 45},
                {"Math", 63, 58, 64},
                {"English", 59, 54, 65}
            };

            //Create a box&whisker chart.
            IShape shape = worksheet.Shapes.AddChart(ChartType.BoxWhisker, 300, 20, 300, 200);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D16"]);

            //Config value axis's scale.
            IAxis value_axis = shape.Chart.Axes.Item(AxisType.Value, AxisGroup.Primary);
            value_axis.MinimumScale = 40;
            value_axis.MaximumScale = 70;

            //Config the display of box&whisker plot.  
            ISeries series = shape.Chart.SeriesCollection[0];
            series.ShowInnerPoints = true;
            series.ShowOutlierPoints = false;
            series.ShowMeanMarkers = false;
            series.ShowMeanLine = true;
            series.QuartileCalculationInclusiveMedian = true;
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

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
