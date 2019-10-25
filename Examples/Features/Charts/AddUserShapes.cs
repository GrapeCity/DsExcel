using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts
{
    public class AddUserShapes : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1:C10"].Value = new object[,]
            {
                {"Task", "Worker 1", "Worker 2"},
                {"Task 1", 7, 10},
                {"Task 2", 5, 1},
                {"Task 3", 3, 6},
                {"Task 4", 10, 5},
                {"Task 5", 4, 4},
                {"Task 6", 5, 8},
                {"Task 7", 8, 7},
                {"Task 8", 2, 5},
                {"Task 9", 6, 4}
            };

            //Add a line chart.
            IShape shape = worksheet.Shapes.AddChart(ChartType.Line, 250, 20, 400, 250);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:C10"]);
            shape.Chart.HasLegend = false;
            shape.Chart.ChartTitle.Text = " ";

            //Add a bussiness logo in the line chart area
            System.IO.Stream stream = this.GetResourceStream("logo.png");
            shape.Chart.AddPicture(stream, ImageType.PNG, 170, 10, 60, 10);

            //Add shapes in the line chart area 
            IShape userShape1 = shape.Chart.AddShape(AutoShapeType.Rectangle, 30, 45, 60, 20);
            userShape1.Fill.Color.ObjectThemeColor = ThemeColor.Accent2;
            userShape1.Line.Color.ObjectThemeColor = ThemeColor.Accent2;
            userShape1.TextFrame.TextRange[0].Text = "Worker 2";

            IShape userShape2 = shape.Chart.AddShape(AutoShapeType.Rectangle, 330, 110, 60, 20);
            userShape2.Fill.Color.ObjectThemeColor = ThemeColor.Accent1;
            userShape2.Line.Color.ObjectThemeColor = ThemeColor.Accent1;
            userShape2.TextFrame.TextRange[0].Text = "Worker 1";
        }

        public override bool CanDownload
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

        public override bool ShowScreenshot
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
                return new string[] { "logo.png" };
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
