using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ConfigGradientFillWithPathShape : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Open an excel file
            var fileStream = this.GetResourceStream("xlsx\\WebsiteFlowChart.xlsx");
            workbook.Open(fileStream);

            IWorksheet worksheet = workbook.Worksheets[0];

            //Get "Idea"
            IShape idea = worksheet.Shapes["Idea"];
            idea.Fill.TwoColorGradient(GrapeCity.Documents.Excel.Drawing.GradientStyle.FromCenter, 1);

            idea.Fill.GradientStops[0].Position = 0.33;
            idea.Fill.GradientStops[0].Color.RGB = System.Drawing.Color.FromArgb(0, 112, 192);

            idea.Fill.GradientStops[1].Position = 1;
            idea.Fill.GradientStops[1].Color.RGB = System.Drawing.Color.White;
            
            //Set gradient path type as "Path"
            idea.Fill.GradientPathType = PathShapeType.Path;

            //Get "Functionality"
            IShape functionality = worksheet.Shapes["Functionality"];
            functionality.Fill.TwoColorGradient(GrapeCity.Documents.Excel.Drawing.GradientStyle.FromCenter, 1);

            functionality.Fill.GradientStops[0].Position = 0.33;
            functionality.Fill.GradientStops[0].Color.RGB = System.Drawing.Color.FromArgb(0, 112, 192);

            functionality.Fill.GradientStops[1].Position = 1;
            functionality.Fill.GradientStops[1].Color.RGB = System.Drawing.Color.White;

            //Set gradient path type as "Path"
            functionality.Fill.GradientPathType = PathShapeType.Path;
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
                return new string[] { "xlsx\\WebsiteFlowChart.xlsx" };
            }
        }

    }
}
