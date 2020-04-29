using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class SetShapeTextAlignment : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\WebsiteFlowChart[Template].xlsx");
            workbook.Open(fileStream);

            IWorksheet worksheet = workbook.Worksheets[0];

            foreach (IShape shape in worksheet.Shapes)
            {
                //Centers text vertically. 
                shape.TextFrame.VerticalAnchor = VerticalAnchor.AnchorMiddle;
                //Centers text horizontally.
                shape.TextFrame.HorizontalAnchor = HorizontalAnchor.Center;
            }
        }

        public override bool ShowViewer
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
                return new string[] { "xlsx\\WebsiteFlowChart[Template].xlsx" };
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
