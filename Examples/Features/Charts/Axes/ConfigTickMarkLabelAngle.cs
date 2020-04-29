using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.Axes
{
    public class ConfigTickMarkLabelAngle : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Open an excel file
            var fileStream = this.GetResourceStream("xlsx\\Manufacturing output chart.xlsx");
            workbook.Open(fileStream);

            IWorksheet worksheet = workbook.Worksheets[0];

            IShape shape = worksheet.Shapes[0];
            
            IAxis category_axis = shape.Chart.Axes.Item(AxisType.Category);

            //config tick label's angle
            category_axis.TickLabels.Orientation = -45;
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
                return new string[] { "xlsx\\Manufacturing output chart.xlsx" };
            }
        }
    }
}
