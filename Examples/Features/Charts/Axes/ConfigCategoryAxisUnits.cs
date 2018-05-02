using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Charts.Axes
{
    public class ConfigCategoryAxisUnits : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A2:A6"].NumberFormat = "m/d/yyyy";
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {new DateTime(2015, 10, 7), 10, 25, 25},
                {new DateTime(2015, 10, 24), 51, 36, 27},
                {new DateTime(2015, 11, 8), 52, 85, 30},
                {new DateTime(2015, 11, 25), 22, 65, 65},
                {new DateTime(2015, 12, 10), 23, 69, 69}
            };

            GrapeCity.Documents.Excel.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Excel.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Excel.Drawing.RowCol.Columns, true, true);

            GrapeCity.Documents.Excel.Drawing.IAxis category_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Excel.Drawing.AxisType.Category);
            category_axis.MaximumScale = new DateTime(2015, 12, 20).ToOADate();
            category_axis.MinimumScale = new DateTime(2015, 10, 1).ToOADate();
            category_axis.BaseUnit = GrapeCity.Documents.Excel.Drawing.TimeUnit.Months;
            category_axis.MajorUnitScale = GrapeCity.Documents.Excel.Drawing.TimeUnit.Months;
            category_axis.MajorUnit = 1;
            category_axis.MinorUnitScale = GrapeCity.Documents.Excel.Drawing.TimeUnit.Days;
            category_axis.MinorUnit = 15;
        }


    }
}
