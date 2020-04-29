namespace GrapeCity.Documents.Excel.Examples.Features.Shape
{
    public class ShapeUngroup : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            Drawing.IShapes shapes = worksheet.Shapes;
            Drawing.IShape pentagon = shapes.AddShape(Drawing.AutoShapeType.RegularPentagon, 89.4, 57.0, 153.6, 90.6);
            Drawing.IShape pie = shapes.AddShape(Drawing.AutoShapeType.Pie, 344.4, 156.8, 50.4, 60.0);
            Drawing.IShapeRange shpRange = shapes.Range[new[] { pentagon.Name, pie.Name }];

            // Group the shape range
            Drawing.IShape grouped = shpRange.Group();

            // Ungroup the group shape
            shpRange = grouped.Ungroup();
        }

        public override bool ShowViewer
        {
            get
            {
                return true;
            }
        }
    }
}
