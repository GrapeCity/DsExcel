using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelTemplates
{
    public class BloodSugarTracker : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Monthly business budget.xlsx from resource
            var fileStream = this.GetResourceStream("xlsx\\Blood sugar tracker.xlsx");

            workbook.Open(fileStream);

            var worksheet = workbook.ActiveSheet;

            //insert 19 rows  
            worksheet.Range["1:19"].Insert();

            //Change the rows(2~5) RowHeight
            worksheet.Rows[1].RowHeight = 34.5;
            worksheet.Rows[2].RowHeight = 15.75;
            worksheet.Rows[3].RowHeight = 19.5;
            worksheet.Rows[4].RowHeight = 15.75;

            //Set values
            worksheet.Range["B2"].Value = "BLOOD SUGAR TRACKING";
            worksheet.Range["B4"].Value = "CHARTED PROGRESS";

            //Set Styles
            worksheet.Range["B2"].Font.ThemeFont = ThemeFont.Major;
            worksheet.Range["B2"].Font.Size = 26;
            worksheet.Range["B2"].Font.ThemeColor = ThemeColor.Dark1;
            worksheet.Range["B2"].Font.TintAndShade = 0.34998626667073579;
            worksheet.Range["B2:D2"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;
            worksheet.Range["B2:D2"].Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Accent1;

            worksheet.Range["B4"].Font.ThemeFont = ThemeFont.Major;
            worksheet.Range["B4"].Font.Bold = true;
            worksheet.Range["B4"].Font.Size = 14;
            worksheet.Range["B4"].Font.ThemeColor = ThemeColor.Dark1;
            worksheet.Range["B4"].Font.TintAndShade = 0.34998626667073579;
            worksheet.Range["B4:D4"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;
            worksheet.Range["B4:D4"].Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Accent2;

            //Add chart
            IShape shape = worksheet.Shapes.AddChart(ChartType.Line, 9.75, 100.5, 365, 203.25);
            shape.Name = "BloodSugarProgress";

            //Add Series.
            ISeries series1 = shape.Chart.SeriesCollection.NewSeries();
            series1.Formula = "=SERIES('BLOOD SUGAR DATA'!$C$23,'BLOOD SUGAR DATA'!$B$24:$B$45,'BLOOD SUGAR DATA'!$C$24:$C$45,1)";
            series1.Format.Line.Color.ObjectThemeColor = ThemeColor.Accent1;
            series1.Format.Line.Weight = 2.5;

            ISeries series2 = shape.Chart.SeriesCollection.NewSeries();
            series2.Formula = "=SERIES('BLOOD SUGAR DATA'!$D$23,'BLOOD SUGAR DATA'!$B$24:$B$45,'BLOOD SUGAR DATA'!$D$24:$D$45,2)";
            series2.Format.Line.Color.ObjectThemeColor = ThemeColor.Accent2;
            series2.Format.Line.Weight = 2.5;

            //Hidden the chart title
            shape.Chart.HasTitle = false;

            //Hidden the category axis
            IAxis category_axis = shape.Chart.Axes.Item(AxisType.Category);
            category_axis.Visible = false;

            //Set value axis units
            IAxis value_axis = shape.Chart.Axes.Item(AxisType.Value);
            value_axis.MaximumScale = 140;
            value_axis.MinimumScale = 50;
            value_axis.MajorUnit = 10;
            value_axis.MinorUnit = 2;

            //Add a rectange shape
            IShape shape1 = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 385, 84.75, 102, 218.25);
            shape1.Fill.Solid();
            shape1.Fill.Color.ObjectThemeColor = ThemeColor.Accent1;
            shape1.Fill.Color.Brightness = 0.6;

            //set shape's border to no line
            shape1.Line.Color.ColorType = SolidColorType.None;

            //set shape rich text
            ITextRange shape1_p1 = shape1.TextFrame.TextRange.Paragraphs[0];
            shape1_p1.Text = "INFO:";
            shape1_p1.Runs.Add(" Blood sugar levels will vary from person-to-person.  There are many factors to keeping it within your normal range and isn't based on sugar alone.  Consult a physician for additional information or follow-up.");

            ITextRange shape1_p2 = shape1.TextFrame.TextRange.Paragraphs.Add("");
            ITextRange shape1_p3 = shape1.TextFrame.TextRange.Paragraphs.Add("More info can be found here:");
            ITextRange shape1_p4 = shape1.TextFrame.TextRange.Paragraphs.Add("http://diabetes.webmd.com/blood-glucose");

            shape1.TextFrame.TextRange.Font.Size = 10;
            shape1.TextFrame.TextRange.Font.ThemeFont = ThemeFont.Minor;
            shape1.TextFrame.TextRange.Font.Color.ObjectThemeColor = ThemeColor.Dark1;
            shape1.TextFrame.TextRange.Font.Color.Brightness = 0.25;
            shape1_p1.Runs[0].Font.Bold = true;
            shape1_p3.Runs[0].Font.Bold = true;

            //Do table filter
            var table = worksheet.Tables[0];
            table.Range.AutoFilter(1, ">=102");
        }

        public override string TemplateName
        {
            get
            {
                return "Blood sugar tracker.xlsx";
            }
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
        

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\Blood sugar tracker.xlsx" };
            }
        }
    }
}
