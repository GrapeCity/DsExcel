using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Tables
{
    public class AddCustomTableStyle : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Add one custom table style.
            ITableStyle style = workbook.TableStyles.Add("test");

            //Set WholeTable element style.
            style.TableStyleElements[TableStyleElementType.WholeTable].Font.Italic = true;
            style.TableStyleElements[TableStyleElementType.WholeTable].Font.Color = Color.White;
            style.TableStyleElements[TableStyleElementType.WholeTable].Font.Strikethrough = true;
            style.TableStyleElements[TableStyleElementType.WholeTable].Borders.LineStyle = BorderLineStyle.Dotted;
            style.TableStyleElements[TableStyleElementType.WholeTable].Borders.Color = Color.FromArgb(0, 193, 213);
            style.TableStyleElements[TableStyleElementType.WholeTable].Interior.Color = Color.FromArgb(59, 92, 170);

            //Set FirstColumnStripe element style.
            style.TableStyleElements[TableStyleElementType.FirstColumnStripe].Font.Bold = true;
            style.TableStyleElements[TableStyleElementType.FirstColumnStripe].Font.Color = Color.FromArgb(255, 0, 0);
            style.TableStyleElements[TableStyleElementType.FirstColumnStripe].Borders.LineStyle = BorderLineStyle.Thick;
            style.TableStyleElements[TableStyleElementType.FirstColumnStripe].Borders.ThemeColor = ThemeColor.Accent5;
            style.TableStyleElements[TableStyleElementType.FirstColumnStripe].Interior.Color = Color.FromArgb(255, 255, 0);
            style.TableStyleElements[TableStyleElementType.FirstColumnStripe].StripeSize = 2;

            //Set SecondColumnStripe element style.
            style.TableStyleElements[TableStyleElementType.SecondColumnStripe].Font.Color = Color.FromArgb(255, 0, 255);
            style.TableStyleElements[TableStyleElementType.SecondColumnStripe].Borders.LineStyle = BorderLineStyle.DashDot;
            style.TableStyleElements[TableStyleElementType.SecondColumnStripe].Borders.Color = Color.FromArgb(42, 105, 162);
            style.TableStyleElements[TableStyleElementType.SecondColumnStripe].Interior.Color = Color.FromArgb(204, 204, 255);

            //add table.
            IWorksheet worksheet = workbook.Worksheets[0];
            ITable table = worksheet.Tables.Add(worksheet.Range["A1:F7"], true);
            worksheet.Range["A:F"].ColumnWidth = 15;

            //set custom table style to table.
            table.TableStyle = style;
        }
    }
}
