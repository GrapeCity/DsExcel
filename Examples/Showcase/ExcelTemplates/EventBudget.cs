using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Showcase
{
    public class EventBudget : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Event budget.xlsx from resource
            var fileStream = this.GetResourceStream("xlsx\\Event budget.xlsx");

            workbook.Open(fileStream);

            var worksheet = workbook.ActiveSheet;

            //change range B2's font size.
            worksheet.Range["B2"].Font.Size = 22;

            //change range E4's font style to bold.
            worksheet.Range["E4"].Font.Bold = true;

            //change table style.
            worksheet.Tables["tblAdmissions"].TableStyle = workbook.TableStyles["TableStyleLight10"];
            worksheet.Tables["tblAds"].TableStyle = workbook.TableStyles["TableStyleLight10"];
            worksheet.Tables["tblVendors"].TableStyle = workbook.TableStyles["TableStyleLight10"];
            worksheet.Tables["tblItems"].TableStyle = workbook.TableStyles["TableStyleLight10"];

            //modify range F4:G5's cell style.
            worksheet.Range["F4:G5"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["F4:G5"].Interior.TintAndShade = -0.15;
            worksheet.Range["F4:G5"].Font.ThemeFont = ThemeFont.Major;
            worksheet.Range["F4:G5"].Font.Size = 12;
            worksheet.Range["F4:G5"].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.None;
            worksheet.Range["F5:G5"].NumberFormat = "$#,##0.00";

            //modify table columns' style.
            worksheet.Range["F8:G11, F15:G18, F22:G25, F29:G33"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["F8:G11, F15:G18, F22:G25, F29:G33"].Interior.TintAndShade = -0.15;
            worksheet.Range["E8:G11, E15:G18, E22:G25, E29:G33"].NumberFormat = "$#,##0.00";
        }

        public override string TemplateName
        {
            get
            {
                return "Event budget.xlsx";
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
                return new string[] { "xlsx\\Event budget.xlsx" };
            }
        }
    }
}
