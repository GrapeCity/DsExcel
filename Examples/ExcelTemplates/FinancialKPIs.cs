using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelTemplates
{
    public class FinancialKPIs : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Financial KPIs.xlsx from resource
            var fileStream = this.GetTemplateStream();

            workbook.Open(fileStream);

            var worksheet = workbook.ActiveSheet;

            //set values
            object[,] ProData = new object[,]
            {
                {1483550,121386 },
                {0.4336,0.32 },
                {0.1236,-0.0438 },
                {12.36,-0.0438 }
            };

            object[] ProData1 = new object[]
            {
                0.14,0.0029,0.0361,0.0361
            };

            worksheet.Range["E7:F10"].Value = ProData;
            worksheet.Range["I7:I10"].Value = ProData1;

            object[,] ActiveData = new object[,]
            {
                {"4.32 item","2.00 item" },
                {"72 days","45 days" },
                {"53 days","55 days" }
            };

            object[] ActiveData1 = new object[]
            {
                "0.45 items","-5 days","-6 days"
            };

            worksheet.Range["E12:F14"].Value = ActiveData;
            worksheet.Range["I12:I14"].Value = ActiveData1;

            object[,] EffData = new object[,]
            {
                {0.3446,0.25 },
                {0.5335,0.11 }
            };

            object[] EffData1 = new object[]
            {
                0.1245,0.1946
            };

            worksheet.Range["E16:F17"].Value = EffData;
            worksheet.Range["I16:I17"].Value = EffData1;

            object[,] LinData = new object[,]
            {
                {"0.91:1","'2:1" },
                {"0.58:1","'1:1" }
            };

            object[] LinData1 = new object[]
            {
                "'0.02:1","'0.03:1"
            };

            worksheet.Range["E19:F20"].Value = LinData;
            worksheet.Range["I19:I20"].Value = LinData1;

            object[,] GeaData = new object[,]
            {
                {-9.60,0.85 },
                {0.68,0.5 }
            };

            object[] GeaData1 = new object[]
            {
                6.65,0.0282
            };

            worksheet.Range["E22:F23"].Value = GeaData;
            worksheet.Range["I22:I23"].Value = GeaData1;

            object[,] CasData = new object[,]
            {
                {0.0735,1.2 },
                {0.1442,0.1442 }
            };

            object[] CasData1 = new object[]
            {
                -0.0046,0.023
            };

            worksheet.Range["E25:F26"].Value = CasData;
            worksheet.Range["I25:I26"].Value = CasData1;
        }

        public override string TemplateName
        {
            get
            {
                return "Financial KPIs.xlsx";
            }
        }

        public override bool HasTemplate
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

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
