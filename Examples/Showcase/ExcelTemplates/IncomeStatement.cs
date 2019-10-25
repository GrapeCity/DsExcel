using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Showcase
{
    public class IncomeStatement:ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Income statement.xlsx from resource
            var fileStream = this.GetResourceStream("xlsx\\Income statement.xlsx");

            workbook.Open(fileStream);

            var worksheet = workbook.ActiveSheet;

            //set values
            object[,] RevData = new object[,]
            {
                { 454597, 687802, 469902, 335846, 776667, 806250, 552036, 346734, 787205, 720496, 686963, 835688 },
                { 821687, 679306, 568503, 715827, 627459, 822625, 271114, 797991, 823966, 304521, 758665, 322694 },
                { 426510, 767295, 799314, 408533, 353862, 261403, 344075, 277433, 289814, 386539, 625447, 489178 },
            };

            object[,] CostData = new object[,]
            {
                { 113053, 134668, 192784, 199522, 163169, 109389, 187568, 170346, 130920, 180119, 184725, 162780 },
                { 42667, 14921, 42007, 10264, 48864, 10360, 42378, 31445, 15208, 34418, 37529, 22077 },
            };

            object[,] OpeData = new object[,]
            {
                { 264290, 282301, 266021, 272333, 285359, 256072, 289430, 280307, 278656, 272674, 299784, 284062 },
                { 29536, 28709, 27579, 27862, 26750, 26970, 27847, 27680, 28375, 25474, 27514, 26753 },
                { 23473, 22902, 21680, 19232, 23077, 20839, 19461, 15504, 21117, 20102, 16322, 21160 },
                { 11340, 11514, 11207, 9186, 9243, 10003, 11092, 10761, 9032, 10366, 9571, 11774 },
                { 9562, 10787, 11809, 9949, 9559, 9753, 11761, 10702, 10545, 10347, 10593, 10719 },
                { 9410, 9646, 9969, 10758, 9183, 10777, 11249, 9763, 9445, 11218, 10579, 10203 },
            };

            object[,] EarnData = new object[,]
            {
                { 1038, 1430, 627, 1352, 765, 1190, 1082, 1261, 1240, 645, 552, 669 },
                { 1227, 1372, 544, 1172, 1479, 821, 891, 1361, 1187, 760, 1126, 873 },
                { 661, 1342, 1119, 1085, 887, 1274, 883, 582, 601, 845, 948, 817 },
            };

            worksheet.Range["E10:P12"].Value = RevData;
            worksheet.Range["E15:P16"].Value = CostData;
            worksheet.Range["E21:P26"].Value = OpeData;
            worksheet.Range["E31:P33"].Value = EarnData;

            //set formulars
            worksheet.Range["E13:P13"].Formula = "=SUM(E10:E12)";
            worksheet.Range["E17:P17"].Formula = "=SUM(E15:E16)";
            worksheet.Range["E19:P19"].Formula = "=E13-E17";
            worksheet.Range["E27:P27"].Formula = "=SUM(E21:E26)";
            worksheet.Range["E29:P29"].Formula = "=E19-E27";
            worksheet.Range["E34:P34"].Formula = "=E29-E31+E32-E33";
            worksheet.Range["E36:P36"].Formula = "=E34*0.17";
            worksheet.Range["E38:P38"].Formula = "=E34-E36";
        }

        public override string TemplateName
        {
            get
            {
                return "Income statement.xlsx";
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
                return new string[] { "xlsx\\Income statement.xlsx" };
            }
        }
    }
}
