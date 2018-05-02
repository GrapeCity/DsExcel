using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.ExcelTemplates
{
    public class AnnualFinancialReport:ExampleBase
    {

        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            //Load template file Annual financial report.xlsx from resource
            var fileStream = this.GetTemplateStream("Annual financial report.xlsx");

            workbook.Open(fileStream);

            var worksheet1 = workbook.Worksheets["Financial Report"];
            var worksheet2 = workbook.Worksheets["Financial Data Input"];

            //set values for sheet "Financial Report"
            object[,] MetricData = new object[,] {
                { 125000, 134137.45, 142728.38, 150687.46, 165044.56, 180026.63, 180583.88 },
                { 65000, 70962.31, 75924.86, 78901.27, 81674.37, 80883.33, 94419.45 },
                { 60000, 64207.3, 68857.69, 75643.25, 76755.26, 77317.83, 73425.99 },
                { 4500, 4517.77, 4656.92, 4974.21, 5024.11, 5068.42, 5546.88 },
                { 2500, 2745.82, 2893.11, 3136.12, 3148.53, 3338.3, 3789.47 },
                { 54000, 54761.075, 55860.81, 59747.95, 61483.59, 66272.1, 67474.85 },
                { 22000, 23920.54, 25576.74, 27498.86, 28335.67, 29424.53, 31408.25 },
                { 32000, 34943.49, 38418.53, 39895.05, 40607.73, 42438.2, 50247.68 },
            };

            object[,] MetricData2 = new object[,] {
                { 142728.38, 150687.46, 165044.56, 180026.63, 180583.88 },
                { 75924.86, 78901.27, 81674.37, 80883.33, 94419.45 },
                { 68857.69, 75643.25, 76755.26, 77317.83, 73425.99 },
                { 4656.92, 4974.21, 5024.11, 5068.42, 5546.88 },
                { 2893.11, 3136.12, 3148.53, 3338.3, 3789.47 },
                { 55860.81, 59747.95, 61483.59, 66272.1, 67474.85 },
                { 25576.74, 27498.86, 28335.67, 29424.53, 31408.25 },
                { 38418.53, 39895.05, 40607.73, 42438.2, 50247.68 },
            };

            worksheet1.Range["S8:Y15"].Value = MetricData;
            worksheet1.Range["S18:W25"].Value = MetricData2;

            //set values for sheet "Financial Data Input"
            object[,] InputData = new object[,] {
                { 125000, 134137.45, 142728.38, 150687.46, 165044.56, 180026.63, 180583.88 },
                { 65000, 70962.31, 75924.86, 78901.27, 81674.37, 80883.33, 94419.45 },
                { 60000, 64207.3, 68857.69, 75643.25, 76755.26, 77317.83, 73425.99 },
                { 4500, 4517.77, 4656.92, 4974.21, 5024.11, 5068.42, 5546.88 },
                { 2500, 2745.82, 2893.11, 3136.12, 3148.53, 3338.3, 3789.47 },
                { 54000, 54761.075, 55860.81, 59747.95, 61483.59, 66272.1, 67474.85 },
                { 22000, 23920.54, 25576.74, 27498.86, 28335.67, 29424.53, 31408.25 },
                { 32000, 34943.49, 38418.53, 39895.05, 40607.73, 42438.2, 50247.68 },
                { 12.8, 12.81, 13.78, 14.29, 15.57, 16.78, 19.96 },
                { 18.2, 18.59, 19.22, 20.17, 20.48, 21.84, 26.01 },
                { 19.1, 20.55, 21.87, 23.19, 24.67, 26.39, 31.08 },
                { 12.1, 12.21, 12.59, 13.7, 13.76, 14.59, 14.92 },
                { 0.75, 0.79, 0.85, 0.89, 0.91, 1, 1.03 },
                { 0.23, 0.25, 0.27, 0.28, 0.29, 0.3, 0.34 },
            };

            worksheet2.Range["C6:I19"].Value = InputData;
        }

        public override string TemplateName
        {
            get
            {
                return "Annual financial report.xlsx";
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
