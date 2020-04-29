using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateSamples
{
    public class EmployeeAbsenceSchedule : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_EmployeeAbsenceSchedule.xlsx");
            workbook.Open(templateFile);

            #region Init Data
            var datasource = new DataTable();
            datasource.Columns.Add(new DataColumn("Month", typeof(string)));
            datasource.Columns.Add(new DataColumn("Day", typeof(Int32)));
            datasource.Columns.Add(new DataColumn("Name", typeof(string)));
            datasource.Columns.Add(new DataColumn("AbsenceType", typeof(string)));

            datasource.Rows.Add("January", 1, "Mark Walter", null);
            datasource.Rows.Add("January", 2, "Mark Walter", null);
            datasource.Rows.Add("January", 3, "Mark Walter", "V");
            datasource.Rows.Add("January", 4, "Mark Walter", "V");
            datasource.Rows.Add("January", 5, "Mark Walter", "V");
            datasource.Rows.Add("January", 6, "Mark Walter", "V");
            datasource.Rows.Add("January", 7, "Mark Walter", null);
            datasource.Rows.Add("January", 8, "Mark Walter", null);
            datasource.Rows.Add("January", 9, "Mark Walter", null);
            datasource.Rows.Add("January", 10, "Mark Walter", null);
            datasource.Rows.Add("January", 11, "Mark Walter", null);
            datasource.Rows.Add("January", 12, "Mark Walter", null);
            datasource.Rows.Add("January", 13, "Mark Walter", "V");
            datasource.Rows.Add("January", 14, "Mark Walter", null);
            datasource.Rows.Add("January", 15, "Mark Walter", null);
            datasource.Rows.Add("January", 16, "Mark Walter", null);
            datasource.Rows.Add("January", 17, "Mark Walter", null);
            datasource.Rows.Add("January", 18, "Mark Walter", null);
            datasource.Rows.Add("January", 19, "Mark Walter", null);
            datasource.Rows.Add("January", 20, "Mark Walter", null);
            datasource.Rows.Add("January", 21, "Mark Walter", null);
            datasource.Rows.Add("January", 22, "Mark Walter", null);
            datasource.Rows.Add("January", 23, "Mark Walter", null);
            datasource.Rows.Add("January", 24, "Mark Walter", null);
            datasource.Rows.Add("January", 25, "Mark Walter", null);
            datasource.Rows.Add("January", 26, "Mark Walter", null);
            datasource.Rows.Add("January", 27, "Mark Walter", null);
            datasource.Rows.Add("January", 28, "Mark Walter", null);
            datasource.Rows.Add("January", 29, "Mark Walter", null);
            datasource.Rows.Add("January", 30, "Mark Walter", null);
            datasource.Rows.Add("January", 31, "Mark Walter", null);
            datasource.Rows.Add("January", 1, "James Collins", null);
            datasource.Rows.Add("January", 2, "James Collins", null);
            datasource.Rows.Add("January", 3, "James Collins", null);
            datasource.Rows.Add("January", 4, "James Collins", null);
            datasource.Rows.Add("January", 5, "James Collins", "S");
            datasource.Rows.Add("January", 6, "James Collins", "S");
            datasource.Rows.Add("January", 7, "James Collins", null);
            datasource.Rows.Add("January", 8, "James Collins", null);
            datasource.Rows.Add("January", 9, "James Collins", null);
            datasource.Rows.Add("January", 10, "James Collins", null);
            datasource.Rows.Add("January", 11, "James Collins", "P");
            datasource.Rows.Add("January", 12, "James Collins", null);
            datasource.Rows.Add("January", 13, "James Collins", null);
            datasource.Rows.Add("January", 14, "James Collins", null);
            datasource.Rows.Add("January", 15, "James Collins", null);
            datasource.Rows.Add("January", 16, "James Collins", null);
            datasource.Rows.Add("January", 17, "James Collins", null);
            datasource.Rows.Add("January", 18, "James Collins", null);
            datasource.Rows.Add("January", 19, "James Collins", null);
            datasource.Rows.Add("January", 20, "James Collins", "S");
            datasource.Rows.Add("January", 21, "James Collins", null);
            datasource.Rows.Add("January", 22, "James Collins", null);
            datasource.Rows.Add("January", 23, "James Collins", null);
            datasource.Rows.Add("January", 24, "James Collins", null);
            datasource.Rows.Add("January", 25, "James Collins", "V");
            datasource.Rows.Add("January", 26, "James Collins", "V");
            datasource.Rows.Add("January", 27, "James Collins", "V");
            datasource.Rows.Add("January", 28, "James Collins", null);
            datasource.Rows.Add("January", 29, "James Collins", null);
            datasource.Rows.Add("January", 30, "James Collins", null);
            datasource.Rows.Add("January", 31, "James Collins", null);
            datasource.Rows.Add("January", 1, "Andrew Fuller", null);
            datasource.Rows.Add("January", 2, "Andrew Fuller", null);
            datasource.Rows.Add("January", 3, "Andrew Fuller", "P");
            datasource.Rows.Add("January", 4, "Andrew Fuller", null);
            datasource.Rows.Add("January", 5, "Andrew Fuller", null);
            datasource.Rows.Add("January", 6, "Andrew Fuller", null);
            datasource.Rows.Add("January", 7, "Andrew Fuller", null);
            datasource.Rows.Add("January", 8, "Andrew Fuller", null);
            datasource.Rows.Add("January", 9, "Andrew Fuller", null);
            datasource.Rows.Add("January", 10, "Andrew Fuller", null);
            datasource.Rows.Add("January", 11, "Andrew Fuller", null);
            datasource.Rows.Add("January", 12, "Andrew Fuller", null);
            datasource.Rows.Add("January", 13, "Andrew Fuller", null);
            datasource.Rows.Add("January", 14, "Andrew Fuller", "S");
            datasource.Rows.Add("January", 15, "Andrew Fuller", null);
            datasource.Rows.Add("January", 16, "Andrew Fuller", null);
            datasource.Rows.Add("January", 17, "Andrew Fuller", null);
            datasource.Rows.Add("January", 18, "Andrew Fuller", null);
            datasource.Rows.Add("January", 19, "Andrew Fuller", null);
            datasource.Rows.Add("January", 20, "Andrew Fuller", null);
            datasource.Rows.Add("January", 21, "Andrew Fuller", null);
            datasource.Rows.Add("January", 22, "Andrew Fuller", null);
            datasource.Rows.Add("January", 23, "Andrew Fuller", null);
            datasource.Rows.Add("January", 24, "Andrew Fuller", null);
            datasource.Rows.Add("January", 25, "Andrew Fuller", null);
            datasource.Rows.Add("January", 26, "Andrew Fuller", null);
            datasource.Rows.Add("January", 27, "Andrew Fuller", null);
            datasource.Rows.Add("January", 28, "Andrew Fuller", null);
            datasource.Rows.Add("January", 29, "Andrew Fuller", "S");
            datasource.Rows.Add("January", 30, "Andrew Fuller", null);
            datasource.Rows.Add("January", 31, "Andrew Fuller", null);
            datasource.Rows.Add("January", 1, "Kiara Davidson", null);
            datasource.Rows.Add("January", 2, "Kiara Davidson", null);
            datasource.Rows.Add("January", 3, "Kiara Davidson", null);
            datasource.Rows.Add("January", 4, "Kiara Davidson", null);
            datasource.Rows.Add("January", 5, "Kiara Davidson", null);
            datasource.Rows.Add("January", 6, "Kiara Davidson", null);
            datasource.Rows.Add("January", 7, "Kiara Davidson", "P");
            datasource.Rows.Add("January", 8, "Kiara Davidson", null);
            datasource.Rows.Add("January", 9, "Kiara Davidson", null);
            datasource.Rows.Add("January", 10, "Kiara Davidson", null);
            datasource.Rows.Add("January", 11, "Kiara Davidson", null);
            datasource.Rows.Add("January", 12, "Kiara Davidson", null);
            datasource.Rows.Add("January", 13, "Kiara Davidson", null);
            datasource.Rows.Add("January", 14, "Kiara Davidson", null);
            datasource.Rows.Add("January", 15, "Kiara Davidson", null);
            datasource.Rows.Add("January", 16, "Kiara Davidson", null);
            datasource.Rows.Add("January", 17, "Kiara Davidson", null);
            datasource.Rows.Add("January", 18, "Kiara Davidson", null);
            datasource.Rows.Add("January", 19, "Kiara Davidson", "V");
            datasource.Rows.Add("January", 20, "Kiara Davidson", "V");
            datasource.Rows.Add("January", 21, "Kiara Davidson", "V");
            datasource.Rows.Add("January", 22, "Kiara Davidson", null);
            datasource.Rows.Add("January", 23, "Kiara Davidson", null);
            datasource.Rows.Add("January", 24, "Kiara Davidson", null);
            datasource.Rows.Add("January", 25, "Kiara Davidson", null);
            datasource.Rows.Add("January", 26, "Kiara Davidson", null);
            datasource.Rows.Add("January", 27, "Kiara Davidson", null);
            datasource.Rows.Add("January", 28, "Kiara Davidson", null);
            datasource.Rows.Add("January", 29, "Kiara Davidson", null);
            datasource.Rows.Add("January", 30, "Kiara Davidson", null);
            datasource.Rows.Add("January", 31, "Kiara Davidson", null);
            datasource.Rows.Add("January", 1, "Edward Williams", null);
            datasource.Rows.Add("January", 2, "Edward Williams", null);
            datasource.Rows.Add("January", 3, "Edward Williams", null);
            datasource.Rows.Add("January", 4, "Edward Williams", "S");
            datasource.Rows.Add("January", 5, "Edward Williams", "V");
            datasource.Rows.Add("January", 6, "Edward Williams", "V");
            datasource.Rows.Add("January", 7, "Edward Williams", null);
            datasource.Rows.Add("January", 8, "Edward Williams", null);
            datasource.Rows.Add("January", 9, "Edward Williams", null);
            datasource.Rows.Add("January", 10, "Edward Williams", null);
            datasource.Rows.Add("January", 11, "Edward Williams", null);
            datasource.Rows.Add("January", 12, "Edward Williams", null);
            datasource.Rows.Add("January", 13, "Edward Williams", null);
            datasource.Rows.Add("January", 14, "Edward Williams", null);
            datasource.Rows.Add("January", 15, "Edward Williams", null);
            datasource.Rows.Add("January", 16, "Edward Williams", null);
            datasource.Rows.Add("January", 17, "Edward Williams", "S");
            datasource.Rows.Add("January", 18, "Edward Williams", null);
            datasource.Rows.Add("January", 19, "Edward Williams", null);
            datasource.Rows.Add("January", 20, "Edward Williams", null);
            datasource.Rows.Add("January", 21, "Edward Williams", null);
            datasource.Rows.Add("January", 22, "Edward Williams", null);
            datasource.Rows.Add("January", 23, "Edward Williams", null);
            datasource.Rows.Add("January", 24, "Edward Williams", "S");
            datasource.Rows.Add("January", 25, "Edward Williams", null);
            datasource.Rows.Add("January", 26, "Edward Williams", null);
            datasource.Rows.Add("January", 27, "Edward Williams", null);
            datasource.Rows.Add("January", 28, "Edward Williams", null);
            datasource.Rows.Add("January", 29, "Edward Williams", null);
            datasource.Rows.Add("January", 30, "Edward Williams", null);
            datasource.Rows.Add("January", 31, "Edward Williams", "V");

            datasource.Rows.Add("February", 1, "Mark Walter", null);
            datasource.Rows.Add("February", 2, "Mark Walter", null);
            datasource.Rows.Add("February", 3, "Mark Walter", "V");
            datasource.Rows.Add("February", 4, "Mark Walter", "V");
            datasource.Rows.Add("February", 5, "Mark Walter", null);
            datasource.Rows.Add("February", 6, "Mark Walter", null);
            datasource.Rows.Add("February", 7, "Mark Walter", null);
            datasource.Rows.Add("February", 8, "Mark Walter", null);
            datasource.Rows.Add("February", 9, "Mark Walter", null);
            datasource.Rows.Add("February", 10, "Mark Walter", null);
            datasource.Rows.Add("February", 11, "Mark Walter", "P");
            datasource.Rows.Add("February", 12, "Mark Walter", null);
            datasource.Rows.Add("February", 13, "Mark Walter", "V");
            datasource.Rows.Add("February", 14, "Mark Walter", null);
            datasource.Rows.Add("February", 15, "Mark Walter", null);
            datasource.Rows.Add("February", 16, "Mark Walter", null);
            datasource.Rows.Add("February", 17, "Mark Walter", null);
            datasource.Rows.Add("February", 18, "Mark Walter", null);
            datasource.Rows.Add("February", 19, "Mark Walter", null);
            datasource.Rows.Add("February", 20, "Mark Walter", null);
            datasource.Rows.Add("February", 21, "Mark Walter", null);
            datasource.Rows.Add("February", 22, "Mark Walter", "S");
            datasource.Rows.Add("February", 23, "Mark Walter", null);
            datasource.Rows.Add("February", 24, "Mark Walter", null);
            datasource.Rows.Add("February", 25, "Mark Walter", null);
            datasource.Rows.Add("February", 26, "Mark Walter", null);
            datasource.Rows.Add("February", 27, "Mark Walter", null);
            datasource.Rows.Add("February", 28, "Mark Walter", null);
            datasource.Rows.Add("February", 1, "James Collins", null);
            datasource.Rows.Add("February", 2, "James Collins", null);
            datasource.Rows.Add("February", 3, "James Collins", "S");
            datasource.Rows.Add("February", 4, "James Collins", null);
            datasource.Rows.Add("February", 5, "James Collins", null);
            datasource.Rows.Add("February", 6, "James Collins", null);
            datasource.Rows.Add("February", 7, "James Collins", null);
            datasource.Rows.Add("February", 8, "James Collins", null);
            datasource.Rows.Add("February", 9, "James Collins", null);
            datasource.Rows.Add("February", 10, "James Collins", null);
            datasource.Rows.Add("February", 11, "James Collins", "V");
            datasource.Rows.Add("February", 12, "James Collins", null);
            datasource.Rows.Add("February", 13, "James Collins", null);
            datasource.Rows.Add("February", 14, "James Collins", null);
            datasource.Rows.Add("February", 15, "James Collins", null);
            datasource.Rows.Add("February", 16, "James Collins", null);
            datasource.Rows.Add("February", 17, "James Collins", null);
            datasource.Rows.Add("February", 18, "James Collins", null);
            datasource.Rows.Add("February", 19, "James Collins", null);
            datasource.Rows.Add("February", 20, "James Collins", "S");
            datasource.Rows.Add("February", 21, "James Collins", null);
            datasource.Rows.Add("February", 22, "James Collins", null);
            datasource.Rows.Add("February", 23, "James Collins", null);
            datasource.Rows.Add("February", 24, "James Collins", null);
            datasource.Rows.Add("February", 25, "James Collins", "P");
            datasource.Rows.Add("February", 26, "James Collins", null);
            datasource.Rows.Add("February", 27, "James Collins", null);
            datasource.Rows.Add("February", 28, "James Collins", null);
            datasource.Rows.Add("February", 1, "Andrew Fuller", null);
            datasource.Rows.Add("February", 2, "Andrew Fuller", null);
            datasource.Rows.Add("February", 3, "Andrew Fuller", "P");
            datasource.Rows.Add("February", 4, "Andrew Fuller", null);
            datasource.Rows.Add("February", 5, "Andrew Fuller", null);
            datasource.Rows.Add("February", 6, "Andrew Fuller", null);
            datasource.Rows.Add("February", 7, "Andrew Fuller", null);
            datasource.Rows.Add("February", 8, "Andrew Fuller", "V");
            datasource.Rows.Add("February", 9, "Andrew Fuller", null);
            datasource.Rows.Add("February", 10, "Andrew Fuller", null);
            datasource.Rows.Add("February", 11, "Andrew Fuller", null);
            datasource.Rows.Add("February", 12, "Andrew Fuller", null);
            datasource.Rows.Add("February", 13, "Andrew Fuller", null);
            datasource.Rows.Add("February", 14, "Andrew Fuller", "S");
            datasource.Rows.Add("February", 15, "Andrew Fuller", null);
            datasource.Rows.Add("February", 16, "Andrew Fuller", null);
            datasource.Rows.Add("February", 17, "Andrew Fuller", null);
            datasource.Rows.Add("February", 18, "Andrew Fuller", null);
            datasource.Rows.Add("February", 19, "Andrew Fuller", null);
            datasource.Rows.Add("February", 20, "Andrew Fuller", null);
            datasource.Rows.Add("February", 21, "Andrew Fuller", null);
            datasource.Rows.Add("February", 22, "Andrew Fuller", "V");
            datasource.Rows.Add("February", 23, "Andrew Fuller", "V");
            datasource.Rows.Add("February", 24, "Andrew Fuller", "V");
            datasource.Rows.Add("February", 25, "Andrew Fuller", null);
            datasource.Rows.Add("February", 26, "Andrew Fuller", null);
            datasource.Rows.Add("February", 27, "Andrew Fuller", null);
            datasource.Rows.Add("February", 28, "Andrew Fuller", null);
            datasource.Rows.Add("February", 1, "Kiara Davidson", null);
            datasource.Rows.Add("February", 2, "Kiara Davidson", null);
            datasource.Rows.Add("February", 3, "Kiara Davidson", null);
            datasource.Rows.Add("February", 4, "Kiara Davidson", null);
            datasource.Rows.Add("February", 5, "Kiara Davidson", null);
            datasource.Rows.Add("February", 6, "Kiara Davidson", null);
            datasource.Rows.Add("February", 7, "Kiara Davidson", "P");
            datasource.Rows.Add("February", 8, "Kiara Davidson", null);
            datasource.Rows.Add("February", 9, "Kiara Davidson", null);
            datasource.Rows.Add("February", 10, "Kiara Davidson", null);
            datasource.Rows.Add("February", 11, "Kiara Davidson", "S");
            datasource.Rows.Add("February", 12, "Kiara Davidson", null);
            datasource.Rows.Add("February", 13, "Kiara Davidson", null);
            datasource.Rows.Add("February", 14, "Kiara Davidson", null);
            datasource.Rows.Add("February", 15, "Kiara Davidson", null);
            datasource.Rows.Add("February", 16, "Kiara Davidson", null);
            datasource.Rows.Add("February", 17, "Kiara Davidson", null);
            datasource.Rows.Add("February", 18, "Kiara Davidson", null);
            datasource.Rows.Add("February", 19, "Kiara Davidson", "V");
            datasource.Rows.Add("February", 20, "Kiara Davidson", "V");
            datasource.Rows.Add("February", 21, "Kiara Davidson", "V");
            datasource.Rows.Add("February", 22, "Kiara Davidson", null);
            datasource.Rows.Add("February", 23, "Kiara Davidson", null);
            datasource.Rows.Add("February", 24, "Kiara Davidson", null);
            datasource.Rows.Add("February", 25, "Kiara Davidson", null);
            datasource.Rows.Add("February", 26, "Kiara Davidson", null);
            datasource.Rows.Add("February", 27, "Kiara Davidson", null);
            datasource.Rows.Add("February", 28, "Kiara Davidson", null);
            datasource.Rows.Add("February", 1, "Edward Williams", null);
            datasource.Rows.Add("February", 2, "Edward Williams", null);
            datasource.Rows.Add("February", 3, "Edward Williams", null);
            datasource.Rows.Add("February", 4, "Edward Williams", "S");
            datasource.Rows.Add("February", 5, "Edward Williams", null);
            datasource.Rows.Add("February", 6, "Edward Williams", null);
            datasource.Rows.Add("February", 7, "Edward Williams", null);
            datasource.Rows.Add("February", 8, "Edward Williams", null);
            datasource.Rows.Add("February", 9, "Edward Williams", null);
            datasource.Rows.Add("February", 10, "Edward Williams", null);
            datasource.Rows.Add("February", 11, "Edward Williams", null);
            datasource.Rows.Add("February", 12, "Edward Williams", null);
            datasource.Rows.Add("February", 13, "Edward Williams", null);
            datasource.Rows.Add("February", 14, "Edward Williams", null);
            datasource.Rows.Add("February", 15, "Edward Williams", null);
            datasource.Rows.Add("February", 16, "Edward Williams", null);
            datasource.Rows.Add("February", 17, "Edward Williams", "S");
            datasource.Rows.Add("February", 18, "Edward Williams", null);
            datasource.Rows.Add("February", 19, "Edward Williams", null);
            datasource.Rows.Add("February", 20, "Edward Williams", null);
            datasource.Rows.Add("February", 21, "Edward Williams", null);
            datasource.Rows.Add("February", 22, "Edward Williams", null);
            datasource.Rows.Add("February", 23, "Edward Williams", null);
            datasource.Rows.Add("February", 24, "Edward Williams", "P");
            datasource.Rows.Add("February", 25, "Edward Williams", null);
            datasource.Rows.Add("February", 26, "Edward Williams", null);
            datasource.Rows.Add("February", 27, "Edward Williams", "V");
            datasource.Rows.Add("February", 28, "Edward Williams", "V");
            #endregion

            //Init template global settings
            workbook.Names.Add("TemplateOptions.KeepLineSize", "true");

            //Add data source
            workbook.AddDataSource("ds", datasource);
            //Invoke to process the template
            workbook.ProcessTemplate();
        }

        public override string TemplateName
        {
            get
            {
                return "Template_EmployeeAbsenceSchedule.xlsx";
            }
        }

        public override bool ShowTemplate
        {
            get
            {
                return true;
            }
        }

        public override bool HasTemplate
        {
            get
            {
                return true;
            }
        }

        public override bool CanDownloadZip
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
                return new string[] { "xlsx\\Template_EmployeeAbsenceSchedule.xlsx" };
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
