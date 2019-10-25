using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateSamples
{
    public class FinancialDashboard : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Template_Score.xlsx from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_FinancialDashboard.xlsx");
            workbook.Open(templateFile);

            #region Init Data
            DataTable datasource = new DataTable();
            datasource.Columns.Add(new DataColumn("season", typeof(string)));
            datasource.Columns.Add(new DataColumn("country", typeof(string)));
            datasource.Columns.Add(new DataColumn("expect", typeof(double)));
            datasource.Columns.Add(new DataColumn("actual", typeof(double)));

            datasource.Rows.Add("2016 Q1", "USA", 236047, 328554);
            datasource.Rows.Add("2016 Q2", "USA", 373060, 238136);
            datasource.Rows.Add("2016 Q3", "USA", 224132, 300822);
            datasource.Rows.Add("2016 Q4", "USA", 269305, 315337);
            datasource.Rows.Add("2017 Q1", "USA", 265397, 279008);
            datasource.Rows.Add("2017 Q2", "USA", 214079, 206019);
            datasource.Rows.Add("2017 Q3", "USA", 370191, 238294);
            datasource.Rows.Add("2017 Q4", "USA", 266843, 242323);
            datasource.Rows.Add("2016 Q1", "Japan", 350156, 370834);
            datasource.Rows.Add("2016 Q2", "Japan", 369399, 247324);
            datasource.Rows.Add("2016 Q3", "Japan", 278834, 237385);
            datasource.Rows.Add("2016 Q4", "Japan", 264277, 245048);
            datasource.Rows.Add("2017 Q1", "Japan", 203006, 295389);
            datasource.Rows.Add("2017 Q2", "Japan", 276987, 215804);
            datasource.Rows.Add("2017 Q3", "Japan", 330315, 330443);
            datasource.Rows.Add("2017 Q4", "Japan", 307477, 262512);
            datasource.Rows.Add("2016 Q1", "Korea", 229432, 330368);
            datasource.Rows.Add("2016 Q2", "Korea", 321904, 279114);
            datasource.Rows.Add("2016 Q3", "Korea", 230496, 219257);
            datasource.Rows.Add("2016 Q4", "Korea", 254328, 361880);
            datasource.Rows.Add("2017 Q1", "Korea", 272263, 355419);
            datasource.Rows.Add("2017 Q2", "Korea", 214079, 231510);
            datasource.Rows.Add("2017 Q3", "Korea", 238392, 237430);
            datasource.Rows.Add("2017 Q4", "Korea", 294097, 257680);
            datasource.Rows.Add("2016 Q1", "China", 238175, 266070);
            datasource.Rows.Add("2016 Q2", "China", 202721, 353563);
            datasource.Rows.Add("2016 Q3", "China", 253279, 312586);
            datasource.Rows.Add("2016 Q4", "China", 211847, 306970);
            datasource.Rows.Add("2017 Q1", "China", 369314, 315718);
            datasource.Rows.Add("2017 Q2", "China", 201224, 368630);
            datasource.Rows.Add("2017 Q3", "China", 239792, 255108);
            datasource.Rows.Add("2017 Q4", "China", 271096, 297354);
            datasource.Rows.Add("2016 Q1", "India", 236047, 328554);
            datasource.Rows.Add("2016 Q2", "India", 373060, 238136);
            datasource.Rows.Add("2016 Q3", "India", 224132, 300822);
            datasource.Rows.Add("2016 Q4", "India", 269305, 315337);
            datasource.Rows.Add("2017 Q1", "India", 265397, 279008);
            datasource.Rows.Add("2017 Q2", "India", 214079, 206019);
            datasource.Rows.Add("2017 Q3", "India", 370191, 238294);
            datasource.Rows.Add("2017 Q4", "India", 266843, 242323);

            #endregion

            //Add data source
            workbook.AddDataSource("ds", datasource);
            //Invoke to process the template
            workbook.ProcessTemplate();
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }

        public override string TemplateName
        {
            get
            {
                return "Template_FinancialDashboard.xlsx";
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
                return new string[] { "xlsx\\Template_FinancialDashboard.xlsx" };
            }
        }
    }
}
