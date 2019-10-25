using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.DataSource
{
    public class DataSet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Template_Score.xlsx from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_Scores.xlsx");
            workbook.Open(templateFile);

            #region Init Data
            var team1 = new System.Data.DataTable();
            team1.Columns.Add(new DataColumn("ID", typeof(Int32)));
            team1.Columns.Add(new DataColumn("Name", typeof(string)));
            team1.Columns.Add(new DataColumn("Score", typeof(Int32)));
            team1.Columns.Add(new DataColumn("Team", typeof(string)));

            team1.Rows.Add(10, "Bob", 12, "Xi'An");
            team1.Rows.Add(11, "Tommy", 6, "Xi'An");
            team1.Rows.Add(12, "Jaguar", 15, "Xi'An");
            team1.Rows.Add(12, "Lusia", 9, "Xi'An");

            var team2 = new System.Data.DataTable();
            team2.Columns.Add(new DataColumn("ID", typeof(Int32)));
            team2.Columns.Add(new DataColumn("Name", typeof(string)));
            team2.Columns.Add(new DataColumn("Score", typeof(Int32)));
            team2.Columns.Add(new DataColumn("Team", typeof(string)));

            team2.Rows.Add(2, "Phillip", 9, "BeiJing");
            team2.Rows.Add(3, "Hunter", 10, "BeiJing");
            team2.Rows.Add(4, "Hellen", 8, "BeiJing");
            team2.Rows.Add(5, "Jim", 9, "BeiJing");

            var datasource = new System.Data.DataSet();
            datasource.Tables.Add(team1);
            datasource.Tables.Add(team2);
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
                return "Template_Scores.xlsx";
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
                return new string[] { "xlsx\\Template_Scores.xlsx" };
            }
        }
    }
}
