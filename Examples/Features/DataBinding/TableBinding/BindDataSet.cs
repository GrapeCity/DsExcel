using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataBinding.TableBinding
{
    public class BindDataSet : ExampleBase
    {
        public override void Execute(Excel.Workbook workbook)
        {
            #region Init data
            // DataSet
            var team1 = new System.Data.DataTable("T1");
            team1.Columns.Add(new DataColumn("ID", typeof(Int32)));
            team1.Columns.Add(new DataColumn("Name", typeof(string)));
            team1.Columns.Add(new DataColumn("Score", typeof(Int32)));
            team1.Columns.Add(new DataColumn("Team", typeof(string)));

            team1.Rows.Add(10, "Bob", 12, "Xi'An");
            team1.Rows.Add(11, "Tommy", 6, "Xi'An");
            team1.Rows.Add(12, "Jaguar", 15, "Xi'An");
            team1.Rows.Add(12, "Lusia", 9, "Xi'An");

            var team2 = new System.Data.DataTable("T2");
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

            IWorksheet worksheet = workbook.Worksheets[0];

            // Add tables
            ITable table = worksheet.Tables.Add(worksheet.Range["B2:E6"], true);
            ITable table2 = worksheet.Tables.Add(worksheet.Range["G2:J6"], true);

            // Set not to auto generate table columns
            table.AutoGenerateColumns = false;
            table2.AutoGenerateColumns = false;

            // Set table binding path
            table.BindingPath = "T1";
            table2.BindingPath = "T2";

            // Set table column data field
            table.Columns[0].DataField = "ID";
            table.Columns[1].DataField = "Name";
            table.Columns[2].DataField = "Score";
            table.Columns[3].DataField = "Team";

            table2.Columns[0].DataField = "ID";
            table2.Columns[1].DataField = "Name";
            table2.Columns[2].DataField = "Score";
            table2.Columns[3].DataField = "Team";

            // Set DataSet as datasource
            worksheet.DataSource = datasource;
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
