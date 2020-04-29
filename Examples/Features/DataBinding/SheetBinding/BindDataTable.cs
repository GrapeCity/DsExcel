using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataBinding.SheetBinding
{
    public class BindDataTable : ExampleBase
    {
        public override void Execute(Excel.Workbook workbook)
        {
            #region Init data
            DataTable teamInfo = new System.Data.DataTable();
            teamInfo.Columns.Add(new DataColumn("ID", typeof(Int32)));
            teamInfo.Columns.Add(new DataColumn("Name", typeof(string)));
            teamInfo.Columns.Add(new DataColumn("Score", typeof(Int32)));
            teamInfo.Columns.Add(new DataColumn("Team", typeof(string)));

            teamInfo.Rows.Add(10, "Bob", 12, "Xi'An");
            teamInfo.Rows.Add(11, "Tommy", 6, "Xi'An");
            teamInfo.Rows.Add(12, "Jaguar", 15, "Xi'An");
            teamInfo.Rows.Add(12, "Lusia", 9, "Xi'An");
            #endregion

            IWorksheet worksheet = workbook.Worksheets[0];

            // Set AutoGenerateColumns as false
            worksheet.AutoGenerateColumns = false;

            // Bind columns manually.
            worksheet.Range["A:A"].EntireColumn.BindingPath = "ID";
            worksheet.Range["B:B"].EntireColumn.BindingPath = "Name";
            worksheet.Range["C:C"].EntireColumn.BindingPath = "Score";
            worksheet.Range["D:D"].EntireColumn.BindingPath = "Team";

            // Set data source
            worksheet.DataSource = teamInfo;
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
