using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateCell
{
    public class ExpressionTemplate : ExampleBase
    {

        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Template_Score.xlsx from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_ExpressionTemplate.xlsx");
            workbook.Open(templateFile);

            #region Init Data
            var datasource = new System.Data.DataTable();
            datasource.Columns.Add(new DataColumn("ID", typeof(Int32)));
            datasource.Columns.Add(new DataColumn("Name", typeof(string)));
            datasource.Columns.Add(new DataColumn("Score", typeof(Int32)));
            datasource.Columns.Add(new DataColumn("Team", typeof(string)));

            datasource.Rows.Add(10, "Bob", 12, "Xi'An");
            datasource.Rows.Add(11, "Tommy", 6, "Xi'An");
            datasource.Rows.Add(12, "Jaguar", 15, "Xi'An");
            datasource.Rows.Add(2, "Phillip", 9, "BeiJing");
            datasource.Rows.Add(3, "Hunter", 10, "BeiJing");
            datasource.Rows.Add(4, "Hellen", 8, "BeiJing");
            datasource.Rows.Add(5, "Jim", 9, "BeiJing");
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
                return "Template_ExpressionTemplate.xlsx";
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
                return new string[] { "xlsx\\Template_ExpressionTemplate.xlsx" };
            }
        }
    }
}
