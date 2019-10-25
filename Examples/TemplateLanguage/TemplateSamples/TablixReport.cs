using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateSamples
{
    public class TablixReport : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Template_Score.xlsx from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_TablixReport.xlsx");
            workbook.Open(templateFile);

            #region Init Data
            var datasource = new DataTable();
            datasource.Columns.Add(new DataColumn("OrderID", typeof(Int32)));
            datasource.Columns.Add(new DataColumn("Product", typeof(string)));
            datasource.Columns.Add(new DataColumn("Sales", typeof(double)));
            datasource.Columns.Add(new DataColumn("ProductType", typeof(string)));
            datasource.Columns.Add(new DataColumn("Year", typeof(string)));
            datasource.Columns.Add(new DataColumn("Season", typeof(string)));

            datasource.Rows.Add(1, "Röd Kaviar", 300, "Seafood", "2017", "Q3");
            datasource.Rows.Add(2, "Spegesild", 144, "Seafood", "2017", "Q3");
            datasource.Rows.Add(3, "Carnarvon Tigers", 600, "Seafood", "2017", "Q3");
            datasource.Rows.Add(4, "Spegesild", 288, "Seafood", "2017", "Q4");
            datasource.Rows.Add(5, "Carnarvon Tigers", 4250, "Seafood", "2017", "Q4");
            datasource.Rows.Add(6, "Escargots de Bourgogne", 636, "Seafood", "2017", "Q4");
            datasource.Rows.Add(7, "Röd Kaviar", 240, "Seafood", "2018", "Q1");
            datasource.Rows.Add(8, "Carnarvon Tigers", 450, "Seafood", "2018", "Q1");
            datasource.Rows.Add(9, "Röd Kaviar", 735, "Seafood", "2018", "Q2");
            datasource.Rows.Add(10, "Røgede sild", 1377, "Seafood", "2018", "Q2");
            datasource.Rows.Add(11, "Röd Kaviar", 1020, "Seafood", "2018", "Q3");
            datasource.Rows.Add(12, "Røgede sild", 190, "Seafood", "2018", "Q3");
            datasource.Rows.Add(13, "Röd Kaviar", 1725, "Seafood", "2018", "Q4");
            datasource.Rows.Add(14, "Carnarvon Tigers", 3562, "Seafood", "2018", "Q4");
            datasource.Rows.Add(15, "Sir Rodney's Marmalade", 4276, "Confections", "2017", "Q3");
            datasource.Rows.Add(16, "Maxilaku", 880, "Confections", "2017", "Q3");
            datasource.Rows.Add(17, "Maxilaku", 1040, "Confections", "2017", "Q4");
            datasource.Rows.Add(18, "NuNuCa Nuß-Nougat-Creme", 716.8, "Confections", "2017", "Q4");
            datasource.Rows.Add(19, "Sir Rodney's Marmalade", 2592, "Confections", "2018", "Q1");
            datasource.Rows.Add(20, "Maxilaku", 1296, "Confections", "2018", "Q1");
            datasource.Rows.Add(21, "Pavlova", 1473.4, "Confections", "2018", "Q1");
            datasource.Rows.Add(22, "Sir Rodney's Marmalade", 4374, "Confections", "2018", "Q2");
            datasource.Rows.Add(23, "Maxilaku", 1004, "Confections", "2018", "Q2");
            datasource.Rows.Add(24, "Pavlova", 3075, "Confections", "2018", "Q2");
            datasource.Rows.Add(25, "Sir Rodney's Marmalade", 1071, "Confections", "2018", "Q3");
            datasource.Rows.Add(26, "Maxilaku", 860, "Confections", "2018", "Q3");
            datasource.Rows.Add(27, "Pavlova", 732, "Confections", "2018", "Q3");
            datasource.Rows.Add(28, "Sir Rodney's Marmalade", 1071, "Confections", "2018", "Q4");
            datasource.Rows.Add(29, "Pavlova", 2634, "Confections", "2018", "Q4");
            datasource.Rows.Add(30, "Sir Rodney's Scones", 1790, "Confections", "2018", "Q4");
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
                return "Template_TablixReport.xlsx";
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
                return new string[] { "xlsx\\Template_TablixReport.xlsx" };
            }
        }
    }

    
}
