using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateSamples
{
    public class SalesTracker : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_SalesTracker.xlsx");
            workbook.Open(templateFile);

            #region Init Data
            var datasource = new DataTable();
            datasource.Columns.Add(new DataColumn("ProductName", typeof(string)));
            datasource.Columns.Add(new DataColumn("CostPerItem", typeof(double)));
            datasource.Columns.Add(new DataColumn("PercentMarkup", typeof(double)));
            datasource.Columns.Add(new DataColumn("TotalSold", typeof(Int32)));
            datasource.Columns.Add(new DataColumn("ShippingCharge", typeof(Int32)));
            datasource.Columns.Add(new DataColumn("ShippingCost", typeof(double)));
            datasource.Columns.Add(new DataColumn("Returns", typeof(Int32)));

            datasource.Rows.Add("Beverages", 10, 1, 15, 10, 5.75, 2);
            datasource.Rows.Add("Condiments", 11.5, 0.75, 18, 10, 5.75, 1);
            datasource.Rows.Add("Dairy Products", 13, 0.65, 20, 10, 6.25, 0);
            datasource.Rows.Add("Confections", 5, 0.9, 50, 5, 3.5, 0);
            datasource.Rows.Add("Sea Food", 4, 0.9, 42, 5, 3.25, 3);
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
                return "Template_SalesTracker.xlsx";
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
                return new string[] { "xlsx\\Template_SalesTracker.xlsx" };
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
