using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateCell
{
    public class InlineTemplate : ExampleBase
    {

        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_InlineTemplate.xlsx");
            workbook.Open(templateFile);

            #region Init Data
            var datasource = new System.Data.DataTable();
            datasource.Columns.Add(new DataColumn("Area", typeof(string)));
            datasource.Columns.Add(new DataColumn("City", typeof(string)));
            datasource.Columns.Add(new DataColumn("Category", typeof(string)));
            datasource.Columns.Add(new DataColumn("Name", typeof(string)));
            datasource.Columns.Add(new DataColumn("Revenue", typeof(double)));

            datasource.Rows.Add("North America", "Chicago", "Consumer Electronics", "Bose 785593-0050", 92800);
            datasource.Rows.Add("North America", "New York", "Consumer Electronics", "Bose 785593-0050", 92800);
            datasource.Rows.Add("South America", "Santiago", "Consumer Electronics", "Bose 785593-0050", 19550);
            datasource.Rows.Add("North America", "Chicago", "Consumer Electronics", "Canon EOS 1500D", 98650);
            datasource.Rows.Add("North America", "Minnesota", "Consumer Electronics", "Canon EOS 1500D", 89110);
            datasource.Rows.Add("South America", "Santiago", "Consumer Electronics", "Canon EOS 1500D", 459000);
            datasource.Rows.Add("North America", "Chicago", "Consumer Electronics", "Haier 394L 4Star", 367050);
            datasource.Rows.Add("South America", "Quito", "Consumer Electronics", "Haier 394L 4Star", 729100);
            datasource.Rows.Add("South America", "Santiago", "Consumer Electronics", "Haier 394L 4Star", 578900);
            datasource.Rows.Add("North America", "Fremont", "Consumer Electronics", "IFB 6.5 Kg FullyAuto", 904930);
            datasource.Rows.Add("South America", "Buenos Aires", "Consumer Electronics", "IFB 6.5 Kg FullyAuto", 673800);
            datasource.Rows.Add("South America", "Medillin", "Consumer Electronics", "IFB 6.5 Kg FullyAuto", 82910);
            datasource.Rows.Add("North America", "Chicago", "Consumer Electronics", "Mi LED 40inch", 550010);
            datasource.Rows.Add("North America", "Minnesota", "Consumer Electronics", "Mi LED 40inch", 1784702);
            datasource.Rows.Add("South America", "Santiago", "Consumer Electronics", "Mi LED 40inch", 102905);
            datasource.Rows.Add("North America", "Chicago", "Consumer Electronics", "Sennheiser HD 4.40-BT", 178100);
            datasource.Rows.Add("South America", "Quito", "Consumer Electronics", "Sennheiser HD 4.40-BT", 234459);
            datasource.Rows.Add("North America", "Minnesota", "Mobile", "Iphone XR", 1734621);
            datasource.Rows.Add("South America", "Santiago", "Mobile", "Iphone XR", 109300);
            datasource.Rows.Add("North America", "Chicago", "Mobile", "OnePlus 7Pro", 499100);
            datasource.Rows.Add("South America", "Quito", "Mobile", "OnePlus 7Pro", 215000);
            datasource.Rows.Add("North America", "Minnesota", "Mobile", "Redmi 7", 81650);
            datasource.Rows.Add("South America", "Quito", "Mobile", "Redmi 7", 276390);
            datasource.Rows.Add("North America", "Minnesota", "Mobile", "Samsung S9", 896250);
            datasource.Rows.Add("South America", "Buenos Aires", "Mobile", "Samsung S9", 896250);
            datasource.Rows.Add("South America", "Quito", "Mobile", "Samsung S9", 716520);
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
                return "Template_InlineTemplate.xlsx";
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
                return new string[] { "xlsx\\Template_InlineTemplate.xlsx" };
            }
        }
    }
}
