using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateSamples
{
    public class Sales : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Template_Score.xlsx from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_Sales.xlsx");
            workbook.Open(templateFile);

            #region Init Data
            var datasource = new DataTable();
            datasource.Columns.Add(new DataColumn("Area", typeof(string)));
            datasource.Columns.Add(new DataColumn("Salesman", typeof(string)));
            datasource.Columns.Add(new DataColumn("Product", typeof(string)));
            datasource.Columns.Add(new DataColumn("ProductType", typeof(string)));
            datasource.Columns.Add(new DataColumn("Sales", typeof(Int32)));

            datasource.Rows.Add("NorthChina", "Hellen", "Apple", "Fruit", 120);
            datasource.Rows.Add("NorthChina", "Hellen", "Banana", "Fruit", 143);
            datasource.Rows.Add("NorthChina", "Hellen", "Kiwi", "Fruit", 322);
            datasource.Rows.Add("NorthChina", "Hellen", "Carrots", "Vegetable", 154);
            datasource.Rows.Add("NorthChina", "Fancy", "Carrots", "Vegetable", 131);
            datasource.Rows.Add("NorthChina", "Fancy", "Cabbage", "Vegetable", 98);
            datasource.Rows.Add("NorthChina", "Fancy", "Potato", "Vegetable", 212);
            datasource.Rows.Add("NorthChina", "Fancy", "Apple", "Fruit", 102);
            datasource.Rows.Add("NorthChina", "Ivan", "Apple", "Fruit", 164);
            datasource.Rows.Add("NorthChina", "Ivan", "Kiwi", "Fruit", 213);
            datasource.Rows.Add("NorthChina", "Ivan", "Potato", "Vegetable", 56);
            datasource.Rows.Add("NorthChina", "Ivan", "Cabbage", "Vegetable", 265);
            datasource.Rows.Add("SouthChina", "Adam", "Cabbage", "Vegetable", 112);
            datasource.Rows.Add("SouthChina", "Adam", "Carrots", "Vegetable", 354);
            datasource.Rows.Add("SouthChina", "Adam", "Banana", "Fruit", 277);
            datasource.Rows.Add("SouthChina", "Adam", "Apple", "Fruit", 105);
            datasource.Rows.Add("SouthChina", "Bob", "Kiwi", "Fruit", 402);
            datasource.Rows.Add("SouthChina", "Bob", "Banana", "Fruit", 133);
            datasource.Rows.Add("SouthChina", "Bob", "Cabbage", "Vegetable", 252);
            datasource.Rows.Add("SouthChina", "Bob", "Potato", "Vegetable", 265);
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
                return "Template_Sales.xlsx";
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
                return new string[] { "xlsx\\Template_Sales.xlsx" };
            }
        }
    }
}
