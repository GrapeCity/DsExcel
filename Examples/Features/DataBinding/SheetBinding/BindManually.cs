using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataBinding.SheetBinding
{
    public class BindManually : ExampleBase
    {
        public override void Execute(Excel.Workbook workbook)
        {
            #region Define custom classes
            //public class SalesData
            //{
            //    public List<SalesRecord> Records;
            //}

            //public class SalesRecord
            //{
            //    public string Area;
            //    public string Salesman;
            //    public string Product;
            //    public string ProductType;
            //    public int Sales;
            //}
            #endregion

            #region Init data
            var datasource = new SalesData
            {
                Records = new List<SalesRecord>()
            };

            var record1 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Hellen",
                Product = "Apple",
                ProductType = "Fruit",
                Sales = 120
            };
            datasource.Records.Add(record1);

            var record2 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Hellen",
                Product = "Banana",
                ProductType = "Fruit",
                Sales = 143
            };
            datasource.Records.Add(record2);

            var record3 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Hellen",
                Product = "Kiwi",
                ProductType = "Fruit",
                Sales = 322
            };
            datasource.Records.Add(record3);
            #endregion

            IWorksheet worksheet = workbook.Worksheets[0];

            // Set AutoGenerateColumns as false
            worksheet.AutoGenerateColumns = false;

            // Bind columns manually.
            worksheet.Range["A:A"].EntireColumn.BindingPath = "Area";
            worksheet.Range["B:B"].EntireColumn.BindingPath = "Salesman";
            worksheet.Range["C:C"].EntireColumn.BindingPath = "Product";
            worksheet.Range["D:D"].EntireColumn.BindingPath = "ProductType";

            // Set data source
            worksheet.DataSource = datasource.Records;
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }

        public override string[] Refs
        {
            get
            {
                return new string[] { "SalesData", "SalesRecord" };
            }
        }
    }
}
