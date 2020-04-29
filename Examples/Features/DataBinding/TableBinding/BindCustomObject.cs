using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataBinding.TableBinding
{
    public class BindCustomObject : ExampleBase
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

            // Add a table
            ITable table = worksheet.Tables.Add(worksheet.Range["B2:F5"], true);

            // Set not to auto generate table columns
            table.AutoGenerateColumns = false;

            // Set table binding path
            table.BindingPath = "Records";

            // Set table column data field
            table.Columns[0].DataField = "Area";
            table.Columns[1].DataField = "Salesman";
            table.Columns[2].DataField = "Product";
            table.Columns[3].DataField = "ProductType";
            table.Columns[4].DataField = "Sales";

            //Set custom object as data source
            worksheet.DataSource = datasource;
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

    public class SalesData
    {
        public List<SalesRecord> Records;
    }

    public class SalesRecord
    {
        public string Area;
        public string Salesman;
        public string Product;
        public string ProductType;
        public int Sales;
    }
}
