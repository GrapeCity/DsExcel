using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.DataSource
{
    public class CustomObject : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Template_SalesDataGroup.xlsx from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_SalesDataGroup.xlsx");
            workbook.Open(templateFile);

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

            var datasource = new SalesData
            {
                Records = new List<SalesRecord>()
            };

            #region Init Data
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

            var record4 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Hellen",
                Product = "Carrots",
                ProductType = "Vegetable",
                Sales = 154
            };
            datasource.Records.Add(record4);

            var record5 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Fancy",
                Product = "Carrots",
                ProductType = "Vegetable",
                Sales = 131
            };
            datasource.Records.Add(record5);

            var record6 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Fancy",
                Product = "Cabbage",
                ProductType = "Vegetable",
                Sales = 98
            };
            datasource.Records.Add(record6);

            var record7 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Fancy",
                Product = "Potato",
                ProductType = "Vegetable",
                Sales = 212
            };
            datasource.Records.Add(record7);

            var record8 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Fancy",
                Product = "Apple",
                ProductType = "Fruit",
                Sales = 102
            };
            datasource.Records.Add(record8);

            var record9 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Ivan",
                Product = "Apple",
                ProductType = "Fruit",
                Sales = 164
            };
            datasource.Records.Add(record9);

            var record10 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Ivan",
                Product = "Kiwi",
                ProductType = "Fruit",
                Sales = 213
            };
            datasource.Records.Add(record10);

            var record11 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Ivan",
                Product = "Potato",
                ProductType = "Vegetable",
                Sales = 56
            };
            datasource.Records.Add(record11);

            var record12 = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Ivan",
                Product = "Cabbage",
                ProductType = "Vegetable",
                Sales = 265
            };
            datasource.Records.Add(record12);

            var record13 = new SalesRecord
            {
                Area = "SouthChina",
                Salesman = "Adam",
                Product = "Cabbage",
                ProductType = "Vegetable",
                Sales = 112
            };
            datasource.Records.Add(record13);

            var record14 = new SalesRecord
            {
                Area = "SouthChina",
                Salesman = "Adam",
                Product = "Carrots",
                ProductType = "Vegetable",
                Sales = 354
            };
            datasource.Records.Add(record14);

            var record15 = new SalesRecord
            {
                Area = "SouthChina",
                Salesman = "Adam",
                Product = "Banana",
                ProductType = "Fruit",
                Sales = 277
            };
            datasource.Records.Add(record15);

            var record16 = new SalesRecord
            {
                Area = "SouthChina",
                Salesman = "Adam",
                Product = "Apple",
                ProductType = "Fruit",
                Sales = 105
            };
            datasource.Records.Add(record16);

            var record17 = new SalesRecord
            {
                Area = "SouthChina",
                Salesman = "Bob",
                Product = "Banana",
                ProductType = "Fruit",
                Sales = 133
            };
            datasource.Records.Add(record17);

            var record18 = new SalesRecord
            {
                Area = "SouthChina",
                Salesman = "Bob",
                Product = "Cabbage",
                ProductType = "Vegetable",
                Sales = 252
            };
            datasource.Records.Add(record18);

            var record19 = new SalesRecord
            {
                Area = "SouthChina",
                Salesman = "Bob",
                Product = "Potato",
                ProductType = "Vegetable",
                Sales = 265
            };
            datasource.Records.Add(record19);

            var record20 = new SalesRecord
            {
                Area = "SouthChina",
                Salesman = "Bob",
                Product = "Kiwi",
                ProductType = "Fruit",
                Sales = 402
            };
            datasource.Records.Add(record20);
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
                return "Template_SalesDataGroup.xlsx";
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
                return new string[] { "xlsx\\Template_SalesDataGroup.xlsx" };
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
