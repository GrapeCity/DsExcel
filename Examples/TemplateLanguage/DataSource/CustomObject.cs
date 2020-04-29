using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.DataSource
{
    public class CustomObject : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_SalesDataGroup.xlsx");
            workbook.Open(templateFile);

            #region Define custom classes
            //public class SalesData
            //{
            //    public List<SalesRecord> Sales;
            //}

            //public class SalesRecord
            //{
            //    public string Area;
            //    public string City;
            //    public string Category;
            //    public string Name;
            //    public double Revenue;
            //}
            #endregion

            var datasource = new SalesData
            {
                Sales = new List<SalesRecord>()
            };

            #region Init Data
            var record1 = new SalesRecord
            {
                Area = "North America",
                City = "Chicago",
                Category = "Consumer Electronics",
                Name = "Bose 785593-0050",
                Revenue = 92800
            };
            datasource.Sales.Add(record1);

            var record2 = new SalesRecord
            {
                Area = "North America",
                City = "New York",
                Category = "Consumer Electronics",
                Name = "Bose 785593-0050",
                Revenue = 92800
            };
            datasource.Sales.Add(record2);

            var record3 = new SalesRecord
            {
                Area = "South America",
                City = "Santiago",
                Category = "Consumer Electronics",
                Name = "Bose 785593-0050",
                Revenue = 19550
            };
            datasource.Sales.Add(record3);

            var record4 = new SalesRecord
            {
                Area = "North America",
                City = "Chicago",
                Category = "Consumer Electronics",
                Name = "Canon EOS 1500D",
                Revenue = 98650
            };
            datasource.Sales.Add(record4);

            var record5 = new SalesRecord
            {
                Area = "North America",
                City = "Minnesota",
                Category = "Consumer Electronics",
                Name = "Canon EOS 1500D",
                Revenue = 89110
            };
            datasource.Sales.Add(record5);

            var record6 = new SalesRecord
            {
                Area = "South America",
                City = "Santiago",
                Category = "Consumer Electronics",
                Name = "Canon EOS 1500D",
                Revenue = 459000
            };
            datasource.Sales.Add(record6);

            var record7 = new SalesRecord
            {
                Area = "North America",
                City = "Chicago",
                Category = "Consumer Electronics",
                Name = "Haier 394L 4Star",
                Revenue = 367050
            };
            datasource.Sales.Add(record7);

            var record8 = new SalesRecord
            {
                Area = "South America",
                City = "Quito",
                Category = "Consumer Electronics",
                Name = "Haier 394L 4Star",
                Revenue = 729100
            };
            datasource.Sales.Add(record8);

            var record9 = new SalesRecord
            {
                Area = "South America",
                City = "Santiago",
                Category = "Consumer Electronics",
                Name = "Haier 394L 4Star",
                Revenue = 578900
            };
            datasource.Sales.Add(record9);

            var record10 = new SalesRecord
            {
                Area = "North America",
                City = "Fremont",
                Category = "Consumer Electronics",
                Name = "IFB 6.5 Kg FullyAuto",
                Revenue = 904930
            };
            datasource.Sales.Add(record10);

            var record11 = new SalesRecord
            {
                Area = "South America",
                City = "Buenos Aires",
                Category = "Consumer Electronics",
                Name = "IFB 6.5 Kg FullyAuto",
                Revenue = 673800
            };
            datasource.Sales.Add(record11);

            var record12 = new SalesRecord
            {
                Area = "South America",
                City = "Medillin",
                Category = "Consumer Electronics",
                Name = "IFB 6.5 Kg FullyAuto",
                Revenue = 82910
            };
            datasource.Sales.Add(record12);

            var record13 = new SalesRecord
            {
                Area = "North America",
                City = "Chicago",
                Category = "Consumer Electronics",
                Name = "Mi LED 40inch",
                Revenue = 550010
            };
            datasource.Sales.Add(record13);

            var record14 = new SalesRecord
            {
                Area = "North America",
                City = "Minnesota",
                Category = "Consumer Electronics",
                Name = "Mi LED 40inch",
                Revenue = 1784702
            };
            datasource.Sales.Add(record14);

            var record15 = new SalesRecord
            {
                Area = "South America",
                City = "Santiago",
                Category = "Consumer Electronics",
                Name = "Mi LED 40inch",
                Revenue = 102905
            };
            datasource.Sales.Add(record15);

            var record16 = new SalesRecord
            {
                Area = "North America",
                City = "Chicago",
                Category = "Consumer Electronics",
                Name = "Sennheiser HD 4.40-BT",
                Revenue = 178100
            };
            datasource.Sales.Add(record16);

            var record17 = new SalesRecord
            {
                Area = "South America",
                City = "Quito",
                Category = "Consumer Electronics",
                Name = "Sennheiser HD 4.40-BT",
                Revenue = 234459
            };
            datasource.Sales.Add(record17);

            var record18 = new SalesRecord
            {
                Area = "North America",
                City = "Minnesota",
                Category = "Mobile",
                Name = "Iphone XR",
                Revenue = 1734621
            };
            datasource.Sales.Add(record18);

            var record19 = new SalesRecord
            {
                Area = "South America",
                City = "Santiago",
                Category = "Mobile",
                Name = "Iphone XR",
                Revenue = 109300
            };
            datasource.Sales.Add(record19);

            var record20 = new SalesRecord
            {
                Area = "North America",
                City = "Chicago",
                Category = "Mobile",
                Name = "OnePlus 7Pro",
                Revenue = 499100
            };
            datasource.Sales.Add(record20);

            var record21 = new SalesRecord
            {
                Area = "South America",
                City = "Quito",
                Category = "Mobile",
                Name = "OnePlus 7Pro",
                Revenue = 215000
            };
            datasource.Sales.Add(record21);

            var record22 = new SalesRecord
            {
                Area = "North America",
                City = "Minnesota",
                Category = "Mobile",
                Name = "Redmi 7",
                Revenue = 81650
            };
            datasource.Sales.Add(record22);

            var record23 = new SalesRecord
            {
                Area = "South America",
                City = "Quito",
                Category = "Mobile",
                Name = "Redmi 7",
                Revenue = 276390
            };
            datasource.Sales.Add(record23);

            var record24 = new SalesRecord
            {
                Area = "North America",
                City = "Minnesota",
                Category = "Mobile",
                Name = "Samsung S9",
                Revenue = 896250
            };
            datasource.Sales.Add(record24);

            var record25 = new SalesRecord
            {
                Area = "South America",
                City = "Buenos Aires",
                Category = "Mobile",
                Name = "Samsung S9",
                Revenue = 896250
            };
            datasource.Sales.Add(record25);

            var record26 = new SalesRecord
            {
                Area = "South America",
                City = "Quito",
                Category = "Mobile",
                Name = "Samsung S9",
                Revenue = 716520
            };
            datasource.Sales.Add(record26);
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

        public override bool ShowTemplate
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
        public List<SalesRecord> Sales;
    }

    public class SalesRecord
    {
        public string Area;
        public string City;
        public string Category;
        public string Name;
        public double Revenue;
    }
}
