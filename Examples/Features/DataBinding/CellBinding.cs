using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.DataBinding
{
    public class CellBinding : ExampleBase
    {
        public override void Execute(Excel.Workbook workbook)
        {
            #region Define custom classes
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
            var record = new SalesRecord
            {
                Area = "NorthChina",
                Salesman = "Hellen",
                Product = "Apple",
                ProductType = "Fruit",
                Sales = 120
            };
            #endregion

            IWorksheet worksheet = workbook.Worksheets[0];

            // Set binding path for cell.
            worksheet.Range["A1"].BindingPath = "Area";
            worksheet.Range["B2"].BindingPath = "Salesman";
            worksheet.Range["C2"].BindingPath = "Product";
            worksheet.Range["D3"].BindingPath = "ProductType";

            // Set data source.
            worksheet.DataSource = record;
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
                return new string[] { "SalesRecord" };
            }
        }
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
