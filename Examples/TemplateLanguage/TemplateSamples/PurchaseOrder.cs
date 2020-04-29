using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateSamples
{
    public class PurchaseOrder : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_PurchaseOrder.xlsx");
            workbook.Open(templateFile);

            #region Define custom classes
            //public class PurchaseOrderBasicInfo
            //{
            //    public string ID;
            //    public DateTime OrderDate;
            //    public string CreditTerms;
            //    public string PONumber;
            //    public string Ref;
            //    public string DeliverToCompany;
            //    public string DeliverToAddress;
            //    public string PostalCode;
            //    public string Country;

            //}
            #endregion

            #region Init Data
            var po = new DataTable();

            po.Columns.Add(new DataColumn("s_no", typeof(Int32)));
            po.Columns.Add(new DataColumn("itemnumber", typeof(string)));
            po.Columns.Add(new DataColumn("itemdescription", typeof(string)));
            po.Columns.Add(new DataColumn("quantity", typeof(Int32)));
            po.Columns.Add(new DataColumn("um", typeof(string)));
            po.Columns.Add(new DataColumn("price", typeof(Int32)));

            po.Rows.Add(1, "P1001", "Pencils HB", 5, "dozen", 10);
            po.Rows.Add(2, "P1003", "Pencils 2B", 4, "dozen", 10);
            po.Rows.Add(3, "P1003", "Paper A4 - Photo Copier", 10, "ream", 3);
            po.Rows.Add(4, "P1234", "Pens - Ball point", 15, "boxes", 2);
            po.Rows.Add(5, "P3221", "Highligter", 8, "sets", 10);

            PurchaseOrderBasicInfo orderbasicInfo = new PurchaseOrderBasicInfo
            {
                ID = "US120499",
                OrderDate = new DateTime(2019, 7, 7),
                CreditTerms = "30",
                PONumber = "PO1011",
                Ref = "QT1231",
                DeliverToCompany = "Sanfort Pvt. Ltd.",
                DeliverToAddress = "1322, High Street, Geln Waverlay",
                PostalCode = "Victoria 3456",
                Country = "Australia"
            };
            #endregion

            //Add data source
            workbook.AddDataSource("po", po);
            workbook.AddDataSource("tax", 5);
            workbook.AddDataSource("ds", orderbasicInfo);
            //Invoke to process the template
            workbook.ProcessTemplate();
        }

        public override string TemplateName
        {
            get
            {
                return "Template_PurchaseOrder.xlsx";
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
                return new string[] { "xlsx\\Template_PurchaseOrder.xlsx" };
            }
        }

        public override string[] Refs
        {
            get
            {
                return new string[] { "PurchaseOrderBasicInfo" };
            }
        }
    }

    public class PurchaseOrderBasicInfo
    {
        public string ID;
        public DateTime OrderDate;
        public string CreditTerms;
        public string PONumber;
        public string Ref;
        public string DeliverToCompany;
        public string DeliverToAddress;
        public string PostalCode;
        public string Country;

    }
}
