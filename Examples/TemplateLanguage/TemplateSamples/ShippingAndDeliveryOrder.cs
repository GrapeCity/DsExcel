using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateSamples
{
    public class ShippingAndDeliveryOrder : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file Template_Score.xlsx from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_ShippingAndDeliveryOrder.xlsx");
            workbook.Open(templateFile);

            #region Define custom classes
            //class PackingList
            //{
            //    public string exporter;
            //    public string address_exporter;
            //    public string country_exporter;
            //    public string phonenumber_shipper;
            //    public string shipper;

            //    public string imports;
            //    public string address_consignee;
            //    public string country_consignee;
            //    public string phonenumber_consignee;
            //    public string consignee;

            //    public int invoice_No;
            //    public DateTime date;
            //    public int reference;

            //    public string dispatchMethod;
            //    public string shipmentType;
            //    public string VA;
            //    public string voyageNo;
            //    public string portofLanding;
            //    public DateTime departureDate;
            //    public string dischargePort;
            //    public string finalDestination;

            //    public string goodsOriginCountry;
            //    public string destinationCountry;

            //    public List<Product> item;

            //    public string issuePlace;
            //    public DateTime issueDate;
            //    public string SignatoryCompany;
            //    public string SignatoryName;


            //}

            //class Product
            //{
            //    public string productcode;
            //    public string Goods;
            //    public double quantity;
            //    public double netweight;
            //    public string kindAndPackagesCount;
            //    public double grossweight;
            //    public double measurements;
            //}
            #endregion

            #region Init Data
            var packinginfo = new PackingList
            {
                exporter = "DEL Exports",
                address_exporter = "4243 Longline Vlvd Longline, CA - 98020",
                country_exporter = "United States",
                phonenumber_shipper = "010-510-22424",
                shipper = "Diana Thompson",
                imports = "Deanna Imports",
                address_consignee = "113/23, Lombard Street Halford Townsville, Melbourne, 4323",
                country_consignee = "Australia",
                phonenumber_consignee = "010-510-33232",
                consignee = "James Williams",
                invoice_No = 1934,
                date = new DateTime(2019, 1, 30),
                reference = 1934,
                dispatchMethod = "Sea",
                shipmentType = "FCL",
                goodsOriginCountry = "United States",
                destinationCountry = "Australia",
                VA = "MAKERS DYER",
                voyageNo = "6E",
                portofLanding = "Longline - California",
                departureDate = new DateTime(2019, 2, 1),
                dischargePort = "Melbourne - Australia",
                finalDestination = "Australia",
                item = new List<Product>()
                {
                    new Product
                    {
                        productcode = "P1001",
                        Goods = "Pencils - HB",
                        quantity = 5,
                        netweight = 0.1,
                        kindAndPackagesCount = "PALLET X 1",
                        grossweight = 750,
                        measurements = 1.7
                    },
                    new Product
                    {
                        productcode = "P1002",
                        Goods = "Paper - A4",
                        quantity = 3,
                        netweight = 2,
                        kindAndPackagesCount = "PALLET X 2",
                        grossweight = 250,
                        measurements = 1.4
                    }
                },
                issuePlace = "Longline",
                issueDate = new DateTime(2019, 1, 30),
                SignatoryCompany = "DEL Exports",
                SignatoryName = "Rayna Johnson"
            };

            #endregion

            //Add data source
            workbook.AddDataSource("ds", packinginfo);
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
                return "Template_ShippingAndDeliveryOrder.xlsx";
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
                return new string[] { "xlsx\\Template_ShippingAndDeliveryOrder.xlsx" };
            }
        }

        public override string[] Refs
        {
            get
            {
                return new string[] { "PackingList", "Product" };
            }
        }
    }

    class PackingList
    {
        public string exporter;
        public string address_exporter;
        public string country_exporter;
        public string phonenumber_shipper;
        public string shipper;

        public string imports;
        public string address_consignee;
        public string country_consignee;
        public string phonenumber_consignee;
        public string consignee;

        public int invoice_No;
        public DateTime date;
        public int reference;

        public string dispatchMethod;
        public string shipmentType;
        public string VA;
        public string voyageNo;
        public string portofLanding;
        public DateTime departureDate;
        public string dischargePort;
        public string finalDestination;

        public string goodsOriginCountry;
        public string destinationCountry;

        public List<Product> item;

        public string issuePlace;
        public DateTime issueDate;
        public string SignatoryCompany;
        public string SignatoryName;


    }

    class Product
    {
        public string productcode;
        public string Goods;
        public double quantity;
        public double netweight;
        public string kindAndPackagesCount;
        public double grossweight;
        public double measurements;
    }
}
