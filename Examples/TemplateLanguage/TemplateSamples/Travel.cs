using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateSamples
{
    public class Travel : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_Travel.xlsx");
            workbook.Open(templateFile);

            #region Init Data
            var ds1 = new DataTable();
            ds1.Columns.Add(new DataColumn("Carrier", typeof(string)));
            ds1.Columns.Add(new DataColumn("FlightNo", typeof(int)));
            ds1.Columns.Add(new DataColumn("Date", typeof(DateTime)));
            ds1.Columns.Add(new DataColumn("From", typeof(string)));
            ds1.Columns.Add(new DataColumn("DepartureTime", typeof(TimeSpan)));
            ds1.Columns.Add(new DataColumn("To", typeof(string)));
            ds1.Columns.Add(new DataColumn("ArrivalTime", typeof(TimeSpan)));
            ds1.Columns.Add(new DataColumn("ReservationNo", typeof(string)));

            ds1.Rows.Add("Trenz Airlines", 1623, new DateTime(2018, 10, 25),
                "Lorem International", new TimeSpan(7, 56, 0), 
                "Dolor Airport", new TimeSpan(9, 15, 0), "AG4567997");

            ds1.Rows.Add("Trenz Airlines", 1323, new DateTime(2018, 10, 30),
                "Lorem International", new TimeSpan(20, 25, 0),
                "Dolor Airport", new TimeSpan(21, 45, 0), "AG4567998");

            var ds2 = new DataTable();
            ds2.Columns.Add(new DataColumn("Accommodations", typeof(string)));
            ds2.Columns.Add(new DataColumn("Date", typeof(DateTime)));
            ds2.Columns.Add(new DataColumn("Concierge", typeof(string)));
            ds2.Columns.Add(new DataColumn("Phone", typeof(string)));
            ds2.Columns.Add(new DataColumn("Email", typeof(string)));
            ds2.Columns.Add(new DataColumn("AddressPart1", typeof(string)));
            ds2.Columns.Add(new DataColumn("AddressPart2", typeof(string)));
            ds2.Columns.Add(new DataColumn("ConfirmNo", typeof(string)));
            ds2.Columns.Add(new DataColumn("Days", typeof(int)));
            ds2.Columns.Add(new DataColumn("TotalCost", typeof(double)));

            ds2.Rows.Add("Lorem Hotel", new DateTime(2018, 10, 25), 
                "Charles", "01234 567 890", "charles@lorem.com",
                "123 High Street, ", "Anytown, County, Postcode", "A4567", 2, 800);

            ds2.Rows.Add("Deloz Hotel", new DateTime(2018, 10, 27),
                "James", "01234 567 890", "no_reply@example.com",
                "202 Halford Street, ", "Anytown, County, Postcode", "A4568", 3, 900);

            var ds3 = new DataTable();
            ds3.Columns.Add(new DataColumn("Contact", typeof(string)));
            ds3.Columns.Add(new DataColumn("Phone", typeof(string)));

            ds3.Rows.Add("Airline Reservations", "01234 567 890");
            ds3.Rows.Add("Hotel Reservations", "12342322232");

            var ds4 = new DataTable();
            ds4.Columns.Add(new DataColumn("Contact", typeof(string)));
            ds4.Columns.Add(new DataColumn("Phone", typeof(string)));
            ds4.Columns.Add(new DataColumn("Notes", typeof(string)));

            ds4.Rows.Add("Tom Jenkins", "01234 567 890", "tom.jerkins@trenz.com");
            ds4.Rows.Add("Rayna James", "19234222456", "ratna.james@deloz.com");
            #endregion

            //Init template global settings
            workbook.Names.Add("TemplateOptions.KeepLineSize", "true");

            //Add data source
            workbook.AddDataSource("ds1", ds1);
            workbook.AddDataSource("ds2", ds2);
            workbook.AddDataSource("ds3", ds3);
            workbook.AddDataSource("ds4", ds4);
            //Invoke to process the template
            workbook.ProcessTemplate();
        }
        
        public override string TemplateName
        {
            get
            {
                return "Template_Travel.xlsx";
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
                return new string[] { "xlsx\\Template_Travel.xlsx" };
            }
        }
    }

    
}
