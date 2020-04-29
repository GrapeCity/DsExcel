using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Templates.TemplateCell
{
    public class ImageTemplate : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Load template file from resource
            var templateFile = this.GetResourceStream("xlsx\\Template_ImageTemplate.xlsx");
            workbook.Open(templateFile);

            #region Define custom classes
            //public class BikeInfo
            //{
            //    public string BikeType;
            //    public List<BikeSeries> BikeSeries;
            //}

            //public class BikeSeries
            //{
            //    public string Name;
            //    public string Description;
            //    public byte[] BikeImage;
            //    public List<Bike> Items;
            //}

            //public class Bike
            //{
            //    public string ProductNo;
            //    public string ProductName;
            //    public string Color;
            //    public int Size;
            //    public double Weight;
            //    public double Dealer;
            //    public double ListPrice;
            //}
            #endregion

            #region Init Data
            var imageStream1 = this.GetResourceStream("image\\Mountain-100.jpg");
            byte[] image1 = new byte[imageStream1.Length];
            imageStream1.Read(image1, 0, image1.Length);
            imageStream1.Close();
            var imageStream2 = this.GetResourceStream("image\\Mountain-200.jpg");
            byte[] image2 = new byte[imageStream2.Length];
            imageStream2.Read(image2, 0, image2.Length);
            imageStream2.Close();
            var imageStream3 = this.GetResourceStream("image\\Mountain-300.jpg");
            byte[] image3 = new byte[imageStream3.Length];
            imageStream3.Read(image3, 0, image3.Length);
            imageStream3.Close();
            var imageStream4 = this.GetResourceStream("image\\Mountain-400-W.jpg");
            byte[] image4 = new byte[imageStream4.Length];
            imageStream4.Read(image4, 0, image4.Length);
            imageStream4.Close();
            var imageStream5 = this.GetResourceStream("image\\Mountain-500.jpg");
            byte[] image5 = new byte[imageStream5.Length];
            imageStream5.Read(image5, 0, image5.Length);
            imageStream5.Close();
            var imageStream6 = this.GetResourceStream("image\\Road-150.jpg");
            byte[] image6 = new byte[imageStream6.Length];
            imageStream6.Read(image6, 0, image6.Length);
            imageStream6.Close();
            var imageStream7 = this.GetResourceStream("image\\Road-350-W.jpg");
            byte[] image7 = new byte[imageStream7.Length];
            imageStream7.Read(image7, 0, image7.Length);
            imageStream7.Close();
            var imageStream8 = this.GetResourceStream("image\\Touring-1000.jpg");
            byte[] image8 = new byte[imageStream8.Length];
            imageStream8.Read(image8, 0, image8.Length);
            imageStream8.Close();
            var imageStream9 = this.GetResourceStream("image\\Touring-2000.jpg");
            byte[] image9 = new byte[imageStream9.Length];
            imageStream9.Read(image9, 0, image9.Length);
            imageStream9.Close();

            var datasource = new List<BikeInfo>();

            var bike1 = new BikeInfo();
            datasource.Add(bike1);
            bike1.BikeType = "Mountain Bikes";
            bike1.BikeSeries = new List<BikeSeries>();

            var bs1 = new BikeSeries();
            bike1.BikeSeries.Add(bs1);
            bs1.Name = "Mountain-100";
            bs1.BikeImage = image1;
            bs1.Description = "Top-of-the-line competition mountain bike. Performance-enhancing options include the innovative HL Frame, super-smooth front suspension, and traction for all terrain.";
            bs1.Items = new List<Bike>();
            var bItem1 = new Bike() { ProductNo = "BK-M82S-38", ProductName = "Mountain-100 Silver, 38", Color = "Silver", Size = 38, Weight = 20.35, Dealer = 1912.1544, ListPrice = 3399.99 };
            bs1.Items.Add(bItem1);
            var bItem2 = new Bike() { ProductNo = "BK-M82B-38", ProductName = "Mountain-100 Black, 38", Color = "Black", Size = 38, Weight = 20.35, Dealer = 1898.0944, ListPrice = 3374.99 };
            bs1.Items.Add(bItem2);

            var bs2 = new BikeSeries();
            bike1.BikeSeries.Add(bs2);
            bs2.Name = "Mountain-200";
            bs2.BikeImage = image2;
            bs2.Description = "Serious back-country riding. Perfect for all levels of competition. Uses the same HL Frame as the Mountain-100.";
            bs2.Items = new List<Bike>();
            var bItem3 = new Bike() { ProductNo = "BK-M68S-42", ProductName = "Mountain-200 Silver, 42", Color = "Silver", Size = 42, Weight = 23.77, Dealer = 1265.6195, ListPrice = 2319.99 };
            bs2.Items.Add(bItem3);
            var bItem4 = new Bike() { ProductNo = "BK-M68B-38", ProductName = "Mountain-200 Black, 38", Color = "Black", Size = 38, Weight = 23.35, Dealer = 1251.9813, ListPrice = 2294.99 };
            bs2.Items.Add(bItem4);

            var bs3 = new BikeSeries();
            bike1.BikeSeries.Add(bs3);
            bs3.Name = "Mountain-300";
            bs3.BikeImage = image3;
            bs3.Description = "For true trail addicts.  An extremely durable bike that will go anywhere and keep you in control on challenging terrain - without breaking your budget.";
            bs3.Items = new List<Bike>();
            var bItem5 = new Bike() { ProductNo = "BK-M47B-38", ProductName = "Mountain-300 Black, 38", Color = "Black", Size = 38, Weight = 25.35, Dealer = 598.4354, ListPrice = 1079.99 };
            bs3.Items.Add(bItem5);
            var bItem6 = new Bike() { ProductNo = "BK-M47B-40", ProductName = "Mountain-300 Black, 40", Color = "Black", Size = 40, Weight = 25.77, Dealer = 598.4354, ListPrice = 1079.99 };
            bs3.Items.Add(bItem6);

            var bs4 = new BikeSeries();
            bike1.BikeSeries.Add(bs4);
            bs4.Name = "Mountain-400-W";
            bs4.BikeImage = image4;
            bs4.Description = "This bike delivers a high-level of performance on a budget. It is responsive and maneuverable, and offers peace-of-mind when you decide to go off-road.";
            bs4.Items = new List<Bike>();
            var bItem7 = new Bike() { ProductNo = "BKBK-M38S-38", ProductName = "Mountain-400-W Silver, 38", Color = "Silver", Size = 38, Weight = 26.35, Dealer = 419.7784, ListPrice = 769.49 };
            bs4.Items.Add(bItem7);
            var bItem8 = new Bike() { ProductNo = "BK-M38S-40", ProductName = "Mountain-400-W Silver, 40", Color = "Silver", Size = 40, Weight = 26.77, Dealer = 419.7784, ListPrice = 769.49 };
            bs4.Items.Add(bItem8);

            var bs5 = new BikeSeries();
            bike1.BikeSeries.Add(bs5);
            bs5.Name = "Mountain-500";
            bs5.BikeImage = image5;
            bs5.Description = "Suitable for any type of riding, on or off-road. Fits any budget. Smooth-shifting with a comfortable ride.";
            bs5.Items = new List<Bike>();
            var bItem9 = new Bike() { ProductNo = "BK-M18S-40", ProductName = "Mountain-500 Silver, 40", Color = "Silver", Size = 40, Weight = 27.35, Dealer = 308.2179, ListPrice = 564.99 };
            bs5.Items.Add(bItem9);
            var bItem10 = new Bike() { ProductNo = "BK-M18B-40", ProductName = "Mountain-500 Black, 40", Color = "Black", Size = 40, Weight = 27.35, Dealer = 294.5797, ListPrice = 539.99 };
            bs5.Items.Add(bItem10);


            var bike2 = new BikeInfo();
            datasource.Add(bike2);
            bike2.BikeType = "Road Bikes";
            bike2.BikeSeries = new List<BikeSeries>();

            var bs6 = new BikeSeries();
            bike2.BikeSeries.Add(bs6);
            bs6.Name = "Road-150";
            bs6.BikeImage = image6;
            bs6.Description = "This bike is ridden by race winners. Developed with the Adventure Works Cycles professional race team, it has a extremely light heat-treated aluminum frame, and steering that allows precision control.";
            bs6.Items = new List<Bike>();
            var bItem11 = new Bike() { ProductNo = "BK-R93R-62", ProductName = "Road-150 Red, 62", Color = "Red", Size = 62, Weight = 15, Dealer = 2171.2942, ListPrice = 3578.27 };
            bs6.Items.Add(bItem11);
            var bItem12 = new Bike() { ProductNo = "BK-R93R-44", ProductName = "Road-150 Red, 44", Color = "Red", Size = 44, Weight = 13.77, Dealer = 2171.2942, ListPrice = 3578.27 };
            bs6.Items.Add(bItem12);

            var bs7 = new BikeSeries();
            bike2.BikeSeries.Add(bs7);
            bs7.Name = "Road-350-W";
            bs7.BikeImage = image7;
            bs7.Description = "Cross-train, race, or just socialize on a sleek, aerodynamic bike designed for a woman.  Advanced seat technology provides comfort all day.";
            bs7.Items = new List<Bike>();
            var bItem13 = new Bike() { ProductNo = "BK-R79Y-40", ProductName = "Road-350-W Yellow, 40", Color = "Yellow", Size = 40, Weight = 15.35, Dealer = 1082.51, ListPrice = 1700.99 };
            bs7.Items.Add(bItem13);
            var bItem14 = new Bike() { ProductNo = "BK-R79Y-42", ProductName = "Road-350-W Yellow, 42", Color = "Yellow", Size = 42, Weight = 15.77, Dealer = 1082.51, ListPrice = 1700.99 };
            bs7.Items.Add(bItem14);


            var bike3 = new BikeInfo();
            datasource.Add(bike3);
            bike3.BikeType = "Touring Bikes";
            bike3.BikeSeries = new List<BikeSeries>();

            var bs8 = new BikeSeries();
            bike3.BikeSeries.Add(bs8);
            bs8.Name = "Touring-1000";
            bs8.BikeImage = image8;
            bs8.Description = "Travel in style and comfort. Designed for maximum comfort and safety. Wide gear range takes on all hills. High-tech aluminum alloy construction provides durability without added weight.";
            bs8.Items = new List<Bike>();
            var bItem15 = new Bike() { ProductNo = "BK-T79Y-46", ProductName = "Touring-1000 Yellow, 46", Color = "Yellow", Size = 46, Weight = 25.13, Dealer = 1481.9379, ListPrice = 2384.07 };
            bs8.Items.Add(bItem15);
            var bItem16 = new Bike() { ProductNo = "BK-T79U-46", ProductName = "Touring-1000 Blue, 46", Color = "Blue", Size = 46, Weight = 25.13, Dealer = 1481.9379, ListPrice = 2384.07 };
            bs8.Items.Add(bItem16);

            var bs9 = new BikeSeries();
            bike3.BikeSeries.Add(bs9);
            bs9.Name = "Touring-2000";
            bs9.BikeImage = image9;
            bs9.Description = "The plush custom saddle keeps you riding all day,  and there's plenty of space to add panniers and bike bags to the newly-redesigned carrier.  This bike has stability when fully-loaded.";
            bs9.Items = new List<Bike>();
            var bItem17 = new Bike() { ProductNo = "BK-T44U-60", ProductName = "Touring-2000 Blue, 60", Color = "Blue", Size = 60, Weight = 27.9, Dealer = 755.1508, ListPrice = 1214.85 };
            bs9.Items.Add(bItem17);
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
                return "Template_ImageTemplate.xlsx";
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
                return new string[] { "xlsx\\Template_ImageTemplate.xlsx", "image\\Mountain-100.jpg", "image\\Mountain-200.jpg", "image\\Mountain-300.jpg", "image\\Mountain-400-W.jpg", "image\\Mountain-500.jpg", "image\\Road-150.jpg", "image\\Road-350-W.jpg", "image\\Touring-1000.jpg", "image\\Touring-2000.jpg" };
            }
        }

        public override string[] Refs
        {
            get
            {
                return new string[] { "BikeInfo", "BikeSeries", "Bike" };
            }
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }

    public class BikeInfo
    {
        public string BikeType;
        public List<BikeSeries> BikeSeries;
    }

    public class BikeSeries
    {
        public string Name;
        public string Description;
        public byte[] BikeImage;
        public List<Bike> Items;
    }

    public class Bike
    {
        public string ProductNo;
        public string ProductName;
        public string Color;
        public int Size;
        public double Weight;
        public double Dealer;
        public double ListPrice;
    }
}
