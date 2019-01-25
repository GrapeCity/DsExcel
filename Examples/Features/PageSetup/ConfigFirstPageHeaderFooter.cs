﻿using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.PageSetup
{
    public class ConfigFirstPageHeaderFooter : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            var fileStream = this.GetResourceStream("xlsx\\PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set first page headerfooter
            worksheet.PageSetup.DifferentFirstPageHeaderFooter = true;

            worksheet.PageSetup.FirstPage.CenterHeader.Text = "&T";
            worksheet.PageSetup.FirstPage.RightFooter.Text = "&D";

            //Set first page headerfooter's graphic
            worksheet.PageSetup.FirstPage.LeftFooter.Text = "&G";
            System.IO.Stream stream = this.GetResourceStream("logo.png");
            worksheet.PageSetup.FirstPage.LeftFooter.Picture.SetGraphicStream(stream, ImageType.PNG);
            worksheet.PageSetup.FirstPage.LeftFooter.Picture.Width = 100;
            worksheet.PageSetup.FirstPage.LeftFooter.Picture.Height = 13;

        }

        public override string TemplateName
        {
            get
            {
                return "PageSetup Demo.xlsx";
            }
        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "xlsx\\PageSetup Demo.xlsx", "logo.png" };
            }
        }

    }
}
