using System;
using System.IO;
using System.Net;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

using GrapeCity.Documents.Excel;
using Microsoft.AspNetCore.Cors;
using Microsoft.DotNet.PlatformAbstractions;

namespace SpreadService.WebApi.Controllers
{
    [Route("api/[controller]/[action]")]
    public class SpreadServicesController : Controller
    {
        public SpreadServicesController()
        {
        }

        [HttpPost]
        public IActionResult ImportExcel()
        {
            Workbook workbook = new Workbook();
            workbook.Open(Request.Body);
            var ssjson = workbook.ToJson();
            return Ok(ssjson);
        }

        [HttpPost]
        public FileResult ExportExcel(string fileName = "")
        {
            // create a new Workbook and invoke FromJson to restore workbook from ssjson
            // the ssjson is from spread.sheets by invoking this.spread.toJSON()
            Workbook workbook = new Workbook();
            workbook.FromJson(HttpContext.Request.Body);

            MemoryStream stream = new MemoryStream();
            workbook.Save(stream);
            stream.Seek(0, SeekOrigin.Begin);

            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);

            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var donwloadFile = string.Format("attachment;filename={0}.xlsx;", string.IsNullOrEmpty(fileName) ? Guid.NewGuid().ToString() : WebUtility.UrlEncode(fileName));

            return File(bytes, contentType, donwloadFile);
        }

        [HttpGet("{templateName}")]
        public IActionResult GetSSJsonFromTemplate(string templateName)
        {
            string excel  = ApplicationEnvironment.ApplicationBasePath + Path.DirectorySeparatorChar + "ExcelTemplates" + Path.DirectorySeparatorChar + templateName;
            Workbook workbook = new Workbook();
            workbook.Open(excel);
            var ssjson = workbook.ToJson();
            return Ok(ssjson);
        }

        [HttpGet("{caseName}")]
        public IActionResult GetSSJsonFromUseCase(string caseName)
        {
            Workbook workbook = CreateWorkbookByCaseName(caseName);
            var ssjson = workbook.ToJson();
            return Ok(ssjson);
        }

        private Workbook CreateWorkbookByCaseName(string caseName)
        {
            switch (caseName)
            {
                case "BidTracker":
                    return GetBidTracker();
                case "AddressBook":
                    return GetAddressBook();
                case "ToDoList":
                    return GetToToList();
                default:
                    break;
            }

            return new Workbook();
        }

        private Workbook GetBidTracker()
        {
            Workbook workbook = new Workbook();

            IWorksheet worksheet = workbook.Worksheets[0];


            //**********************Set RowHeight & ColumnWidth**************
            worksheet.StandardHeight = 30;
            worksheet.Range["1:1"].RowHeight = 57.75;
            worksheet.Range["2:9"].RowHeight = 30;
            worksheet.Range["A:A"].ColumnWidth = 2.71;
            worksheet.Range["B:B"].ColumnWidth = 11.71;
            worksheet.Range["C:C"].ColumnWidth = 28;
            worksheet.Range["D:D"].ColumnWidth = 22.425;
            worksheet.Range["E:E"].ColumnWidth = 16.71;
            worksheet.Range["F:F"].ColumnWidth = 28;
            worksheet.Range["G:H"].ColumnWidth = 16.71;
            worksheet.Range["I:I"].ColumnWidth = 2.71;

            //*************************Set Table Value & Formulas********************
            ITable table = worksheet.Tables.Add(worksheet.Range["B2:H9"], true);
            worksheet.Range["B2:H9"].Value = new object[,]
            {
                { "BID #", "DESCRIPTION", "DATE RECEIVED", "AMOUNT", "PERCENT COMPLETE", "DEADLINE", "DAYS LEFT" },
                { 1, "Bid number 1", null, 2000, 0.5, null, null },
                { 2, "Bid number 2", null, 3500, 0.25, null, null },
                { 3, "Bid number 3", null, 5000, 0.3, null, null },
                { 4, "Bid number 4", null, 4000, 0.2, null, null },
                { 5, "Bid number 5", null, 4000, 0.75, null , null },
                { 6, "Bid number 6", null, 1500, 0.45, null , null },
                { 7, "Bid number 7", null, 5000, 0.65, null, null },
            };
            worksheet.Range["B1"].Value = "Bid Details";
            worksheet.Range["D3"].Formula = "=TODAY()-10";
            worksheet.Range["D4:D5"].Formula = "=TODAY()-20";
            worksheet.Range["D6"].Formula = "=TODAY()-10";
            worksheet.Range["D7"].Formula = "=TODAY()-28";
            worksheet.Range["D8"].Formula = "=TODAY()-17";
            worksheet.Range["D9"].Formula = "=TODAY()-15";
            worksheet.Range["G3:G9"].Formula = "=[@[DATE RECEIVED]]+30";
            worksheet.Range["H3:H9"].Formula = "=[@DEADLINE]-TODAY()";

            //***************************Set Table Style*******************************
            ITableStyle tableStyle = workbook.TableStyles.Add("Bid Tracker");
            workbook.DefaultTableStyle = "Bid Tracker";

            var wholeTableStyle = tableStyle.TableStyleElements[TableStyleElementType.WholeTable];

            //Set WholeTable element style.
            wholeTableStyle.Font.Color = Color.FromRGB(89, 89, 89);
            wholeTableStyle.Borders.Color = Color.FromRGB(89, 89, 89);
            wholeTableStyle.Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;

            var headerRowStyle = tableStyle.TableStyleElements[TableStyleElementType.HeaderRow];

            //Set HeaderRow element style.
            headerRowStyle.Borders.Color = Color.FromRGB(89, 89, 89);
            headerRowStyle.Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            headerRowStyle.Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            headerRowStyle.Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thin;
            headerRowStyle.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            headerRowStyle.Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.Thin;
            headerRowStyle.Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;
            headerRowStyle.Interior.Color = Color.FromRGB(131, 95, 1);
            headerRowStyle.Interior.PatternColor = Color.FromRGB(254, 184, 10);

            var totalRowStyle = tableStyle.TableStyleElements[TableStyleElementType.TotalRow];

            //Set TotalRow element style.
            totalRowStyle.Borders.Color = Color.White;
            totalRowStyle.Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            totalRowStyle.Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            totalRowStyle.Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thin;
            totalRowStyle.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            totalRowStyle.Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.Thin;
            totalRowStyle.Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;
            totalRowStyle.Interior.Color = Color.FromRGB(131, 95, 1);

            //**********************************Set Named Styles****************************
            IStyle titleStyle = workbook.Styles["Title"];
            titleStyle.Font.Name = "Calibri";
            titleStyle.Font.Size = 36;
            titleStyle.Font.Color = Color.FromRGB(56, 145, 167);
            titleStyle.IncludeAlignment = true;
            titleStyle.VerticalAlignment = VerticalAlignment.Center;

            IStyle heading1Style = workbook.Styles["Heading 1"];
            heading1Style.IncludeAlignment = true;
            heading1Style.HorizontalAlignment = HorizontalAlignment.Left;
            heading1Style.IndentLevel = 1;
            heading1Style.VerticalAlignment = VerticalAlignment.Bottom;
            heading1Style.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            heading1Style.Font.Size = 14;
            heading1Style.Font.Color = Color.White;
            heading1Style.Font.Bold = false;
            heading1Style.IncludePatterns = true;
            heading1Style.Interior.Color = Color.White;

            IStyle dateStyle = workbook.Styles.Add("Date");
            dateStyle.IncludeNumber = true;
            dateStyle.NumberFormat = "m/d/yyyy";
            dateStyle.IncludeAlignment = true;
            dateStyle.HorizontalAlignment = HorizontalAlignment.Left;
            dateStyle.IndentLevel = 1;
            dateStyle.VerticalAlignment = VerticalAlignment.Center;
            dateStyle.IncludeFont = false;
            dateStyle.IncludeBorder = false;
            dateStyle.IncludePatterns = false;

            IStyle commaStyle = workbook.Styles["Comma"];
            commaStyle.IncludeNumber = true;
            commaStyle.NumberFormat = "#,##0_);(#,##0)";
            commaStyle.IncludeAlignment = true;
            commaStyle.HorizontalAlignment = HorizontalAlignment.Left;
            commaStyle.IndentLevel = 1;
            commaStyle.VerticalAlignment = VerticalAlignment.Center;

            IStyle normalStyle = workbook.Styles["Normal"];
            normalStyle.HorizontalAlignment = HorizontalAlignment.Left;
            normalStyle.IndentLevel = 1;
            normalStyle.VerticalAlignment = VerticalAlignment.Center;
            normalStyle.WrapText = true;
            normalStyle.Font.Color = Color.FromRGB(89, 89, 89);

            IStyle currencyStyle = workbook.Styles["Currency"];
            currencyStyle.NumberFormat = "$#,##0.00";
            currencyStyle.IncludeAlignment = true;
            currencyStyle.HorizontalAlignment = HorizontalAlignment.Right;
            currencyStyle.IndentLevel = 1;
            currencyStyle.VerticalAlignment = VerticalAlignment.Center;

            IStyle percentStyle = workbook.Styles["Percent"];
            percentStyle.IncludeAlignment = true;
            percentStyle.HorizontalAlignment = HorizontalAlignment.Right;
            percentStyle.VerticalAlignment = VerticalAlignment.Center;
            percentStyle.IncludeFont = true;
            percentStyle.Font.Name = "Calibri";
            percentStyle.Font.Size = 20;
            percentStyle.Font.Bold = true;
            percentStyle.Font.Color = Color.FromRGB(89, 89, 89);

            IStyle comma0Style = workbook.Styles["Comma [0]"];
            comma0Style.NumberFormat = "#,##0_);(#,##0)";
            comma0Style.IncludeAlignment = true;
            comma0Style.HorizontalAlignment = HorizontalAlignment.Right;
            comma0Style.IndentLevel = 3;
            comma0Style.VerticalAlignment = VerticalAlignment.Center;

            //***********************************Add Conditional Formatting***************
            IDataBar dataBar = worksheet.Range["F3:F9"].FormatConditions.AddDatabar();
            dataBar.MinPoint.Type = ConditionValueTypes.Number;
            dataBar.MinPoint.Value = 1;
            dataBar.MaxPoint.Type = ConditionValueTypes.Number;
            dataBar.MaxPoint.Value = 0;

            dataBar.BarFillType = DataBarFillType.Gradient;
            dataBar.BarColor.Color = Color.FromRGB(126, 194, 211);
            dataBar.Direction = DataBarDirection.Context;

            dataBar.AxisColor.Color = Color.Black;
            dataBar.AxisPosition = DataBarAxisPosition.Automatic;

            dataBar.NegativeBarFormat.ColorType = DataBarNegativeColorType.Color;
            dataBar.NegativeBarFormat.Color.Color = Color.Red;
            dataBar.ShowValue = true;

            //***************************************Use NamedStyle*************************
            worksheet.SheetView.DisplayGridlines = false;
            table.TableStyle = tableStyle;
            worksheet.Range["B1"].Style = titleStyle;
            worksheet.Range["B1"].WrapText = false;
            worksheet.Range["B2:H2"].Style = heading1Style;
            worksheet.Range["B3:B9"].Style = commaStyle;
            worksheet.Range["C3:C9"].Style = normalStyle;
            worksheet.Range["D3:D9"].Style = dateStyle;
            worksheet.Range["E3:E9"].Style = currencyStyle;
            worksheet.Range["F3:F9"].Style = percentStyle;
            worksheet.Range["G3:G9"].Style = dateStyle;
            worksheet.Range["H3:H9"].Style = comma0Style;

            return workbook;
        }

        private Workbook GetToToList()
        {
            Workbook workbook = new Workbook();

            object[,] data = new object[,]{
                { "TASK", "PRIORITY", "STATUS", "START DATE", "DUE DATE","% COMPLETE", "DONE?", "NOTES" },
                { "First Thing I Need To Do", "Normal", "Not Started", null, null, 0, null, null },
                { "Other Thing I Need To Finish", "High", "In Progress", null , null, 0.5, null, null},
                { "Something Else To Get Done", "Low", "Complete",null, null, 1, null, null },
                { "More Errands And Things", "Normal", "In Progress", null, null, 0.75, null, null },
                { "So Much To Get Done This Week", "High", "In Progress", null, null, 0.25, null, null }
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "To-Do List";
            worksheet.TabColor = Color.FromRGB(148, 112, 135);
            worksheet.SheetView.DisplayGridlines = false;

            //Set Value.
            worksheet.Range["B1"].Value = "To-Do List";
            worksheet.Range["B2:I7"].Value = data;

            //Set formula.
            worksheet.Range["E3"].Formula = "=TODAY()";
            worksheet.Range["E4"].Formula = "=TODAY()-30";
            worksheet.Range["E5"].Formula = "=TODAY()-23";
            worksheet.Range["E6"].Formula = "=TODAY()-15";
            worksheet.Range["E7"].Formula = "=TODAY()-5";

            //Change the range's RowHeight and ColumnWidth.
            worksheet.StandardHeight = 30;
            worksheet.StandardWidth = 8.88671875;
            worksheet.Range["1:1"].RowHeight = 72.75;
            worksheet.Range["2:2"].RowHeight = 33;
            worksheet.Range["3:7"].RowHeight = 30;
            worksheet.Range["A:A"].ColumnWidth = 2.77734375;
            worksheet.Range["B:B"].ColumnWidth = 29.109375;
            worksheet.Range["C:G"].ColumnWidth = 16.77734375;
            worksheet.Range["H:H"].ColumnWidth = 10.77734375;
            worksheet.Range["I:I"].ColumnWidth = 29.6640625;
            worksheet.Range["J:J"].ColumnWidth = 2.77734375;

            //Modify the build in name styles.
            var nameStyle_Normal = workbook.Styles["Normal"];
            nameStyle_Normal.VerticalAlignment = VerticalAlignment.Center;
            nameStyle_Normal.WrapText = true;
            nameStyle_Normal.Font.ThemeFont = ThemeFont.Minor;
            nameStyle_Normal.Font.ThemeColor = ThemeColor.Dark1;
            nameStyle_Normal.Font.TintAndShade = 0.25;

            var nameStyle_Title = workbook.Styles["Title"];
            nameStyle_Title.HorizontalAlignment = HorizontalAlignment.General;
            nameStyle_Title.VerticalAlignment = VerticalAlignment.Bottom;
            nameStyle_Title.Font.ThemeFont = ThemeFont.Minor;
            nameStyle_Title.Font.Bold = true;
            nameStyle_Title.Font.Size = 38;
            nameStyle_Title.Font.ThemeColor = ThemeColor.Dark1;
            nameStyle_Title.Font.TintAndShade = 0.249946592608417;
            nameStyle_Title.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;
            nameStyle_Title.Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Dark1;
            nameStyle_Title.IncludeAlignment = true;
            nameStyle_Title.IncludeBorder = true;

            var nameStyle_Percent = workbook.Styles["Percent"];
            nameStyle_Percent.HorizontalAlignment = HorizontalAlignment.Right;
            nameStyle_Percent.IndentLevel = 1;
            nameStyle_Percent.VerticalAlignment = VerticalAlignment.Center;
            nameStyle_Percent.IncludeAlignment = true;

            var nameStyle_Heading_1 = workbook.Styles["Heading 1"];
            nameStyle_Heading_1.HorizontalAlignment = HorizontalAlignment.Left;
            nameStyle_Heading_1.VerticalAlignment = VerticalAlignment.Bottom;
            nameStyle_Heading_1.Font.ThemeFont = ThemeFont.Major;
            nameStyle_Heading_1.Font.Bold = false;
            nameStyle_Heading_1.Font.Size = 11;
            nameStyle_Heading_1.Font.ThemeColor = ThemeColor.Dark1;
            nameStyle_Heading_1.Font.TintAndShade = 0.249946592608417;
            nameStyle_Heading_1.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            nameStyle_Heading_1.IncludeNumber = true;
            nameStyle_Heading_1.IncludeAlignment = true;
            nameStyle_Heading_1.IncludeBorder = false;

            var nameStyle_Heading_2 = workbook.Styles["Heading 2"];
            nameStyle_Heading_2.HorizontalAlignment = HorizontalAlignment.Right;
            nameStyle_Heading_2.IndentLevel = 2;
            nameStyle_Heading_2.VerticalAlignment = VerticalAlignment.Bottom;
            nameStyle_Heading_2.Font.ThemeFont = ThemeFont.Major;
            nameStyle_Heading_2.Font.Bold = false;
            nameStyle_Heading_2.Font.Size = 11;
            nameStyle_Heading_2.Font.ThemeColor = ThemeColor.Dark2;
            nameStyle_Heading_2.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            nameStyle_Heading_2.IncludeNumber = true;
            nameStyle_Heading_2.IncludeAlignment = true;

            //Create custom name styes.
            IStyle nameStyle_Done = workbook.Styles.Add("Done");
            nameStyle_Done.NumberFormat = "\"Done\";\"\";\"\"";
            nameStyle_Done.HorizontalAlignment = HorizontalAlignment.Center;
            nameStyle_Done.VerticalAlignment = VerticalAlignment.Center;
            nameStyle_Done.Font.ThemeFont = ThemeFont.Minor;
            nameStyle_Done.Font.ThemeColor = ThemeColor.Light1;

            IStyle nameStyle_Date = workbook.Styles.Add("Date");
            nameStyle_Date.NumberFormat = "yyyy/m/d";
            nameStyle_Date.HorizontalAlignment = HorizontalAlignment.Right;
            nameStyle_Date.VerticalAlignment = VerticalAlignment.Center;
            nameStyle_Date.Font.ThemeFont = ThemeFont.Minor;
            nameStyle_Date.Font.ThemeColor = ThemeColor.Dark1;
            nameStyle_Date.Font.TintAndShade = 0.249946592608417;
            nameStyle_Date.IncludeBorder = false;
            nameStyle_Date.IncludePatterns = false;

            //Apply the above name styles on ranges.
            worksheet.Range["B1:I1"].Style = workbook.Styles["Title"];
            worksheet.Range["B2:D2"].Style = workbook.Styles["Heading 1"];
            worksheet.Range["E2:F2"].Style = workbook.Styles["Heading 2"];
            worksheet.Range["G2"].Style = workbook.Styles["Heading 1"];
            worksheet.Range["H2:H7"].Style = workbook.Styles["Done"];
            worksheet.Range["I2"].Style = workbook.Styles["Heading 1"];
            worksheet.Range["E3:F7"].Style = workbook.Styles["Date"];
            worksheet.Range["G3:G7"].Style = workbook.Styles["Percent"];

            //Add one custom table style.
            ITableStyle style = workbook.TableStyles.Add("To-do List");
            var wholeTableStyle = style.TableStyleElements[TableStyleElementType.WholeTable];
            wholeTableStyle.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Light1;
            wholeTableStyle.Borders[BordersIndex.EdgeBottom].TintAndShade = -0.14993743705557422;
            wholeTableStyle.Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.InsideHorizontal].ThemeColor = ThemeColor.Light1;
            wholeTableStyle.Borders[BordersIndex.InsideHorizontal].TintAndShade = -0.14993743705557422;

            //Create a table and apply the above table style.
            ITable table = worksheet.Tables.Add(worksheet.Range["B2:I7"], true);
            table.Name = "ToDoList";
            table.TableStyle = style;

            //Use table formula in table range.
            worksheet.Range["F3"].Formula = "=[@[START DATE]]+7";
            worksheet.Range["F4"].Formula = "=[@[START DATE]]+35";
            worksheet.Range["F5"].Formula = "=[@[START DATE]]+10";
            worksheet.Range["F6"].Formula = "=[@[START DATE]]+36";
            worksheet.Range["F7"].Formula = "=[@[START DATE]]+14";
            worksheet.Range["H3:H7"].Formula = "=--([@[% COMPLETE]]>=1)";

            //Add a expression rule.
            IFormatCondition expression = worksheet.Range["B3:I7"].FormatConditions.Add(FormatConditionType.Expression, FormatConditionOperator.Between, "=AND($G3=0,$G3<>\"\")", null) as IFormatCondition;
            expression.Interior.ThemeColor = ThemeColor.Light1;
            expression.Interior.TintAndShade = -0.0499893185216834;

            //Add a data bar rule.
            IDataBar dataBar = worksheet.Range["G3:G7"].FormatConditions.AddDatabar();
            dataBar.BarFillType = DataBarFillType.Solid;
            dataBar.BarColor.ThemeColor = ThemeColor.Accent1;
            dataBar.BarColor.TintAndShade = 0.39997558519241921;

            //Add an icon set rule.
            IIconSetCondition iconSet = worksheet.Range["H3:H7"].FormatConditions.AddIconSetCondition();
            iconSet.IconSet = workbook.IconSets[IconSetType.Icon3Symbols];
            iconSet.IconCriteria[0].Operator = FormatConditionOperator.GreaterEqual;
            iconSet.IconCriteria[0].Value = 1;
            iconSet.IconCriteria[0].Type = ConditionValueTypes.Number;

            //Create list validations.
            worksheet.Range["C3:C7"].Validation.Add(ValidationType.List, ValidationAlertStyle.Warning, ValidationOperator.Between, "Low, Normal, High", null);
            IValidation validation = worksheet.Range["C3:C7"].Validation;
            validation.ErrorMessage = "Select entry from the list. Select CANCEL, then press ALT+DOWN ARROW to navigate the list. Select ENTER to make selection";

            worksheet.Range["D3:D7"].Validation.Add(ValidationType.List, ValidationAlertStyle.Warning, ValidationOperator.Between, "Not Started,In Progress, Deferred, Complete", null);
            validation = worksheet.Range["D3:D7"].Validation;
            validation.ErrorMessage = "Select entry from the list. Select CANCEL, then press ALT+DOWN ARROW to navigate the list. Select ENTER to make selection";

            worksheet.Range["G3:G7"].Validation.Add(ValidationType.List, ValidationAlertStyle.Warning, ValidationOperator.Between, "0%,25%,50%,75%,100%", null);
            validation = worksheet.Range["G3:G7"].Validation;
            validation.ErrorMessage = "Select entry from the list. Select CANCEL, then press ALT+DOWN ARROW to navigate the list. Select ENTER to make selection";

            //Create custom validation.
            worksheet.Range["F3:F7"].Validation.Add(ValidationType.Custom, ValidationAlertStyle.Warning, ValidationOperator.Between, "=F3>=E3", null);
            validation = worksheet.Range["F3:F7"].Validation;
            validation.ErrorMessage = "The Due Date must be greater than or equal to the Start Date. Select YES to keep the value, NO to retry or CANCEL to clear the entry";

            //Create none validations, set inputmessage.
            worksheet.Range["B2"].Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
            validation = worksheet.Range["B2"].Validation;
            validation.InputMessage = "Enter Task in this column under this heading. Use heading filters to find specific entries";

            worksheet.Range["C2"].Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
            validation = worksheet.Range["C2"].Validation;
            validation.InputMessage = "Select Priority in this column under this heading. Press ALT+DOWN ARROW to open the drop-down list, then ENTER to make selection";

            worksheet.Range["D2"].Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
            validation = worksheet.Range["D2"].Validation;
            validation.InputMessage = "Select Status in this column under this heading.  Press ALT+DOWN ARROW to open the drop-down list, then ENTER to make selection";

            worksheet.Range["E2"].Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
            validation = worksheet.Range["E2"].Validation;
            validation.InputMessage = "Enter Start Date in this column under this heading";

            worksheet.Range["F2"].Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
            validation = worksheet.Range["F2"].Validation;
            validation.InputMessage = "Enter Due Date in this column under this heading";

            worksheet.Range["G2"].Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
            validation = worksheet.Range["G2"].Validation;
            validation.InputMessage = "Select % Complete in this column. Press ALT+DOWN ARROW to open the drop-down list, then ENTER to make selection. A status bar indicates progress toward completion";

            worksheet.Range["H2"].Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
            validation = worksheet.Range["H2"].Validation;
            validation.InputMessage = "Icon indicator for task completion in this column under this heading is automatically updated as tasks complete";

            worksheet.Range["I2"].Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
            validation = worksheet.Range["I2"].Validation;
            validation.InputMessage = "Enter Notes in this column under this heading";

            //Create customize theme.
            Themes themes = new Themes();
            ITheme theme = themes.Add("test");
            theme.ThemeColorScheme[ThemeColor.Dark1].RGB = Color.FromRGB(0, 0, 0);
            theme.ThemeColorScheme[ThemeColor.Light1].RGB = Color.FromRGB(255, 255, 255);
            theme.ThemeColorScheme[ThemeColor.Dark2].RGB = Color.FromRGB(37, 28, 34);
            theme.ThemeColorScheme[ThemeColor.Light2].RGB = Color.FromRGB(240, 248, 246);
            theme.ThemeColorScheme[ThemeColor.Accent1].RGB = Color.FromRGB(148, 112, 135);
            theme.ThemeColorScheme[ThemeColor.Accent2].RGB = Color.FromRGB(71, 166, 181);
            theme.ThemeColorScheme[ThemeColor.Accent3].RGB = Color.FromRGB(234, 194, 53);
            theme.ThemeColorScheme[ThemeColor.Accent4].RGB = Color.FromRGB(107, 192, 129);
            theme.ThemeColorScheme[ThemeColor.Accent5].RGB = Color.FromRGB(233, 115, 61);
            theme.ThemeColorScheme[ThemeColor.Accent6].RGB = Color.FromRGB(251, 147, 59);
            theme.ThemeColorScheme[ThemeColor.Hyperlink].RGB = Color.FromRGB(71, 166, 181);
            theme.ThemeColorScheme[ThemeColor.FollowedHyperlink].RGB = Color.FromRGB(148, 112, 135);
            theme.ThemeFontScheme.Major[FontLanguageIndex.Latin].Name = "Franklin Gothic Medium";
            theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Bookman Old Style";

            //Apply the above custom theme.
            workbook.Theme = theme;

            //Set active cell.
            worksheet.Range["G4"].Activate();

            return workbook;
        }

        private Workbook GetAddressBook()
        {
            Workbook workbook = new Workbook();

            IWorksheet worksheet = workbook.Worksheets[0];

            //***************************Set RowHeight & Width****************************
            worksheet.StandardHeight = 30;
            worksheet.Range["3:4"].RowHeight = 30;
            worksheet.Range["1:1"].RowHeight = 103.50;
            worksheet.Range["2:2"].RowHeight = 38.25;
            worksheet.Range["A:A"].ColumnWidth = 2.625;
            worksheet.Range["B:B"].ColumnWidth = 22.25;
            worksheet.Range["C:E"].ColumnWidth = 17.25;
            worksheet.Range["F:F"].ColumnWidth = 31.875;
            worksheet.Range["G:G"].ColumnWidth = 22.625;
            worksheet.Range["H:H"].ColumnWidth = 30;
            worksheet.Range["I:I"].ColumnWidth = 20.25;
            worksheet.Range["J:J"].ColumnWidth = 17.625;
            worksheet.Range["K:K"].ColumnWidth = 12.625;
            worksheet.Range["L:L"].ColumnWidth = 37.25;
            worksheet.Range["M:M"].ColumnWidth = 2.625;


            //*******************************Set Table Value & Formulas*************************************
            ITable table = worksheet.Tables.Add(worksheet.Range["B2:L4"], true);
            worksheet.Range["B2:L4"].Value = new object[,]
            {
                { "NAME", "WORK", "CELL", "HOME", "EMAIL", "BIRTHDAY", "ADDRESS", "CITY", "STATE", "ZIP", "NOTE" },
                { "Kim Abercrombie", 1235550123, 1235550123, 1235550123, "someone@example.com", null, "123 N. Maple", "Cherryville", "WA", 98031, "" },
                { "John Smith", 3215550123, "", "", "someone@example.com", null, "456 E. Aspen", "", "", "", "" },
            };
            worksheet.Range["B1"].Value = "ADDRESS BOOK";
            worksheet.Range["G3"].Formula = "=TODAY()";
            worksheet.Range["G4"].Formula = "=TODAY()+5";


            //****************************Set Table Style********************************
            ITableStyle tableStyle = workbook.TableStyles.Add("Personal Address Book");
            workbook.DefaultTableStyle = "Personal Address Book";

            //Set WholeTable element style.
            var wholeTableStyle = tableStyle.TableStyleElements[TableStyleElementType.WholeTable];
            wholeTableStyle.Borders.Color = Color.FromRGB(179, 35, 23);
            wholeTableStyle.Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.Thin;
            wholeTableStyle.Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;

            //Set FirstColumn element style.
            tableStyle.TableStyleElements[TableStyleElementType.FirstColumn].Font.Bold = true;

            //Set SecondColumns element style.
            var headerRowStyle = tableStyle.TableStyleElements[TableStyleElementType.HeaderRow];
            headerRowStyle.Borders.Color = Color.FromRGB(179, 35, 23);
            headerRowStyle.Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thick;
            headerRowStyle.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;



            //***********************************Set Named Styles*****************************
            IStyle normalStyle = workbook.Styles["Normal"];
            normalStyle.Font.Name = "Arial";
            normalStyle.Font.Color = Color.FromRGB(179, 35, 23);
            normalStyle.HorizontalAlignment = HorizontalAlignment.Left;
            normalStyle.IndentLevel = 1;
            normalStyle.VerticalAlignment = VerticalAlignment.Center;
            normalStyle.WrapText = true;

            IStyle titleStyle = workbook.Styles["Title"];
            titleStyle.IncludeAlignment = true;
            titleStyle.HorizontalAlignment = HorizontalAlignment.Left;
            titleStyle.VerticalAlignment = VerticalAlignment.Center;
            titleStyle.Font.Name = "Arial";
            titleStyle.Font.Bold = true;
            titleStyle.Font.Size = 72;
            titleStyle.Font.Color = Color.FromRGB(179, 35, 23);

            IStyle heading1Style = workbook.Styles["Heading 1"];
            heading1Style.IncludeBorder = false;
            heading1Style.Font.Name = "Arial";
            heading1Style.Font.Size = 18;
            heading1Style.Font.Color = Color.FromRGB(179, 35, 23);

            IStyle dataStyle = workbook.Styles.Add("Data");
            dataStyle.IncludeNumber = true;
            dataStyle.NumberFormat = "m/d/yyyy";

            IStyle phoneStyle = workbook.Styles.Add("Phone");
            phoneStyle.IncludeNumber = true;
            phoneStyle.NumberFormat = "[<=9999999]###-####;(###) ###-####";



            //****************************************Use NamedStyle**************************
            worksheet.SheetView.DisplayGridlines = false;
            worksheet.Range["B2:L2"].Interior.Color = Color.FromRGB(217, 217, 217);
            worksheet.Range["B3:B4"].Font.Bold = true;
            worksheet.Range["2:2"].HorizontalAlignment = HorizontalAlignment.Left;

            table.TableStyle = tableStyle;
            worksheet.Range["B1"].Style = titleStyle;
            worksheet.Range["B2:L2"].Style = heading1Style;
            worksheet.Range["C3:E4"].Style = phoneStyle;
            worksheet.Range["G3:G4"].Style = dataStyle;

            return workbook; ;
        }

    }
}
