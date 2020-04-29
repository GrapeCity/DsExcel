using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelReporting
{
    public class BasicSalesReport : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            #region theme
            //create a custom theme.
            Themes themes = new Themes();
            ITheme theme = themes.Add("testTheme");
            theme.ThemeColorScheme[ThemeColor.Light1].RGB = Color.FromArgb(255, 255, 255);
            theme.ThemeColorScheme[ThemeColor.Dark1].RGB = Color.FromArgb(0, 0, 0);
            theme.ThemeColorScheme[ThemeColor.Light2].RGB = Color.FromArgb(255, 255, 255);
            theme.ThemeColorScheme[ThemeColor.Dark2].RGB = Color.FromArgb(0, 0, 0);
            theme.ThemeColorScheme[ThemeColor.Accent1].RGB = Color.FromArgb(140, 198, 63);
            theme.ThemeColorScheme[ThemeColor.Accent2].RGB = Color.FromArgb(242, 116, 45);
            theme.ThemeColorScheme[ThemeColor.Accent3].RGB = Color.FromArgb(106, 159, 207);
            theme.ThemeColorScheme[ThemeColor.Accent4].RGB = Color.FromArgb(242, 192, 45);
            theme.ThemeColorScheme[ThemeColor.Accent5].RGB = Color.FromArgb(146, 98, 174);
            theme.ThemeColorScheme[ThemeColor.Accent6].RGB = Color.FromArgb(121, 198, 199);
            theme.ThemeColorScheme[ThemeColor.Hyperlink].RGB = Color.FromArgb(106, 159, 207);
            theme.ThemeColorScheme[ThemeColor.FollowedHyperlink].RGB = Color.FromArgb(146, 98, 74);
            theme.ThemeFontScheme.Major[FontLanguageIndex.Latin].Name = "Garamond";
            theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Garamond";

            //assign the custom theme for workbook.
            workbook.Theme = theme;
            #endregion

            #region NameStyle
            //Change built-in custom styles.
            var style_Title = workbook.Styles["Title"];
            style_Title.Font.ThemeFont = ThemeFont.Major;
            style_Title.Font.Size = 26;
            style_Title.Font.Bold = true;
            style_Title.Font.ThemeColor = ThemeColor.Light1;
            style_Title.IncludeAlignment = false;
            style_Title.IncludeFont = true;
            style_Title.IncludeBorder = false;
            style_Title.IncludeNumber = false;
            style_Title.IncludePatterns = false;
            style_Title.IncludeProtection = false;

            var style_Normal = workbook.Styles["Normal"];
            style_Normal.HorizontalAlignment = HorizontalAlignment.General;
            style_Normal.VerticalAlignment = VerticalAlignment.Center;
            style_Normal.Font.ThemeFont = ThemeFont.Minor;
            style_Normal.Font.Size = 9;
            style_Normal.Font.ThemeColor = ThemeColor.Dark1;
            style_Normal.IncludeAlignment = true;
            style_Normal.IncludeFont = true;
            style_Normal.IncludeBorder = true;
            style_Normal.IncludeNumber = true;
            style_Normal.IncludePatterns = true;
            style_Normal.IncludeProtection = true;
            #endregion

            #region Worksheet_DataInput
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Data Input";
            //hide worksheet gridlines.
            worksheet.SheetView.DisplayGridlines = false;

            #region RowHeightColumnWidth
            worksheet.StandardHeight = 18.75;
            worksheet.StandardWidth = 8.43;

            worksheet.Range["1:1"].RowHeight = 51.75;
            worksheet.Range["2:2"].RowHeight = 20.25;
            worksheet.Range["3:87"].RowHeight = 19;

            worksheet.Range["A:A"].ColumnWidth = 2.28515625;
            worksheet.Range["B:B"].ColumnWidth = 16.85546875;
            worksheet.Range["C:C"].ColumnWidth = 20.5703125;
            worksheet.Range["D:D"].ColumnWidth = 27.7109375;
            worksheet.Range["E:E"].ColumnWidth = 17.7109375;
            worksheet.Range["F:F"].ColumnWidth = 18.140625;
            worksheet.Range["G:G"].ColumnWidth = 2.28515625;
            #endregion

            #region Values
            worksheet.Range["B1"].Value = "DATA INPUT";
            #endregion

            #region Table
            worksheet.Range["B2:F87"].Value = new object[,] {
                    { "DATE", "PRODUCT", "CUSTOMER", "AMOUNT", "QUARTER" },
                    { new DateTime(2012, 1, 1), "Product 14", "Fabrikam, Inc.", 1886, "QUARTER 1" },
                    { new DateTime(2012, 1, 3), "Product 23", "Alpine Ski House", 4022, "QUARTER 1" },
                    { new DateTime(2012, 1, 4), "Product 18", "Coho Winery", 8144, "QUARTER 1" },
                    { new DateTime(2012, 1, 7), "Product 10", "Southridge Video", 8002, "QUARTER 1" },
                    { new DateTime(2012, 1, 11), "Product 7", "Coho Winery", 6392, "QUARTER 1" },
                    { new DateTime(2012, 1, 25), "Product 1", "Contoso, Ltd", 6444, "QUARTER 1" },
                    { new DateTime(2012, 1, 30), "Product 27", "Southridge Video", 2772, "QUARTER 1" },
                    { new DateTime(2012, 2, 4), "Product 30", "City Power & Light", 8674, "QUARTER 1" },
                    { new DateTime(2012, 2, 5), "Product 16", "A. Datum Corporation", 2332, "QUARTER 1" },
                    { new DateTime(2012, 2, 8), "Product 21", "Alpine Ski House", 5370, "QUARTER 1" },
                    { new DateTime(2012, 2, 10), "Product 6", "City Power & Light", 1768, "QUARTER 1" },
                    { new DateTime(2012, 2, 17), "Product 24", "Coho Winery", 5474, "QUARTER 1" },
                    { new DateTime(2012, 2, 22), "Product 28", "Fabrikam, Inc.", 3494, "QUARTER 1" },
                    { new DateTime(2012, 2, 24), "Product 22", "City Power & Light", 1484, "QUARTER 1" },
                    { new DateTime(2012, 2, 29), "Product 26", "Humongous Insurance", 5454, "QUARTER 1" },
                    { new DateTime(2012, 3, 1), "Product 15", "City Power & Light", 2306, "QUARTER 1" },
                    { new DateTime(2012, 3, 8), "Product 9", "A. Datum Corporation", 8652, "QUARTER 1" },
                    { new DateTime(2012, 3, 14), "Product 16", "Alpine Ski House", 3594, "QUARTER 1" },
                    { new DateTime(2012, 3, 31), "Product 28", "City Power & Light", 9130, "QUARTER 1" },
                    { new DateTime(2012, 4, 3), "Product 28", "Southridge Video", 9986, "QUARTER 2" },
                    { new DateTime(2012, 4, 9), "Product 2", "Fabrikam, Inc.", 8270, "QUARTER 2" },
                    { new DateTime(2012, 4, 10), "Product 30", "A. Datum Corporation", 5184, "QUARTER 2" },
                    { new DateTime(2012, 4, 11), "Product 25", "Contoso, Ltd", 9426, "QUARTER 2" },
                    { new DateTime(2012, 4, 11), "Product 15", "Humongous Insurance", 4012, "QUARTER 2" },
                    { new DateTime(2012, 4, 15), "Product 28", "Coho Winery", 7724, "QUARTER 2" },
                    { new DateTime(2012, 5, 3), "Product 21", "Northwind Traders", 2264, "QUARTER 2" },
                    { new DateTime(2012, 5, 4), "Product 30", "Coho Winery", 9374, "QUARTER 2" },
                    { new DateTime(2012, 5, 5), "Product 17", "Humongous Insurance", 3692, "QUARTER 2" },
                    { new DateTime(2012, 5, 5), "Product 28", "Humongous Insurance", 9124, "QUARTER 2" },
                    { new DateTime(2012, 5, 19), "Product 22", "Humongous Insurance", 4962, "QUARTER 2" },
                    { new DateTime(2012, 5, 22), "Product 8", "A. Datum Corporation", 9166, "QUARTER 2" },
                    { new DateTime(2012, 5, 31), "Product 16", "Coho Winery", 5610, "QUARTER 2" },
                    { new DateTime(2012, 6, 2), "Product 8", "City Power & Light", 3322, "QUARTER 2" },
                    { new DateTime(2012, 6, 2), "Product 3", "Humongous Insurance", 2592, "QUARTER 2" },
                    { new DateTime(2012, 6, 4), "Product 13", "Contoso, Ltd", 4444, "QUARTER 2" },
                    { new DateTime(2012, 6, 9), "Product 10", "Southridge Video", 7166, "QUARTER 2" },
                    { new DateTime(2012, 6, 12), "Product 13", "Fabrikam, Inc.", 5008, "QUARTER 3" },
                    { new DateTime(2012, 6, 26), "Product 2", "Contoso, Ltd", 3578, "QUARTER 3" },
                    { new DateTime(2012, 6, 29), "Product 22", "Southridge Video", 1144, "QUARTER 3" },
                    { new DateTime(2012, 7, 2), "Product 14", "Alpine Ski House", 3696, "QUARTER 3" },
                    { new DateTime(2012, 7, 6), "Product 14", "Coho Winery", 7084, "QUARTER 3" },
                    { new DateTime(2012, 7, 6), "Product 2", "City Power & Light", 4642, "QUARTER 3" },
                    { new DateTime(2012, 7, 8), "Product 5", "Fabrikam, Inc.", 6078, "QUARTER 3" },
                    { new DateTime(2012, 7, 10), "Product 9", "Alpine Ski House", 2394, "QUARTER 3" },
                    { new DateTime(2012, 7, 15), "Product 13", "Southridge Video", 5516, "QUARTER 3" },
                    { new DateTime(2012, 7, 16), "Product 8", "Fabrikam, Inc.", 1948, "QUARTER 3" },
                    { new DateTime(2012, 7, 19), "Product 28", "Contoso, Ltd", 7280, "QUARTER 3" },
                    { new DateTime(2012, 7, 29), "Product 26", "A. Datum Corporation", 9292, "QUARTER 3" },
                    { new DateTime(2012, 8, 19), "Product 26", "Fabrikam, Inc.", 5868, "QUARTER 3" },
                    { new DateTime(2012, 8, 19), "Product 6", "Northwind Traders", 4098, "QUARTER 3" },
                    { new DateTime(2012, 8, 20), "Product 13", "A. Datum Corporation", 1270, "QUARTER 3" },
                    { new DateTime(2012, 8, 20), "Product 23", "A. Datum Corporation", 7744, "QUARTER 3" },
                    { new DateTime(2012, 8, 24), "Product 2", "Humongous Insurance", 5488, "QUARTER 3" },
                    { new DateTime(2012, 8, 24), "Product 5", "Southridge Video", 6944, "QUARTER 3" },
                    { new DateTime(2012, 8, 25), "Product 20", "Fabrikam, Inc.", 4454, "QUARTER 3" },
                    { new DateTime(2012, 8, 27), "Product 13", "City Power & Light", 7100, "QUARTER 3" },
                    { new DateTime(2012, 8, 30), "Product 8", "Humongous Insurance", 4346, "QUARTER 3" },
                    { new DateTime(2012, 9, 1), "Product 25", "Alpine Ski House", 2032, "QUARTER 3" },
                    { new DateTime(2012, 9, 17), "Product 20", "Contoso, Ltd", 2328, "QUARTER 3" },
                    { new DateTime(2012, 9, 22), "Product 18", "Humongous Insurance", 6090, "QUARTER 3" },
                    { new DateTime(2012, 9, 25), "Product 19", "Coho Winery", 8344, "QUARTER 3" },
                    { new DateTime(2012, 9, 29), "Product 11", "Alpine Ski House", 5872, "QUARTER 3" },
                    { new DateTime(2012, 9, 30), "Product 25", "Humongous Insurance", 1578, "QUARTER 3" },
                    { new DateTime(2012, 10, 2), "Product 29", "City Power & Light", 1714, "QUARTER 4" },
                    { new DateTime(2012, 10, 5), "Product 11", "Fabrikam, Inc.", 5716, "QUARTER 4" },
                    { new DateTime(2012, 10, 6), "Product 21", "Coho Winery", 8244, "QUARTER 4" },
                    { new DateTime(2012, 10, 14), "Product 11", "A. Datum Corporation", 8888, "QUARTER 4" },
                    { new DateTime(2012, 10, 14), "Product 23", "City Power & Light", 9438, "QUARTER 4" },
                    { new DateTime(2012, 11, 9), "Product 14", "Southridge Video", 6230, "QUARTER 4" },
                    { new DateTime(2012, 11, 13), "Product 17", "A. Datum Corporation", 4278, "QUARTER 4" },
                    { new DateTime(2012, 11, 17), "Product 6", "Alpine Ski House", 5438, "QUARTER 4" },
                    { new DateTime(2012, 11, 22), "Product 1", "Coho Winery", 6728, "QUARTER 4" },
                    { new DateTime(2012, 11, 25), "Product 30", "Fabrikam, Inc.", 9992, "QUARTER 4" },
                    { new DateTime(2012, 11, 25), "Product 26", "Northwind Traders", 8462, "QUARTER 4" },
                    { new DateTime(2012, 11, 27), "Product 11", "Contoso, Ltd", 7930, "QUARTER 4" },
                    { new DateTime(2012, 11, 30), "Product 30", "Northwind Traders", 8136, "QUARTER 4" },
                    { new DateTime(2012, 12, 2), "Product 13", "Humongous Insurance", 6212, "QUARTER 4" },
                    { new DateTime(2012, 12, 5), "Product 3", "Contoso, Ltd", 4946, "QUARTER 4" },
                    { new DateTime(2012, 12, 5), "Product 3", "Southridge Video", 8554, "QUARTER 4" },
                    { new DateTime(2012, 12, 10), "Product 24", "Northwind Traders", 4508, "QUARTER 4" },
                    { new DateTime(2012, 12, 21), "Product 7", "Humongous Insurance", 7300, "QUARTER 4" },
                    { new DateTime(2012, 12, 24), "Product 17", "Fabrikam, Inc.", 8292, "QUARTER 4" },
                    { new DateTime(2012, 12, 26), "Product 26", "Alpine Ski House", 9782, "QUARTER 4" },
                    { new DateTime(2013, 1, 3), "Product 19", "Fabrikam, Inc.", 8024, "QUARTER 1" },
                    { new DateTime(2013, 1, 4), "Product 22", "A. Datum Corporation", 3758, "QUARTER 1" },
                };

            ITable table_Data = worksheet.Tables.Add(worksheet.Range["B2:F87"], true);

            //set built-in table style for table.
            table_Data.TableStyle = workbook.TableStyles["TableStyleMedium2"];

            //customize table header range's style.
            table_Data.HeaderRange.HorizontalAlignment = HorizontalAlignment.Left;
            table_Data.HeaderRange.IndentLevel = 1;
            table_Data.HeaderRange.VerticalAlignment = VerticalAlignment.Center;
            table_Data.HeaderRange.Font.Size = 11;

            //customize table each column's data body range's style.
            table_Data.Columns[0].DataBodyRange.HorizontalAlignment = HorizontalAlignment.Left;
            table_Data.Columns[0].DataBodyRange.IndentLevel = 1;
            table_Data.Columns[0].DataBodyRange.VerticalAlignment = VerticalAlignment.Center;
            table_Data.Columns[1].DataBodyRange.HorizontalAlignment = HorizontalAlignment.Left;
            table_Data.Columns[1].DataBodyRange.IndentLevel = 1;
            table_Data.Columns[1].DataBodyRange.VerticalAlignment = VerticalAlignment.Center;
            table_Data.Columns[2].DataBodyRange.HorizontalAlignment = HorizontalAlignment.Left;
            table_Data.Columns[2].DataBodyRange.IndentLevel = 1;
            table_Data.Columns[2].DataBodyRange.VerticalAlignment = VerticalAlignment.Center;
            table_Data.Columns[3].DataBodyRange.HorizontalAlignment = HorizontalAlignment.Right;
            table_Data.Columns[3].DataBodyRange.IndentLevel = 1;
            table_Data.Columns[3].DataBodyRange.VerticalAlignment = VerticalAlignment.Bottom;
            table_Data.Columns[3].DataBodyRange.NumberFormat = "$#,##0.00";
            table_Data.Columns[4].DataBodyRange.HorizontalAlignment = HorizontalAlignment.Left;
            table_Data.Columns[4].DataBodyRange.IndentLevel = 1;
            table_Data.Columns[4].DataBodyRange.VerticalAlignment = VerticalAlignment.Bottom;
            #endregion

            #region Slicers
            //create slicer caches.
            ISlicerCache cache_customer = workbook.SlicerCaches.Add(table_Data, "CUSTOMER", "CUSTOMER");
            ISlicerCache cache_product = workbook.SlicerCaches.Add(table_Data, "PRODUCT", "PRODUCT");

            //create slicers.
            ISlicer slicer_customer = cache_customer.Slicers.Add(worksheet, "CUSTOMER", "CUSTOMER", 51.914015748031495, 551, 144, 190);
            ISlicer slicer_product = cache_product.Slicers.Add(worksheet, "PRODUCT", "PRODUCT", 51.914015748031495, 691, 144, 190);

            //assign built-in slicer style for slicers.
            slicer_customer.Style = workbook.TableStyles["SlicerStyleDark1"];
            slicer_product.Style = workbook.TableStyles["SlicerStyleDark1"];
            #endregion

            #region Style
            worksheet.Range["1:1"].Style = workbook.Styles["Title"];
            worksheet.Range["1:1"].Interior.ThemeColor = ThemeColor.Accent1;
            worksheet.Range["1:1"].HorizontalAlignment = HorizontalAlignment.Left;
            worksheet.Range["1:1"].IndentLevel = 1;
            worksheet.Range["1:1"].VerticalAlignment = VerticalAlignment.Center;
            worksheet.Range["A1"].Style = workbook.Styles["Normal"];
        
            #endregion

            #endregion

            #region Worksheet_CustomizableReport
            IWorksheet worksheet2 = workbook.Worksheets.Add();
            worksheet2.Name = "Customizable Report";
            worksheet2.SheetView.DisplayGridlines = false;

            #region RowHeightColumnWidth
            worksheet2.StandardHeight = 16.5;
            worksheet2.StandardWidth = 8.43;
            worksheet2.Range["1:1"].RowHeight = 51.75;
            worksheet2.Range["2:116"].RowHeight = 14.25;
            worksheet2.Range["A:A"].ColumnWidth = 2.28515625;
            worksheet2.Range["B:B"].ColumnWidth = 23.140625;
            worksheet2.Range["C:C"].ColumnWidth = 15.5703125;
            worksheet2.Range["D:F"].ColumnWidth = 11.42578125;
            worksheet2.Range["G:H"].ColumnWidth = 13.42578125;
            #endregion

            #region Values
            worksheet2.Range["B1"].Value = "SALES REPORT";
            #endregion

            #region PivotTable
            //create pivot cache.
            var pivotcache = workbook.PivotCaches.Create(worksheet.Range["B2:F87"]);
            //create pivot table.
            var pivottable = worksheet2.PivotTables.Add(pivotcache, worksheet2.Range["B2"], "ReportPivot");
            //assign built-in table style for pivot table.
            pivottable.TableStyle = "PivotStyleMedium2";

            //config pivot table's fields.
            var field_product = pivottable.PivotFields["PRODUCT"];
            field_product.Orientation = PivotFieldOrientation.RowField;

            var field_customer = pivottable.PivotFields["CUSTOMER"];
            field_customer.Orientation = PivotFieldOrientation.RowField;

            var field_amount = pivottable.PivotFields["AMOUNT"];
            field_amount.Orientation = PivotFieldOrientation.DataField;

            var field_quarter = pivottable.PivotFields["QUARTER"];
            field_quarter.Orientation = PivotFieldOrientation.ColumnField;

            //just a workaround, we dost not provide api to set number format
            //for pivot field. so use follow way.
            pivottable.DataBodyRange.NumberFormat = "$#,##0.00_);($#,##0.00)";
            #endregion

            #region Style
            worksheet2.Range["1:1"].Style = workbook.Styles["Title"];
            worksheet2.Range["1:1"].Interior.ThemeColor = ThemeColor.Accent1;
            worksheet2.Range["1:1"].HorizontalAlignment = HorizontalAlignment.Left;
            worksheet2.Range["1:1"].IndentLevel = 1;
            worksheet2.Range["1:1"].VerticalAlignment = VerticalAlignment.Center;
            worksheet2.Range["A1"].Style = workbook.Styles["Normal"];
            #endregion

            #region Shape
            //create a shape.
            IShape shape = worksheet2.Shapes.AddShape(AutoShapeType.RectangularCallout, 472, 65, 300, 70);
            //config shape's line and fill.
            shape.Line.DashStyle = LineDashStyle.Solid;
            shape.Line.Weight = 4;
            shape.Line.Color.RGB = Color.FromArgb(89, 89, 89);
            shape.Fill.Solid();
            shape.Fill.Color.ColorType = SolidColorType.None;

            //config shape's rich text.
            ITextRange shape_p1 = shape.TextFrame.TextRange.Paragraphs[0];
            shape_p1.Text = "TIP:";
            ITextRange shape_p2 = shape.TextFrame.TextRange.Paragraphs.Add();
            shape_p2.Text = "Customize this PivotTable to fit your needs! Select a cell in the PivotTable to activate the PivotTable Field List pane. Then in the task pane, drag to add, remove, or reorder the fields." + 
                " For example, drag the Product field above the Customer field for a different view. To update PivotTable data, right-click in the PivotTable and then click Refresh.";

            //set first paragraph's font style.
            shape_p1.Font.ThemeFont = ThemeFont.Major;
            shape_p1.Font.Bold = true;
            shape_p1.Font.Size = 12;
            shape_p1.Font.Color.ObjectThemeColor = ThemeColor.Dark1;
            shape_p1.Font.Color.Brightness = 0.25;

            //set second paragraph's font style.
            shape_p2.Font.ThemeFont = ThemeFont.Minor;
            shape_p2.Font.Size = 8;
            shape_p2.Font.Color.ObjectThemeColor = ThemeColor.Dark1;
            shape_p2.Font.Color.Brightness = 0.25;
            #endregion

            #endregion

            #region Worksheet_DataLists
            IWorksheet worksheet3 = workbook.Worksheets.Add();
            worksheet3.Name = "Data Lists";
            worksheet3.SheetView.DisplayGridlines = false;

            #region RowHeightColumnWidth
            worksheet3.StandardHeight = 18.75;
            worksheet3.StandardWidth = 8.43;
            worksheet3.Range["1:1"].RowHeight = 51.75;
            worksheet3.Range["2:32"].RowHeight = 19;
            worksheet3.Range["A:A, D:D"].ColumnWidth = 2.28515625;
            worksheet3.Range["B:C"].ColumnWidth = 34.42578125;
            #endregion

            #region Values
            worksheet3.Range["B1"].Value = "DATA LISTS";
            #endregion

            #region Table
            worksheet3.Range["B2:B32"].Value = new object[,] {
                    { "PRODUCTS" },
                    { "Product 1" },
                    { "Product 2" },
                    { "Product 3" },
                    { "Product 4" },
                    { "Product 5" },
                    { "Product 6" },
                    { "Product 7" },
                    { "Product 8" },
                    { "Product 9" },
                    { "Product 10" },
                    { "Product 11" },
                    { "Product 12" },
                    { "Product 13" },
                    { "Product 14" },
                    { "Product 15" },
                    { "Product 16" },
                    { "Product 17" },
                    { "Product 18" },
                    { "Product 19" },
                    { "Product 20" },
                    { "Product 21" },
                    { "Product 22" },
                    { "Product 23" },
                    { "Product 24" },
                    { "Product 25" },
                    { "Product 26" },
                    { "Product 27" },
                    { "Product 28" },
                    { "Product 29" },
                    { "Product 30" },
                };
            ITable table_Products = worksheet3.Tables.Add(worksheet.Range["B2:B32"], true);
            worksheet3.Range["C2:C30"].Value = new object[,] {
                    { "CUSTOMERS" },
                    { "A. Datum Corporation" },
                    { "Adventure Works" },
                    { "Alpine Ski House" },
                    { "Blue Yonder Airlines" },
                    { "City Power & Light" },
                    { "Coho Vineyard" },
                    { "Coho Winery" },
                    { "Coho Vineyard & Winery" },
                    { "Contoso, Ltd" },
                    { "Contoso Pharmaceuticals" },
                    { "Consolidated Messenger" },
                    { "Fabrikam, Inc." },
                    { "Fourth Coffee" },
                    { "Graphic Design Institute" },
                    { "Humongous Insurance" },
                    { "Litware, Inc." },
                    { "Lucerne Publishing" },
                    { "Margie's Travel" },
                    { "Northwind Traders" },
                    { "Proseware, Inc." },
                    { "School of Fine Art" },
                    { "Southridge Video" },
                    { "Tailspin Toys" },
                    { "Trey Research" },
                    { "The Phone Company" },
                    { "Wide World Importers" },
                    { "Wingtip Toys" },
                    { "Woodgrove Bank" },
                };
            ITable table_Customers = worksheet3.Tables.Add(worksheet.Range["C2:C30"], true);

            table_Products.TableStyle = workbook.TableStyles["TableStyleMedium2"];
            table_Products.Columns[0].Range.HorizontalAlignment = HorizontalAlignment.Left;
            table_Products.Columns[0].Range.IndentLevel = 1;
            table_Products.Columns[0].Range.VerticalAlignment = VerticalAlignment.Center;
            table_Products.HeaderRange.Font.Size = 11;

            table_Customers.TableStyle = workbook.TableStyles["TableStyleMedium2"];
            table_Customers.Columns[0].Range.HorizontalAlignment = HorizontalAlignment.Left;
            table_Customers.Columns[0].Range.IndentLevel = 1;
            table_Customers.Columns[0].Range.VerticalAlignment = VerticalAlignment.Center;
            table_Customers.ShowTableStyleFirstColumn = true;
            table_Customers.HeaderRange.Font.Size = 11;
            #endregion

            #region Style
            worksheet3.Range["1:1"].Style = workbook.Styles["Title"];
            worksheet3.Range["1:1"].Interior.ThemeColor = ThemeColor.Accent1;
            worksheet3.Range["1:1"].HorizontalAlignment = HorizontalAlignment.Left;
            worksheet3.Range["1:1"].IndentLevel = 1;
            worksheet3.Range["1:1"].VerticalAlignment = VerticalAlignment.Center;
            worksheet3.Range["A1"].Style = workbook.Styles["Normal"];
            #endregion

            #region Shape
            IShape shape2 = worksheet3.Shapes.AddShape(AutoShapeType.RectangularCallout, 380, 65, 280, 50);
            shape2.Line.DashStyle = LineDashStyle.Solid;
            shape2.Line.Weight = 4;
            shape2.Line.Color.RGB = Color.FromArgb(89, 89, 89);
            shape2.Fill.Solid();
            shape2.Fill.Color.ColorType = SolidColorType.None;

            ITextRange shape2_p1 = shape2.TextFrame.TextRange.Paragraphs[0];
            shape2_p1.Text = "TIP:";
            ITextRange shape2_p2 = shape2.TextFrame.TextRange.Paragraphs.Add();
            shape2_p2.Text = "To add a new product or customer, start typing below the table and it will automatically expand when you press the Enter or Tab key.";
            shape2_p1.Font.ThemeFont = ThemeFont.Major;
            shape2_p1.Font.Bold = true;
            shape2_p1.Font.Size = 12;
            shape2_p1.Font.Color.ObjectThemeColor = ThemeColor.Dark1;
            shape2_p1.Font.Color.Brightness = 0.25;

            shape2_p2.Font.ThemeFont = ThemeFont.Minor;
            shape2_p2.Font.Size = 8;
            shape2_p2.Font.Color.ObjectThemeColor = ThemeColor.Dark1;
            shape2_p2.Font.Color.Brightness = 0.25;
            #endregion

            #endregion

        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
    }
}
