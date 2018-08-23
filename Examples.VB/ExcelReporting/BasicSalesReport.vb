Imports System.Drawing

Namespace ExcelReporting
    Public Class BasicSalesReport
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Workbook)

#Region "theme"
            'create a custom theme.
            Dim themes As New Themes
            Dim theme As ITheme = themes.Add("testTheme")
            theme.ThemeColorScheme(ThemeColor.Light1).RGB = Color.FromArgb(255, 255, 255)
            theme.ThemeColorScheme(ThemeColor.Dark1).RGB = Color.FromArgb(0, 0, 0)
            theme.ThemeColorScheme(ThemeColor.Light2).RGB = Color.FromArgb(255, 255, 255)
            theme.ThemeColorScheme(ThemeColor.Dark2).RGB = Color.FromArgb(0, 0, 0)
            theme.ThemeColorScheme(ThemeColor.Accent1).RGB = Color.FromArgb(140, 198, 63)
            theme.ThemeColorScheme(ThemeColor.Accent2).RGB = Color.FromArgb(242, 116, 45)
            theme.ThemeColorScheme(ThemeColor.Accent3).RGB = Color.FromArgb(106, 159, 207)
            theme.ThemeColorScheme(ThemeColor.Accent4).RGB = Color.FromArgb(242, 192, 45)
            theme.ThemeColorScheme(ThemeColor.Accent5).RGB = Color.FromArgb(146, 98, 174)
            theme.ThemeColorScheme(ThemeColor.Accent6).RGB = Color.FromArgb(121, 198, 199)
            theme.ThemeColorScheme(ThemeColor.Hyperlink).RGB = Color.FromArgb(106, 159, 207)
            theme.ThemeColorScheme(ThemeColor.FollowedHyperlink).RGB = Color.FromArgb(146, 98, 74)
            theme.ThemeFontScheme.Major(FontLanguageIndex.Latin).Name = "Garamond"
            theme.ThemeFontScheme.Minor(FontLanguageIndex.Latin).Name = "Garamond"

            'assign the custom theme for workbook.
            workbook.Theme = theme
#End Region

#Region "NameStyle"
            'Change built-in custom styles.
            Dim style_Title = workbook.Styles("Title")
            style_Title.Font.ThemeFont = ThemeFont.Major
            style_Title.Font.Size = 26
            style_Title.Font.Bold = True
            style_Title.Font.ThemeColor = ThemeColor.Light1
            style_Title.IncludeAlignment = False
            style_Title.IncludeFont = True
            style_Title.IncludeBorder = False
            style_Title.IncludeNumber = False
            style_Title.IncludePatterns = False
            style_Title.IncludeProtection = False

            Dim style_Normal = workbook.Styles("Normal")
            style_Normal.HorizontalAlignment = HorizontalAlignment.General
            style_Normal.VerticalAlignment = VerticalAlignment.Center
            style_Normal.Font.ThemeFont = ThemeFont.Minor
            style_Normal.Font.Size = 9
            style_Normal.Font.ThemeColor = ThemeColor.Dark1
            style_Normal.IncludeAlignment = True
            style_Normal.IncludeFont = True
            style_Normal.IncludeBorder = True
            style_Normal.IncludeNumber = True
            style_Normal.IncludePatterns = True
            style_Normal.IncludeProtection = True
#End Region

#Region "Worksheet_DataInput"
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Name = "Data Input"

            'hide worksheet gridlines.
            worksheet.SheetView.DisplayGridlines = False

#Region "RowHeightColumnWidth"
            worksheet.StandardHeight = 18.75
            worksheet.StandardWidth = 8.43
            worksheet.Range("1:1").RowHeight = 51.75
            worksheet.Range("2:2").RowHeight = 20.25
            worksheet.Range("3:87").RowHeight = 19
            worksheet.Range("A:A").ColumnWidth = 2.28515625
            worksheet.Range("B:B").ColumnWidth = 16.85546875
            worksheet.Range("C:C").ColumnWidth = 20.5703125
            worksheet.Range("D:D").ColumnWidth = 27.7109375
            worksheet.Range("E:E").ColumnWidth = 17.7109375
            worksheet.Range("F:F").ColumnWidth = 18.140625
            worksheet.Range("G:G").ColumnWidth = 2.28515625
#End Region

#Region "Values"
            worksheet.Range!B1.Value = "DATA INPUT"
#End Region

#Region "Table"
            worksheet.Range("B2:F87").Value = {
                {"DATE", "PRODUCT", "CUSTOMER", "AMOUNT", "QUARTER"},
                {
                    #1/1/2012#,
                    "Product 14",
                    "Fabrikam, Inc.",
                    1886,
                    "QUARTER 1"
                },
                {
                    #1/3/2012#,
                    "Product 23",
                    "Alpine Ski House",
                    4022,
                    "QUARTER 1"
                },
                {
                    #1/4/2012#,
                    "Product 18",
                    "Coho Winery",
                    8144,
                    "QUARTER 1"
                },
                {
                    #1/7/2012#,
                    "Product 10",
                    "Southridge Video",
                    8002,
                    "QUARTER 1"
                },
                {
                    #1/11/2012#,
                    "Product 7",
                    "Coho Winery",
                    6392,
                    "QUARTER 1"
                },
                {
                    #1/25/2012#,
                    "Product 1",
                    "Contoso, Ltd",
                    6444,
                    "QUARTER 1"
                },
                {
                    #1/30/2012#,
                    "Product 27",
                    "Southridge Video",
                    2772,
                    "QUARTER 1"
                },
                {
                    #2/4/2012#,
                    "Product 30",
                    "City Power & Light",
                    8674,
                    "QUARTER 1"
                },
                {
                    #2/5/2012#,
                    "Product 16",
                    "A. Datum Corporation",
                    2332,
                    "QUARTER 1"
                },
                {
                    #2/8/2012#,
                    "Product 21",
                    "Alpine Ski House",
                    5370,
                    "QUARTER 1"
                },
                {
                    #2/10/2012#,
                    "Product 6",
                    "City Power & Light",
                    1768,
                    "QUARTER 1"
                },
                {
                    #2/17/2012#,
                    "Product 24",
                    "Coho Winery",
                    5474,
                    "QUARTER 1"
                },
                {
                    #2/22/2012#,
                    "Product 28",
                    "Fabrikam, Inc.",
                    3494,
                    "QUARTER 1"
                },
                {
                    #2/24/2012#,
                    "Product 22",
                    "City Power & Light",
                    1484,
                    "QUARTER 1"
                },
                {
                    #2/29/2012#,
                    "Product 26",
                    "Humongous Insurance",
                    5454,
                    "QUARTER 1"
                },
                {
                    #3/1/2012#,
                    "Product 15",
                    "City Power & Light",
                    2306,
                    "QUARTER 1"
                },
                {
                    #3/8/2012#,
                    "Product 9",
                    "A. Datum Corporation",
                    8652,
                    "QUARTER 1"
                },
                {
                    #3/14/2012#,
                    "Product 16",
                    "Alpine Ski House",
                    3594,
                    "QUARTER 1"
                },
                {
                    #3/31/2012#,
                    "Product 28",
                    "City Power & Light",
                    9130,
                    "QUARTER 1"
                },
                {
                    #4/3/2012#,
                    "Product 28",
                    "Southridge Video",
                    9986,
                    "QUARTER 2"
                },
                {
                    #4/9/2012#,
                    "Product 2",
                    "Fabrikam, Inc.",
                    8270,
                    "QUARTER 2"
                },
                {
                    #4/10/2012#,
                    "Product 30",
                    "A. Datum Corporation",
                    5184,
                    "QUARTER 2"
                },
                {
                    #4/11/2012#,
                    "Product 25",
                    "Contoso, Ltd",
                    9426,
                    "QUARTER 2"
                },
                {
                    #4/11/2012#,
                    "Product 15",
                    "Humongous Insurance",
                    4012,
                    "QUARTER 2"
                },
                {
                    #4/15/2012#,
                    "Product 28",
                    "Coho Winery",
                    7724,
                    "QUARTER 2"
                },
                {
                    #5/3/2012#,
                    "Product 21",
                    "Northwind Traders",
                    2264,
                    "QUARTER 2"
                },
                {
                    #5/4/2012#,
                    "Product 30",
                    "Coho Winery",
                    9374,
                    "QUARTER 2"
                },
                {
                    #5/5/2012#,
                    "Product 17",
                    "Humongous Insurance",
                    3692,
                    "QUARTER 2"
                },
                {
                    #5/5/2012#,
                    "Product 28",
                    "Humongous Insurance",
                    9124,
                    "QUARTER 2"
                },
                {
                    #5/19/2012#,
                    "Product 22",
                    "Humongous Insurance",
                    4962,
                    "QUARTER 2"
                },
                {
                    #5/22/2012#,
                    "Product 8",
                    "A. Datum Corporation",
                    9166,
                    "QUARTER 2"
                },
                {
                    #5/31/2012#,
                    "Product 16",
                    "Coho Winery",
                    5610,
                    "QUARTER 2"
                },
                {
                    #6/2/2012#,
                    "Product 8",
                    "City Power & Light",
                    3322,
                    "QUARTER 2"
                },
                {
                    #6/2/2012#,
                    "Product 3",
                    "Humongous Insurance",
                    2592,
                    "QUARTER 2"
                },
                {
                    #6/4/2012#,
                    "Product 13",
                    "Contoso, Ltd",
                    4444,
                    "QUARTER 2"
                },
                {
                    #6/9/2012#,
                    "Product 10",
                    "Southridge Video",
                    7166,
                    "QUARTER 2"
                },
                {
                    #6/12/2012#,
                    "Product 13",
                    "Fabrikam, Inc.",
                    5008,
                    "QUARTER 3"
                },
                {
                    #6/26/2012#,
                    "Product 2",
                    "Contoso, Ltd",
                    3578,
                    "QUARTER 3"
                },
                {
                    #6/29/2012#,
                    "Product 22",
                    "Southridge Video",
                    1144,
                    "QUARTER 3"
                },
                {
                    #7/2/2012#,
                    "Product 14",
                    "Alpine Ski House",
                    3696,
                    "QUARTER 3"
                },
                {
                    #7/6/2012#,
                    "Product 14",
                    "Coho Winery",
                    7084,
                    "QUARTER 3"
                },
                {
                    #7/6/2012#,
                    "Product 2",
                    "City Power & Light",
                    4642,
                    "QUARTER 3"
                },
                {
                    #7/8/2012#,
                    "Product 5",
                    "Fabrikam, Inc.",
                    6078,
                    "QUARTER 3"
                },
                {
                    #7/10/2012#,
                    "Product 9",
                    "Alpine Ski House",
                    2394,
                    "QUARTER 3"
                },
                {
                    #7/15/2012#,
                    "Product 13",
                    "Southridge Video",
                    5516,
                    "QUARTER 3"
                },
                {
                    #7/16/2012#,
                    "Product 8",
                    "Fabrikam, Inc.",
                    1948,
                    "QUARTER 3"
                },
                {
                    #7/19/2012#,
                    "Product 28",
                    "Contoso, Ltd",
                    7280,
                    "QUARTER 3"
                },
                {
                    #7/29/2012#,
                    "Product 26",
                    "A. Datum Corporation",
                    9292,
                    "QUARTER 3"
                },
                {
                    #8/19/2012#,
                    "Product 26",
                    "Fabrikam, Inc.",
                    5868,
                    "QUARTER 3"
                },
                {
                    #8/19/2012#,
                    "Product 6",
                    "Northwind Traders",
                    4098,
                    "QUARTER 3"
                },
                {
                    #8/20/2012#,
                    "Product 13",
                    "A. Datum Corporation",
                    1270,
                    "QUARTER 3"
                },
                {
                    #8/20/2012#,
                    "Product 23",
                    "A. Datum Corporation",
                    7744,
                    "QUARTER 3"
                },
                {
                    #8/24/2012#,
                    "Product 2",
                    "Humongous Insurance",
                    5488,
                    "QUARTER 3"
                },
                {
                    #8/24/2012#,
                    "Product 5",
                    "Southridge Video",
                    6944,
                    "QUARTER 3"
                },
                {
                    #8/25/2012#,
                    "Product 20",
                    "Fabrikam, Inc.",
                    4454,
                    "QUARTER 3"
                },
                {
                    #8/27/2012#,
                    "Product 13",
                    "City Power & Light",
                    7100,
                    "QUARTER 3"
                },
                {
                    #8/30/2012#,
                    "Product 8",
                    "Humongous Insurance",
                    4346,
                    "QUARTER 3"
                },
                {
                    #9/1/2012#,
                    "Product 25",
                    "Alpine Ski House",
                    2032,
                    "QUARTER 3"
                },
                {
                    #9/17/2012#,
                    "Product 20",
                    "Contoso, Ltd",
                    2328,
                    "QUARTER 3"
                },
                {
                    #9/22/2012#,
                    "Product 18",
                    "Humongous Insurance",
                    6090,
                    "QUARTER 3"
                },
                {
                    #9/25/2012#,
                    "Product 19",
                    "Coho Winery",
                    8344,
                    "QUARTER 3"
                },
                {
                    #9/29/2012#,
                    "Product 11",
                    "Alpine Ski House",
                    5872,
                    "QUARTER 3"
                },
                {
                    #9/30/2012#,
                    "Product 25",
                    "Humongous Insurance",
                    1578,
                    "QUARTER 3"
                },
                {
                    #10/2/2012#,
                    "Product 29",
                    "City Power & Light",
                    1714,
                    "QUARTER 4"
                },
                {
                    #10/5/2012#,
                    "Product 11",
                    "Fabrikam, Inc.",
                    5716,
                    "QUARTER 4"
                },
                {
                    #10/6/2012#,
                    "Product 21",
                    "Coho Winery",
                    8244,
                    "QUARTER 4"
                },
                {
                    #10/14/2012#,
                    "Product 11",
                    "A. Datum Corporation",
                    8888,
                    "QUARTER 4"
                },
                {
                    #10/14/2012#,
                    "Product 23",
                    "City Power & Light",
                    9438,
                    "QUARTER 4"
                },
                {
                    #11/9/2012#,
                    "Product 14",
                    "Southridge Video",
                    6230,
                    "QUARTER 4"
                },
                {
                    #11/13/2012#,
                    "Product 17",
                    "A. Datum Corporation",
                    4278,
                    "QUARTER 4"
                },
                {
                    #11/17/2012#,
                    "Product 6",
                    "Alpine Ski House",
                    5438,
                    "QUARTER 4"
                },
                {
                    #11/22/2012#,
                    "Product 1",
                    "Coho Winery",
                    6728,
                    "QUARTER 4"
                },
                {
                    #11/25/2012#,
                    "Product 30",
                    "Fabrikam, Inc.",
                    9992,
                    "QUARTER 4"
                },
                {
                    #11/25/2012#,
                    "Product 26",
                    "Northwind Traders",
                    8462,
                    "QUARTER 4"
                },
                {
                    #11/27/2012#,
                    "Product 11",
                    "Contoso, Ltd",
                    7930,
                    "QUARTER 4"
                },
                {
                    #11/30/2012#,
                    "Product 30",
                    "Northwind Traders",
                    8136,
                    "QUARTER 4"
                },
                {
                    #12/2/2012#,
                    "Product 13",
                    "Humongous Insurance",
                    6212,
                    "QUARTER 4"
                },
                {
                    #12/5/2012#,
                    "Product 3",
                    "Contoso, Ltd",
                    4946,
                    "QUARTER 4"
                },
                {
                    #12/5/2012#,
                    "Product 3",
                    "Southridge Video",
                    8554,
                    "QUARTER 4"
                },
                {
                    #12/10/2012#,
                    "Product 24",
                    "Northwind Traders",
                    4508,
                    "QUARTER 4"
                },
                {
                    #12/21/2012#,
                    "Product 7",
                    "Humongous Insurance",
                    7300,
                    "QUARTER 4"
                },
                {
                    #12/24/2012#,
                    "Product 17",
                    "Fabrikam, Inc.",
                    8292,
                    "QUARTER 4"
                },
                {
                    #12/26/2012#,
                    "Product 26",
                    "Alpine Ski House",
                    9782,
                    "QUARTER 4"
                },
                {
                    #1/3/2013#,
                    "Product 19",
                    "Fabrikam, Inc.",
                    8024,
                    "QUARTER 1"
                },
                {
                    #1/4/2013#,
                    "Product 22",
                    "A. Datum Corporation",
                    3758,
                    "QUARTER 1"
                }
            }
            Dim table_Data As ITable = worksheet.Tables.Add(worksheet.Range("B2:F87"), True)

            'set built-in table style for table.
            table_Data.TableStyle = workbook.TableStyles("TableStyleMedium2")

            'customize table header range's style.
            table_Data.HeaderRange.HorizontalAlignment = HorizontalAlignment.Left
            table_Data.HeaderRange.IndentLevel = 1
            table_Data.HeaderRange.VerticalAlignment = VerticalAlignment.Center
            table_Data.HeaderRange.Font.Size = 11

            'customize table each column's data body range's style.
            table_Data.Columns(0).DataBodyRange.HorizontalAlignment = HorizontalAlignment.Left
            table_Data.Columns(0).DataBodyRange.IndentLevel = 1
            table_Data.Columns(0).DataBodyRange.VerticalAlignment = VerticalAlignment.Center
            table_Data.Columns(1).DataBodyRange.HorizontalAlignment = HorizontalAlignment.Left
            table_Data.Columns(1).DataBodyRange.IndentLevel = 1
            table_Data.Columns(1).DataBodyRange.VerticalAlignment = VerticalAlignment.Center
            table_Data.Columns(2).DataBodyRange.HorizontalAlignment = HorizontalAlignment.Left
            table_Data.Columns(2).DataBodyRange.IndentLevel = 1
            table_Data.Columns(2).DataBodyRange.VerticalAlignment = VerticalAlignment.Center
            table_Data.Columns(3).DataBodyRange.HorizontalAlignment = HorizontalAlignment.Right
            table_Data.Columns(3).DataBodyRange.IndentLevel = 1
            table_Data.Columns(3).DataBodyRange.VerticalAlignment = VerticalAlignment.Bottom
            table_Data.Columns(3).DataBodyRange.NumberFormat = "$#,##0.00"
            table_Data.Columns(4).DataBodyRange.HorizontalAlignment = HorizontalAlignment.Left
            table_Data.Columns(4).DataBodyRange.IndentLevel = 1
            table_Data.Columns(4).DataBodyRange.VerticalAlignment = VerticalAlignment.Bottom
#End Region

#Region "Slicers"
            'create slicer caches.
            Dim cache_customer As ISlicerCache = workbook.SlicerCaches.Add(table_Data, "CUSTOMER", "CUSTOMER")
            Dim cache_product As ISlicerCache = workbook.SlicerCaches.Add(table_Data, "PRODUCT", "PRODUCT")

            'create slicers.
            Dim slicer_customer As ISlicer = cache_customer.Slicers.Add(worksheet, "CUSTOMER", "CUSTOMER", 51.914015748031495, 551, 144, 190)
            Dim slicer_product As ISlicer = cache_product.Slicers.Add(worksheet, "PRODUCT", "PRODUCT", 51.914015748031495, 691, 144, 190)

            'assign built-in slicer style for slicers.
            slicer_customer.Style = workbook.TableStyles("SlicerStyleDark1")
            slicer_product.Style = workbook.TableStyles("SlicerStyleDark1")
#End Region

#Region "Style"
            worksheet.Range("1:1").Style = workbook.Styles("Title")
            worksheet.Range("1:1").Interior.ThemeColor = ThemeColor.Accent1
            worksheet.Range("1:1").HorizontalAlignment = HorizontalAlignment.Left
            worksheet.Range("1:1").IndentLevel = 1
            worksheet.Range("1:1").VerticalAlignment = VerticalAlignment.Center
            worksheet.Range!A1.Style = workbook.Styles("Normal")
#End Region

#End Region

#Region "Worksheet_CustomizableReport"
            Dim worksheet2 As IWorksheet = workbook.Worksheets.Add()
            worksheet2.Name = "Customizable Report"
            worksheet2.SheetView.DisplayGridlines = False

#Region "RowHeightColumnWidth"
            worksheet2.StandardHeight = 16.5
            worksheet2.StandardWidth = 8.43
            worksheet2.Range("1:1").RowHeight = 51.75
            worksheet2.Range("2:116").RowHeight = 14.25
            worksheet2.Range("A:A").ColumnWidth = 2.28515625
            worksheet2.Range("B:B").ColumnWidth = 23.140625
            worksheet2.Range("C:C").ColumnWidth = 15.5703125
            worksheet2.Range("D:F").ColumnWidth = 11.42578125
            worksheet2.Range("G:H").ColumnWidth = 13.42578125
#End Region

#Region "Values"
            worksheet2.Range!B1.Value = "SALES REPORT"
#End Region

#Region "PivotTable"
            'create pivot cache.
            Dim pivotcache = workbook.PivotCaches.Create(worksheet.Range("B2:F87"))

            'create pivot table.
            Dim pivottable = worksheet2.PivotTables.Add(pivotcache, worksheet2.Range!B2, "ReportPivot")

            'assign built-in table style for pivot table.
            pivottable.TableStyle = "PivotStyleMedium2"

            'config pivot table's fields.
            Dim field_product = pivottable.PivotFields("PRODUCT")
            field_product.Orientation = PivotFieldOrientation.RowField
            Dim field_customer = pivottable.PivotFields("CUSTOMER")
            field_customer.Orientation = PivotFieldOrientation.RowField
            Dim field_amount = pivottable.PivotFields("AMOUNT")
            field_amount.Orientation = PivotFieldOrientation.DataField
            Dim field_quarter = pivottable.PivotFields("QUARTER")
            field_quarter.Orientation = PivotFieldOrientation.ColumnField

            'just a workaround, we dost not provide api to set number format
            'for pivot field. so use follow way.
            pivottable.DataBodyRange.NumberFormat = "$#,##0.00_);($#,##0.00)"
#End Region

#Region "Style"
            worksheet2.Range("1:1").Style = workbook.Styles("Title")
            worksheet2.Range("1:1").Interior.ThemeColor = ThemeColor.Accent1
            worksheet2.Range("1:1").HorizontalAlignment = HorizontalAlignment.Left
            worksheet2.Range("1:1").IndentLevel = 1
            worksheet2.Range("1:1").VerticalAlignment = VerticalAlignment.Center
            worksheet2.Range!A1.Style = workbook.Styles("Normal")
#End Region

#Region "Shape"
            'create a shape.
            Dim shape As IShape = worksheet2.Shapes.AddShape(AutoShapeType.RectangularCallout, 472, 65, 300, 70)

            'config shape's line and fill.
            shape.Line.DashStyle = LineDashStyle.Solid
            shape.Line.Weight = 4
            shape.Line.Color.RGB = Color.FromArgb(89, 89, 89)
            shape.Fill.Solid()
            shape.Fill.Color.ColorType = SolidColorType.None

            'config shape's rich text.
            Dim shape_p1 As ITextRange = shape.TextFrame.TextRange.Paragraphs(0)
            shape_p1.Text = "TIP:"
            Dim shape_p2 As ITextRange = shape.TextFrame.TextRange.Paragraphs.Add()
            shape_p2.Text = "Customize this PivotTable to fit your needs! Select a cell in the PivotTable to activate the PivotTable Field List pane. Then in the task pane, drag to add, remove, or reorder the fields." &
                " For example, drag the Product field above the Customer field for a different view. To update PivotTable data, right-click in the PivotTable and then click Refresh."

            'set first paragraph's font style.
            shape_p1.Font.ThemeFont = ThemeFont.Major
            shape_p1.Font.Bold = True
            shape_p1.Font.Size = 12
            shape_p1.Font.Color.ObjectThemeColor = ThemeColor.Dark1
            shape_p1.Font.Color.Brightness = 0.25

            'set second paragraph's font style.
            shape_p2.Font.ThemeFont = ThemeFont.Minor
            shape_p2.Font.Size = 8
            shape_p2.Font.Color.ObjectThemeColor = ThemeColor.Dark1
            shape_p2.Font.Color.Brightness = 0.25
#End Region

#End Region

#Region "Worksheet_DataLists"
            Dim worksheet3 As IWorksheet = workbook.Worksheets.Add()
            worksheet3.Name = "Data Lists"
            worksheet3.SheetView.DisplayGridlines = False

#Region "RowHeightColumnWidth"
            worksheet3.StandardHeight = 18.75
            worksheet3.StandardWidth = 8.43
            worksheet3.Range("1:1").RowHeight = 51.75
            worksheet3.Range("2:32").RowHeight = 19
            worksheet3.Range("A:A, D:D").ColumnWidth = 2.28515625
            worksheet3.Range("B:C").ColumnWidth = 34.42578125
#End Region

#Region "Values"
            worksheet3.Range!B1.Value = "DATA LISTS"
#End Region

#Region "Table"
            worksheet3.Range("B2:B32").Value = {
                {"PRODUCTS"},
                {"Product 1"},
                {"Product 2"},
                {"Product 3"},
                {"Product 4"},
                {"Product 5"},
                {"Product 6"},
                {"Product 7"},
                {"Product 8"},
                {"Product 9"},
                {"Product 10"},
                {"Product 11"},
                {"Product 12"},
                {"Product 13"},
                {"Product 14"},
                {"Product 15"},
                {"Product 16"},
                {"Product 17"},
                {"Product 18"},
                {"Product 19"},
                {"Product 20"},
                {"Product 21"},
                {"Product 22"},
                {"Product 23"},
                {"Product 24"},
                {"Product 25"},
                {"Product 26"},
                {"Product 27"},
                {"Product 28"},
                {"Product 29"},
                {"Product 30"}
            }

            Dim table_Products As ITable = worksheet3.Tables.Add(worksheet.Range("B2:B32"), True)
            worksheet3.Range("C2:C30").Value = {
                {"CUSTOMERS"},
                {"A. Datum Corporation"},
                {"Adventure Works"},
                {"Alpine Ski House"},
                {"Blue Yonder Airlines"},
                {"City Power & Light"},
                {"Coho Vineyard"},
                {"Coho Winery"},
                {"Coho Vineyard & Winery"},
                {"Contoso, Ltd"},
                {"Contoso Pharmaceuticals"},
                {"Consolidated Messenger"},
                {"Fabrikam, Inc."},
                {"Fourth Coffee"},
                {"Graphic Design Institute"},
                {"Humongous Insurance"},
                {"Litware, Inc."},
                {"Lucerne Publishing"},
                {"Margie's Travel"},
                {"Northwind Traders"},
                {"Proseware, Inc."},
                {"School of Fine Art"},
                {"Southridge Video"},
                {"Tailspin Toys"},
                {"Trey Research"},
                {"The Phone Company"},
                {"Wide World Importers"},
                {"Wingtip Toys"},
                {"Woodgrove Bank"}
            }

            Dim table_Customers As ITable = worksheet3.Tables.Add(worksheet.Range("C2:C30"), True)
            table_Products.TableStyle = workbook.TableStyles("TableStyleMedium2")
            table_Products.Columns(0).Range.HorizontalAlignment = HorizontalAlignment.Left
            table_Products.Columns(0).Range.IndentLevel = 1
            table_Products.Columns(0).Range.VerticalAlignment = VerticalAlignment.Center
            table_Products.HeaderRange.Font.Size = 11
            table_Customers.TableStyle = workbook.TableStyles("TableStyleMedium2")
            table_Customers.Columns(0).Range.HorizontalAlignment = HorizontalAlignment.Left
            table_Customers.Columns(0).Range.IndentLevel = 1
            table_Customers.Columns(0).Range.VerticalAlignment = VerticalAlignment.Center
            table_Customers.ShowTableStyleFirstColumn = True
            table_Customers.HeaderRange.Font.Size = 11
#End Region

#Region "Style"
            worksheet3.Range("1:1").Style = workbook.Styles("Title")
            worksheet3.Range("1:1").Interior.ThemeColor = ThemeColor.Accent1
            worksheet3.Range("1:1").HorizontalAlignment = HorizontalAlignment.Left
            worksheet3.Range("1:1").IndentLevel = 1
            worksheet3.Range("1:1").VerticalAlignment = VerticalAlignment.Center
            worksheet3.Range!A1.Style = workbook.Styles("Normal")
#End Region

#Region "Shape"
            Dim shape2 As IShape = worksheet3.Shapes.AddShape(AutoShapeType.RectangularCallout, 380, 65, 280, 50)
            shape2.Line.DashStyle = LineDashStyle.Solid
            shape2.Line.Weight = 4
            shape2.Line.Color.RGB = Color.FromArgb(89, 89, 89)
            shape2.Fill.Solid()
            shape2.Fill.Color.ColorType = SolidColorType.None
            Dim shape2_p1 As ITextRange = shape2.TextFrame.TextRange.Paragraphs(0)
            shape2_p1.Text = "TIP:"
            Dim shape2_p2 As ITextRange = shape2.TextFrame.TextRange.Paragraphs.Add()
            shape2_p2.Text = "To add a new product or customer, start typing below the table and it will automatically expand when you press the Enter or Tab key."
            shape2_p1.Font.ThemeFont = ThemeFont.Major
            shape2_p1.Font.Bold = True
            shape2_p1.Font.Size = 12
            shape2_p1.Font.Color.ObjectThemeColor = ThemeColor.Dark1
            shape2_p1.Font.Color.Brightness = 0.25
            shape2_p2.Font.ThemeFont = ThemeFont.Minor
            shape2_p2.Font.Size = 8
            shape2_p2.Font.Color.ObjectThemeColor = ThemeColor.Dark1
            shape2_p2.Font.Color.Brightness = 0.25
#End Region
#End Region

        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
