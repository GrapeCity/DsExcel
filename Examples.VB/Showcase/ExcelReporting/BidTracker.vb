Imports System.Drawing

Namespace Showcase
    Public Class BidTracker
        Inherits ExampleBase

        Protected Overrides Sub BeforeExecute(workbook As Workbook, userAgents() As String)
            If AgentIsMac(userAgents) Then
                Dim themes As New Themes
                Dim theme As ITheme = themes.Add("testTheme", Themes.OfficeTheme)
                theme.ThemeFontScheme.Minor(FontLanguageIndex.Latin).Name = "Trebuchet MS"
                workbook.Theme = theme
                Dim style_Normal = workbook.Styles("Normal")
                style_Normal.Font.ThemeFont = ThemeFont.Minor
            End If
        End Sub

        Public Overrides Sub Execute(workbook As Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            '***********************Set RowHeight & ColumnWidth***************
            worksheet.StandardHeight = 30
            worksheet.Range("1:1").RowHeight = 57.75
            worksheet.Range("2:9").RowHeight = 30.25
            worksheet.Range("A:A").ColumnWidth = 2.71
            worksheet.Range("B:B").ColumnWidth = 11.71
            worksheet.Range("C:C").ColumnWidth = 28
            worksheet.Range("D:D").ColumnWidth = 22.425
            worksheet.Range("E:E").ColumnWidth = 16.71
            worksheet.Range("F:F").ColumnWidth = 28
            worksheet.Range("G:H").ColumnWidth = 16.71
            worksheet.Range("I:I").ColumnWidth = 2.71



            '**************************Set Table Value & Formulas*********************
            Dim table As ITable = worksheet.Tables.Add(worksheet.Range("B2:H9"), True)
            worksheet.Range("B2:H9").Value = New Object(,) {
                {"BID #", "DESCRIPTION", "DATE RECEIVED", "AMOUNT", "PERCENT COMPLETE", "DEADLINE", "DAYS LEFT"},
                {1, "Bid number 1", Nothing, 2000, 0.5, Nothing, Nothing},
                {2, "Bid number 2", Nothing, 3500, 0.25, Nothing, Nothing},
                {3, "Bid number 3", Nothing, 5000, 0.3, Nothing, Nothing},
                {4, "Bid number 4", Nothing, 4000, 0.2, Nothing, Nothing},
                {5, "Bid number 5", Nothing, 4000, 0.75, Nothing, Nothing},
                {6, "Bid number 6", Nothing, 1500, 0.45, Nothing, Nothing},
                {7, "Bid number 7", Nothing, 5000, 0.65, Nothing, Nothing}
            }
            worksheet.Range!B1.Value = "Bid Details"
            worksheet.Range!D3.Formula = "=TODAY()-10"
            worksheet.Range("D4:D5").Formula = "=TODAY()-20"
            worksheet.Range!D6.Formula = "=TODAY()-10"
            worksheet.Range!D7.Formula = "=TODAY()-28"
            worksheet.Range!D8.Formula = "=TODAY()-17"
            worksheet.Range!D9.Formula = "=TODAY()-15"
            worksheet.Range("G3:G9").Formula = "=[@[DATE RECEIVED]]+30"
            worksheet.Range("H3:H9").Formula = "=[@DEADLINE]-TODAY()"



            '****************************Set Table Style********************************
            Dim tableStyle As ITableStyle = workbook.TableStyles.Add("Bid Tracker")
            workbook.DefaultTableStyle = "Bid Tracker"

            'Set WholeTable element style.
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Font.Color = Color.FromArgb(89, 89, 89)
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders.Color = Color.FromArgb(89, 89, 89)
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeLeft).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeRight).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeTop).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.InsideVertical).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.InsideHorizontal).LineStyle = BorderLineStyle.Thin

            'Set HeaderRow element style.
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders.Color = Color.FromArgb(89, 89, 89)
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders(BordersIndex.EdgeLeft).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders(BordersIndex.EdgeRight).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders(BordersIndex.EdgeTop).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders(BordersIndex.InsideVertical).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders(BordersIndex.InsideHorizontal).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Interior.Color = Color.FromArgb(131, 95, 1)
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Interior.PatternColor = Color.FromArgb(254, 184, 10)


            'Set TotalRow element style.
            tableStyle.TableStyleElements(TableStyleElementType.TotalRow).Borders.Color = Color.White
            tableStyle.TableStyleElements(TableStyleElementType.TotalRow).Borders(BordersIndex.EdgeLeft).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.TotalRow).Borders(BordersIndex.EdgeRight).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.TotalRow).Borders(BordersIndex.EdgeTop).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.TotalRow).Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.TotalRow).Borders(BordersIndex.InsideVertical).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.TotalRow).Borders(BordersIndex.InsideHorizontal).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.TotalRow).Interior.Color = Color.FromArgb(131, 95, 1)



            '***********************************Set Named Styles*****************************
            Dim titleStyle As IStyle = workbook.Styles("Title")
            titleStyle.Font.Name = "Trebuchet MS"
            titleStyle.Font.Size = 36
            titleStyle.Font.Color = Color.FromArgb(56, 145, 167)
            titleStyle.IncludeAlignment = True
            titleStyle.VerticalAlignment = VerticalAlignment.Center


            Dim heading1Style As IStyle = workbook.Styles("Heading 1")
            heading1Style.IncludeAlignment = True
            heading1Style.HorizontalAlignment = HorizontalAlignment.Left
            heading1Style.IndentLevel = 1
            heading1Style.VerticalAlignment = VerticalAlignment.Bottom
            heading1Style.Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.None
            heading1Style.Font.Size = 14
            heading1Style.Font.Color = Color.White
            heading1Style.Font.Bold = False
            heading1Style.IncludePatterns = True
            heading1Style.Interior.Color = Color.FromArgb(131, 95, 1)
            heading1Style.Font.Name = "Trebuchet MS"


            Dim dateStyle As IStyle = workbook.Styles.Add("Date")
            dateStyle.IncludeNumber = True
            dateStyle.NumberFormat = "m/d/yyyy"
            dateStyle.IncludeAlignment = True
            dateStyle.HorizontalAlignment = HorizontalAlignment.Left
            dateStyle.IndentLevel = 1
            dateStyle.VerticalAlignment = VerticalAlignment.Center
            dateStyle.IncludeFont = False
            dateStyle.IncludeBorder = False
            dateStyle.IncludePatterns = False
            dateStyle.Font.Name = "Trebuchet MS"


            Dim commaStyle As IStyle = workbook.Styles("Comma")
            commaStyle.IncludeNumber = True
            commaStyle.NumberFormat = "#,##0_);(#,##0)"
            commaStyle.IncludeAlignment = True
            commaStyle.HorizontalAlignment = HorizontalAlignment.Left
            commaStyle.IndentLevel = 1
            commaStyle.VerticalAlignment = VerticalAlignment.Center
            commaStyle.Font.Name = "Trebuchet MS"


            Dim normalStyle As IStyle = workbook.Styles("Normal")
            normalStyle.HorizontalAlignment = HorizontalAlignment.Left
            normalStyle.IndentLevel = 1
            normalStyle.VerticalAlignment = VerticalAlignment.Center
            normalStyle.WrapText = True
            normalStyle.Font.Color = Color.FromArgb(89, 89, 89)


            Dim currencyStyle As IStyle = workbook.Styles("Currency")
            currencyStyle.NumberFormat = "$#,##0.00"
            currencyStyle.IncludeAlignment = True
            currencyStyle.HorizontalAlignment = HorizontalAlignment.Left
            currencyStyle.IndentLevel = 1
            currencyStyle.VerticalAlignment = VerticalAlignment.Center
            currencyStyle.Font.Name = "Trebuchet MS"


            Dim percentStyle As IStyle = workbook.Styles("Percent")
            percentStyle.IncludeAlignment = True
            percentStyle.HorizontalAlignment = HorizontalAlignment.Right
            percentStyle.VerticalAlignment = VerticalAlignment.Center
            percentStyle.IncludeFont = True
            percentStyle.Font.Name = "Trebuchet MS"
            percentStyle.Font.Size = 20
            percentStyle.Font.Bold = True
            percentStyle.Font.Color = Color.FromArgb(89, 89, 89)
            percentStyle.Font.Name = "Trebuchet MS"


            Dim comma0Style As IStyle = workbook.Styles("Comma [0]")
            comma0Style.NumberFormat = "#,##0_);(#,##0)"
            comma0Style.IncludeAlignment = True
            comma0Style.HorizontalAlignment = HorizontalAlignment.Right
            comma0Style.IndentLevel = 3
            comma0Style.VerticalAlignment = VerticalAlignment.Center
            percentStyle.Font.Name = "Trebuchet MS"



            '************************************Add Conditional Formatting****************
            Dim dataBar As IDataBar = worksheet.Range("F3:F9").FormatConditions.AddDatabar()
            dataBar.MinPoint.Type = ConditionValueTypes.Number
            dataBar.MinPoint.Value = 1
            dataBar.MaxPoint.Type = ConditionValueTypes.Number
            dataBar.MaxPoint.Value = 0

            dataBar.BarFillType = DataBarFillType.Gradient
            dataBar.BarColor.Color = Color.FromArgb(126, 194, 211)
            dataBar.Direction = DataBarDirection.Context

            dataBar.AxisColor.Color = Color.Black
            dataBar.AxisPosition = DataBarAxisPosition.Automatic

            dataBar.NegativeBarFormat.ColorType = DataBarNegativeColorType.Color
            dataBar.NegativeBarFormat.Color.Color = Color.Red
            dataBar.ShowValue = True



            '****************************************Use NamedStyle**************************
            worksheet.SheetView.DisplayGridlines = False
            table.TableStyle = tableStyle
            worksheet.Range!B1.Style = titleStyle
            worksheet.Range!B1.WrapText = False
            worksheet.Range("B2:H2").Style = heading1Style
            worksheet.Range("B3:B9").Style = commaStyle
            worksheet.Range("C3:C9").Style = normalStyle
            worksheet.Range("D3:D9").Style = dateStyle
            worksheet.Range("E3:E9").Style = currencyStyle
            worksheet.Range("F3:F9").Style = percentStyle
            worksheet.Range("G3:G9").Style = dateStyle
            worksheet.Range("H3:H9").Style = comma0Style

        End Sub

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
