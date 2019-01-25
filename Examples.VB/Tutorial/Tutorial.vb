Imports System.Drawing
Imports System.IO
Imports System.Reflection
Public Class Tutorial
    Inherits ExampleBase
    Public Overrides ReadOnly Property CanDownload As Boolean
        Get
            Return False
        End Get
    End Property
    Public Overrides ReadOnly Property ShowCode As Boolean
        Get
            Return False
        End Get
    End Property
    Public Overrides ReadOnly Property ShowViewer As Boolean
        Get
            Return False
        End Get
    End Property
    Public Overrides Function GetDescriptionByCulture(culture As String) As String
        Dim resource As String = "GrapeCity.Documents.Excel.Examples.VB.Tutorial_" & culture & ".md"
        Dim assembly = [GetType]().GetTypeInfo().Assembly
        Using stream As Stream = assembly.GetManifestResourceStream(resource)
            Using reader As New StreamReader(stream)
                Return reader.ReadToEnd()
            End Using
        End Using
    End Function
    Public Overrides Sub Execute(workbook As Workbook)
        Dim worksheet As IWorksheet = workbook.Worksheets(0)
        '-----------------------------Set Value------------------------------
        worksheet.Range("B3:C7").Value = New Object(,)  {
            {"ITEM", "AMOUNT"},
            {"Income 1", 2500},
            {"Income 2", 1000},
            {"Income 3", 250},
            {"Other", 250}
        }
        worksheet.Range("B10:C23").Value = New Object(,)  {
            {"ITEM", "AMOUNT"},
            {"Rent/mortgage", 800},
            {"Electric", 120},
            {"Gas", 50},
            {"Cell phone", 45},
            {"Groceries", 500},
            {"Car payment", 273},
            {"Auto expenses", 120},
            {"Student loans", 50},
            {"Credit cards", 100},
            {"Auto Insurance", 78},
            {"Personal care", 50},
            {"Entertainment", 100},
            {"Miscellaneous", 50}
        }

        worksheet.Range("B2:C2").Merge()
        worksheet.Range!B2.Value = "MONTHLY INCOME"
        worksheet.Range("B9:C9").Merge()
        worksheet.Range!B9.Value = "MONTHLY EXPENSES"
        worksheet.Range("E2:G2").Merge()
        worksheet.Range!E2.Value = "PERCENTAGE OF INCOME SPENT"
        worksheet.Range("E5:G5").Merge()
        worksheet.Range!E5.Value = "SUMMARY"
        worksheet.Range("E3:F3").Merge()
        worksheet.Range!E9.Value = "BALANCE"
        worksheet.Range!E6.Value = "Total Monthly Income"
        worksheet.Range!E7.Value = "Total Monthly Expenses"

        '--------------------------------Set Height & Width--------------------------------
        worksheet.StandardHeight = 26.25
        worksheet.StandardWidth = 8.43
        worksheet.Range("2:24").RowHeight = 27
        worksheet.Range("A:A").ColumnWidth = 2.855
        worksheet.Range("B:B").ColumnWidth = 33.285
        worksheet.Range("C:C").ColumnWidth = 25.57
        worksheet.Range("D:D").ColumnWidth = 1
        worksheet.Range("E:F").ColumnWidth = 25.57
        worksheet.Range("G:G").ColumnWidth = 14.285

        '------------------------------Set Table--------------------------------------
        Dim incomeTable As ITable = worksheet.Tables.Add(worksheet.Range("B3:C7"), True)
        incomeTable.Name = "tblIncome"
        incomeTable.TableStyle = workbook.TableStyles("TableStyleMedium4")
        Dim expensesTable As ITable = worksheet.Tables.Add(worksheet.Range("B10:C23"), True)
        expensesTable.Name = "tblExpenses"
        expensesTable.TableStyle = workbook.TableStyles("TableStyleMedium4")

        '------------------------------Set Formulas-----------------------------------
        worksheet.Names.Add("TotalMonthlyIncome", "=SUM(tblIncome[AMOUNT])")
        worksheet.Names.Add("TotalMonthlyExpenses", "=SUM(tblExpenses[AMOUNT])")
        worksheet.Range!E3.Formula = "=TotalMonthlyExpenses"
        worksheet.Range!G3.Formula = "=TotalMonthlyExpenses/TotalMonthlyIncome"
        worksheet.Range!G6.Formula = "=TotalMonthlyIncome"
        worksheet.Range!G7.Formula = "=TotalMonthlyExpenses"
        worksheet.Range!G9.Formula = "=TotalMonthlyIncome-TotalMonthlyExpenses"

        '----------------------------Set Styles-------------------------
        Dim currencyStyle As IStyle = workbook.Styles("Currency")
        currencyStyle.IncludeAlignment = True
        currencyStyle.HorizontalAlignment = HorizontalAlignment.Left
        currencyStyle.VerticalAlignment = VerticalAlignment.Bottom
        currencyStyle.NumberFormat = "$#,##0.00"

        Dim heading1Style As IStyle = workbook.Styles("Heading 1")
        heading1Style.IncludeAlignment = True
        heading1Style.HorizontalAlignment = HorizontalAlignment.Center
        heading1Style.VerticalAlignment = VerticalAlignment.Center
        heading1Style.Font.Name = "Century Gothic"
        heading1Style.Font.Bold = True
        heading1Style.Font.Size = 11
        heading1Style.Font.Color = Color.White
        heading1Style.IncludeBorder = False
        heading1Style.IncludePatterns = True
        heading1Style.Interior.Color = Color.FromArgb(32, 61, 64)

        Dim percentStyle As IStyle = workbook.Styles("Percent")
        percentStyle.IncludeAlignment = True
        percentStyle.HorizontalAlignment = HorizontalAlignment.Center
        percentStyle.IncludeFont = True
        percentStyle.Font.Color = Color.FromArgb(32, 61, 64)
        percentStyle.Font.Name = "Century Gothic"
        percentStyle.Font.Bold = True
        percentStyle.Font.Size = 14
        worksheet.SheetView.DisplayGridlines = False
        worksheet.Range("C4:C7, C11:C23, G6:G7, G9").Style = currencyStyle
        worksheet.Range("B2, B9, E2, E5").Style = heading1Style
        worksheet.Range!G3.Style = percentStyle
        worksheet.Range("E6:G6").Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Medium
        worksheet.Range("E6:G6").Borders(BordersIndex.EdgeBottom).Color = Color.FromArgb(32, 61, 64)
        worksheet.Range("E7:G7").Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Medium
        worksheet.Range("E7:G7").Borders(BordersIndex.EdgeBottom).Color = Color.FromArgb(32, 61, 64)
        worksheet.Range("E9:G9").Interior.Color = Color.FromArgb(32, 61, 64)
        worksheet.Range("E9:G9").HorizontalAlignment = HorizontalAlignment.Left
        worksheet.Range("E9:G9").VerticalAlignment = VerticalAlignment.Center
        worksheet.Range("E9:G9").Font.Name = "Century Gothic"
        worksheet.Range("E9:G9").Font.Bold = True
        worksheet.Range("E9:G9").Font.Size = 11
        worksheet.Range("E9:G9").Font.Color = Color.White
        worksheet.Range("E3:F3").Borders.Color = Color.FromArgb(32, 61, 64)

        '----------------------------Set Conditional Format-------------------------
        Dim dataBar As IDataBar = worksheet.Range!E3.FormatConditions.AddDatabar()
        dataBar.MinPoint.Type = ConditionValueTypes.Number
        dataBar.MinPoint.Value = 1
        dataBar.MaxPoint.Type = ConditionValueTypes.Number
        dataBar.MaxPoint.Value = "=TotalMonthlyIncome"
        dataBar.BarFillType = DataBarFillType.Gradient
        dataBar.BarColor.Color = Color.Red
        dataBar.ShowValue = False

        '--------------------------------Set Shape--------------------------------
        Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 339, 247, 316.5, 346)
        shape.Chart.ChartArea.Format.Line.Transparency = 1
        shape.Chart.ColumnGroups(0).Overlap = 0
        shape.Chart.ColumnGroups(0).GapWidth = 37

        Dim category_axis As IAxis = shape.Chart.Axes.Item(AxisType.Category)
        category_axis.Format.Line.Color.RGB = Color.Black
        category_axis.TickLabels.Font.Size = 11
        category_axis.TickLabels.Font.Color.RGB = Color.Black

        Dim series_axis As IAxis = shape.Chart.Axes.Item(AxisType.Value)
        series_axis.Format.Line.Weight = 1
        series_axis.Format.Line.Color.RGB = Color.Black
        series_axis.TickLabels.NumberFormat = "$###0"
        series_axis.TickLabels.Font.Size = 11
        series_axis.TickLabels.Font.Color.RGB = Color.Black

        Dim chartSeries As ISeries = shape.Chart.SeriesCollection.NewSeries()
        chartSeries.Formula = "=SERIES(""Simple Budget"",{""Income"",""Expenses""},'Sheet1'!$G$6:$G$7,1)"
        chartSeries.Points(0).Format.Fill.Color.RGB = Color.FromArgb(176, 21, 19)
        chartSeries.Points(1).Format.Fill.Color.RGB = Color.FromArgb(234, 99, 18)
        chartSeries.DataLabels.Font.Size = 11
        chartSeries.DataLabels.Font.Color.RGB = Color.Black
        chartSeries.DataLabels.ShowValue = True
        chartSeries.DataLabels.Position = DataLabelPosition.OutsideEnd
    End Sub
End Class
