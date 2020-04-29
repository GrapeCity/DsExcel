# Getting started with Documents for Excel, a spreadsheet API

In this tutorial, we create a real-life scenario with GrapeCity Documents for Excel to give you a fundamental understanding of what it can do. At the end of this tutorial, you will have a simple budget Excel file.

## Prepare

1. Install [.NET Core](https://www.microsoft.com/net/core). This tutorial uses .NET Core, but you can use similar methods in .NET Framework and Mono projects.

2. Create a .NET Core Console Application in **Visual Studio**, or just use the **dotnet CLI**.
> ```csharp
> dotnet new console
> ```

3. Install the **GrapeCity Documents for Excel** nuget package using Visual Studio or the dotnet CLI:
> **Visual Studio**
> - Right-click the project file, then click "Manage NuGet Packages."
> - Select **nuget.org** as the package source, and search for "GrapeCity.Documents.Excel" Click "Install."
>
> **dotnet CLI** 
> - Open a cmd window under the project folder.
> - Execute this command.:
> ```csharp
> dotnet add package GrapeCity.Documents.Excel
> ```

## Add Namespace

Open Program.cs and add these three namespaces.

- C#
```csharp
using System.Drawing;
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Drawing; 
```

- VB
```vbnet
Imports System.Drawing;
Imports GrapeCity.Documents.Excel;
Imports GrapeCity.Documents.Excel.Drawing; 
```

## Create Workbook

The first step in creating an Excel file with the GrapeCity Documents for Excel API is to create a new Workbook.

- C#
```csharp
Workbook workbook = new Workbook();
IWorksheet worksheet = workbook.Worksheets[0];
```

- VB
```vbnet
Dim workbook As Workbook= new Workbook
Dim worksheet As IWorksheet = workbook.Worksheets(0)
```

## Initialize Data

To initialize data in **GrapeCity Documents for Excel**, prepare a two-dimensional array and assign it to the Value of a worksheet Range.

- C#
```csharp
worksheet.Range["B3:C7"].Value = new object[,]
{
    { "ITEM", "AMOUNT" },
    { "Income 1", 2500 },
    { "Income 2", 1000 },
    { "Income 3", 250 },
    { "Other", 250 },
};

worksheet.Range["B10:C23"].Value = new object[,]
{
    { "ITEM", "AMOUNT" },
    { "Rent/mortgage", 800 },
    { "Electric", 120 },
    { "Gas", 50 },
    { "Cell phone", 45 },
    { "Groceries", 500 },
    { "Car payment", 273 },
    { "Auto expenses", 120 },
    { "Student loans", 50 },
    { "Credit cards", 100 },
    { "Auto Insurance", 78 },
    { "Personal care", 50 },
    { "Entertainment", 100 },
    { "Miscellaneous", 50 },
};

worksheet.Range["B2:C2"].Merge();
worksheet.Range["B2"].Value = "MONTHLY INCOME";
worksheet.Range["B9:C9"].Merge();
worksheet.Range["B9"].Value = "MONTHLY EXPENSES";
worksheet.Range["E2:G2"].Merge();
worksheet.Range["E2"].Value = "PERCENTAGE OF INCOME SPENT";
worksheet.Range["E5:G5"].Merge();
worksheet.Range["E5"].Value = "SUMMARY";
worksheet.Range["E3:F3"].Merge();
worksheet.Range["E9"].Value = "BALANCE";
worksheet.Range["E6"].Value = "Total Monthly Income";
worksheet.Range["E7"].Value = "Total Monthly Expenses";
```

- VB
```vbnet
worksheet.Range("B3:C7").Value = {
    {"ITEM", "AMOUNT"},
    {"Income 1", 2500},
    {"Income 2", 1000},
    {"Income 3", 250},
    {"Other", 250}
}
worksheet.Range("B10:C23").Value = {
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
```

## Set Row Heights and Column Widths

Customize row heights and column widths to polish the layout and data presentation. Use "StandardHeight" and "StandardWidth" to set the default row height and column width for the worksheet.

- C#
```csharp
worksheet.StandardHeight = 26.25;
worksheet.StandardWidth = 8.43;

worksheet.Range["2:24"].RowHeight = 27;
worksheet.Range["A:A"].ColumnWidth = 2.855;
worksheet.Range["B:B"].ColumnWidth = 33.285;
worksheet.Range["C:C"].ColumnWidth = 25.57;
worksheet.Range["D:D"].ColumnWidth = 1;
worksheet.Range["E:F"].ColumnWidth = 25.57;
worksheet.Range["G:G"].ColumnWidth = 14.285;
```

- VB
```vbnet
worksheet.StandardHeight = 26.25
worksheet.StandardWidth = 8.43

worksheet.Range("2:24").RowHeight = 27
worksheet.Range("A:A").ColumnWidth = 2.855
worksheet.Range("B:B").ColumnWidth = 33.285
worksheet.Range("C:C").ColumnWidth = 25.57
worksheet.Range("D:D").ColumnWidth = 1
worksheet.Range("E:F").ColumnWidth = 25.57
worksheet.Range("G:G").ColumnWidth = 14.285
```

## Create Table

Add two tables: "Income" and "Expenses," and apply a built-in table style to each.

- C#
```csharp
ITable incomeTable = worksheet.Tables.Add(worksheet.Range["B3:C7"], true);
incomeTable.Name = "tblIncome";
incomeTable.TableStyle = workbook.TableStyles["TableStyleMedium4"];

ITable expensesTable = worksheet.Tables.Add(worksheet.Range["B10:C23"], true);
expensesTable.Name = "tblExpenses";
expensesTable.TableStyle = workbook.TableStyles["TableStyleMedium4"];
```

- VB
```vbnet
Dim incomeTable As ITable = worksheet.Tables.Add(worksheet.Range("B3:C7"), True)
incomeTable.Name = "tblIncome"
incomeTable.TableStyle = workbook.TableStyles("TableStyleMedium4")

Dim expensesTable As ITable = worksheet.Tables.Add(worksheet.Range("B10:C23"), True)
expensesTable.Name = "tblExpenses"
expensesTable.TableStyle = workbook.TableStyles("TableStyleMedium4")
```

## Set Formulas

Create two custom names to summarize the income and expenses for the month, then add formulas that calculate the total monthly income, total monthly expenses, percentage of income spent, and balance.

- C#
```csharp
worksheet.Names.Add("TotalMonthlyIncome", "=SUM(tblIncome[AMOUNT])");
worksheet.Names.Add("TotalMonthlyExpenses", "=SUM(tblExpenses[AMOUNT])");

worksheet.Range["E3"].Formula = "=TotalMonthlyExpenses";
worksheet.Range["G3"].Formula = "=TotalMonthlyExpenses/TotalMonthlyIncome";
worksheet.Range["G6"].Formula = "=TotalMonthlyIncome";
worksheet.Range["G7"].Formula = "=TotalMonthlyExpenses";
worksheet.Range["G9"].Formula = "=TotalMonthlyIncome-TotalMonthlyExpenses";
```

- VB
```vbnet
worksheet.Names.Add("TotalMonthlyIncome", "=SUM(tblIncome[AMOUNT])")
worksheet.Names.Add("TotalMonthlyExpenses", "=SUM(tblExpenses[AMOUNT])")

worksheet.Range!E3.Formula = "=TotalMonthlyExpenses"
worksheet.Range!G3.Formula = "=TotalMonthlyExpenses/TotalMonthlyIncome"
worksheet.Range!G6.Formula = "=TotalMonthlyIncome"
worksheet.Range!G7.Formula = "=TotalMonthlyExpenses"
worksheet.Range!G9.Formula = "=TotalMonthlyIncome-TotalMonthlyExpenses"
```

## Set Styles

There are two ways to change range styles. 
- Apply a built-in or custom style by name
- Set individual styles for each element

Modify the "Currency," "Heading 1," and "Percent" built-in styles, and apply them to ranges of cells. Modify individual style elements for other ranges.

- C#
```csharp
IStyle currencyStyle = workbook.Styles["Currency"];
currencyStyle.IncludeAlignment = true;
currencyStyle.HorizontalAlignment = HorizontalAlignment.Left;
currencyStyle.VerticalAlignment = VerticalAlignment.Bottom;
currencyStyle.NumberFormat = "$#,##0.00";

IStyle heading1Style = workbook.Styles["Heading 1"];
heading1Style.IncludeAlignment = true;
heading1Style.HorizontalAlignment = HorizontalAlignment.Center;
heading1Style.VerticalAlignment = VerticalAlignment.Center;
heading1Style.Font.Name = "Century Gothic";
heading1Style.Font.Bold = true;
heading1Style.Font.Size = 11;
heading1Style.Font.Color = Color.White;
heading1Style.IncludeBorder = false;
heading1Style.IncludePatterns = true;
heading1Style.Interior.Color = Color.FromArgb(32, 61, 64);

IStyle percentStyle = workbook.Styles["Percent"];
percentStyle.IncludeAlignment = true;
percentStyle.HorizontalAlignment = HorizontalAlignment.Center;
percentStyle.IncludeFont = true;
percentStyle.Font.Color = Color.FromArgb(32, 61, 64);
percentStyle.Font.Name = "Century Gothic";
percentStyle.Font.Bold = true;
percentStyle.Font.Size = 14;

worksheet.SheetView.DisplayGridlines = false;
worksheet.Range["C4:C7, C11:C23, G6:G7, G9"].Style = currencyStyle;
worksheet.Range["B2, B9, E2, E5"].Style = heading1Style;
worksheet.Range["G3"].Style = percentStyle;

worksheet.Range["E6:G6"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Medium;
worksheet.Range["E6:G6"].Borders[BordersIndex.EdgeBottom].Color = Color.FromArgb(32, 61, 64);
worksheet.Range["E7:G7"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Medium;
worksheet.Range["E7:G7"].Borders[BordersIndex.EdgeBottom].Color = Color.FromArgb(32, 61, 64);

worksheet.Range["E9:G9"].Interior.Color = Color.FromArgb(32, 61, 64);
worksheet.Range["E9:G9"].HorizontalAlignment = HorizontalAlignment.Left;
worksheet.Range["E9:G9"].VerticalAlignment = VerticalAlignment.Center;
worksheet.Range["E9:G9"].Font.Name = "Century Gothic";
worksheet.Range["E9:G9"].Font.Bold = true;
worksheet.Range["E9:G9"].Font.Size = 11;
worksheet.Range["E9:G9"].Font.Color = Color.White;
worksheet.Range["E3:F3"].Borders.Color = Color.FromArgb(32, 61, 64);
```

- VB
```vbnet
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
```


## Add Conditional Formatting

GrapeCity Documents for Excel supports all types of conditional format rules. Create a gradient data bar rule to show the percentage of income spent. The rule shows a data bar without showing a value.

- C#
```csharp
IDataBar dataBar = worksheet.Range["E3"].FormatConditions.AddDatabar();
dataBar.MinPoint.Type = ConditionValueTypes.Number;
dataBar.MinPoint.Value = 1;
dataBar.MaxPoint.Type = ConditionValueTypes.Number;
dataBar.MaxPoint.Value = "=TotalMonthlyIncome";
dataBar.BarFillType = DataBarFillType.Gradient;
dataBar.BarColor.Color = Color.Red;
dataBar.ShowValue = false;
```

- VB
```vbnet
Dim dataBar As IDataBar = worksheet.Range!E3.FormatConditions.AddDatabar()
dataBar.MinPoint.Type = ConditionValueTypes.Number
dataBar.MinPoint.Value = 1
dataBar.MaxPoint.Type = ConditionValueTypes.Number
dataBar.MaxPoint.Value = "=TotalMonthlyIncome"
dataBar.BarFillType = DataBarFillType.Gradient
dataBar.BarColor.Color = Color.Red
dataBar.ShowValue = False
```

## Add Chart 

Create a column chart to illustrate the gap between income and expenses. To polish the layout, change the series overlap and gap width, then customize the formatting of some of the chart elements: chart area, axis line, tick labels and data points.

- C#
```csharp
IShape shape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 339, 247, 316.5, 346);
shape.Chart.ChartArea.Format.Line.Transparency = 1;
shape.Chart.ColumnGroups[0].Overlap = 0;
shape.Chart.ColumnGroups[0].GapWidth = 37;

IAxis category_axis = shape.Chart.Axes.Item(AxisType.Category);
category_axis.Format.Line.Color.RGB = Color.Black;
category_axis.TickLabels.Font.Size = 11;
category_axis.TickLabels.Font.Color.RGB = Color.Black;

IAxis series_axis = shape.Chart.Axes.Item(AxisType.Value);
series_axis.Format.Line.Weight = 1;
series_axis.Format.Line.Color.RGB = Color.Black;
series_axis.TickLabels.NumberFormat = "$###0";
series_axis.TickLabels.Font.Size = 11;
series_axis.TickLabels.Font.Color.RGB = Color.Black;

ISeries chartSeries = shape.Chart.SeriesCollection.NewSeries();
chartSeries.Formula = "=SERIES(\"Simple Budget\",{\"Income\",\"Expenses\"},'Sheet1'!$G$6:$G$7,1)";
chartSeries.Points[0].Format.Fill.Color.RGB = Color.FromArgb(176, 21, 19);
chartSeries.Points[1].Format.Fill.Color.RGB = Color.FromArgb(234, 99, 18);
chartSeries.DataLabels.Font.Size = 11;
chartSeries.DataLabels.Font.Color.RGB = Color.Black;
chartSeries.DataLabels.ShowValue = true;
chartSeries.DataLabels.Position = DataLabelPosition.OutsideEnd;
```

- VB
```vbnet
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
```

## Save to Excel

Save it to an Excel file named "SimpleBudget.xlsx."

- C#
```csharp
workbook.Save("SimpleBudget.xlsx");
```
- VB
```vbnet
workbook.Save("SimpleBudget.xlsx")
```

You can download and view the saved [SimpleBudget.xlsx](api/examples/xlsx/tutorial?fileName=SimpleBudget). If you prefer to download the [Tutorial Source Project](GrapeCity.Documents.Excel.Tutorial.zip) and run the code yourself, be sure to first install [.NET Core](https://www.microsoft.com/net/core) on your machine. 