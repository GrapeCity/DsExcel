# Getting started

Spread.Services is a new high performing, low memory server component with full API for server-side spreadsheet generation, manipulation, and serialization to various formats including xlsx and ssjson. Spread.Services targets **.NET Standard 1.4** for multi-platform support including: **.NET Framework**, **.NET Core**, and **Mono**.

In this tutorial, we will create a real scenario with Spread.Services to give you a fundamental understanding of how Spread.Services works, in the end of this tutorial, you will get a simple budget excel file.

## Prepare
1. Create a .Net Core Console Application (we use .net core in this tutorial, you can do similar ways on .Net Framework or Xamrin project).
2. Edit references of the project to install Spread.Services nuget package, there are two ways:
> In **Visutal Studio**
> > - Right click "Dependencies", then click "Manage NuGet Packages...".
> > - Select nuget.org as package source, search "Spread.Services", then click "Install".
> 
> or just through **dotnet CLI** 
> > ```csharp
> > dotnet add package Spread.Services
> > ```

## Create Workbook

```csharp
Workbook workbook = new Workbook();
IWorksheet worksheet = workbook.Worksheets[0];
```

## Initialize Data

It is very easy and efficient to initialize data in **Spread.Services**, just needs to prepare a two-dimension array, and assign it to range's value.

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

## Set Row Height And Column Width

 To beautify layout and fit data presentation, we will customize row height and column width. "StandardHeight" represents the default row height, well "StandardWidth" represents the default column width of worksheet.


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

## Create Table

We will add two tables, one is "Income" table, the other is "Expenses" table, and give them a built-in table style.

```csharp
ITable incomeTable = worksheet.Tables.Add(worksheet.Range["B3:C7"], true);
incomeTable.Name = "tblIncome";
incomeTable.TableStyle = workbook.TableStyles["TableStyleMedium4"];

ITable expensesTable = worksheet.Tables.Add(worksheet.Range["B10:C23"], true);
expensesTable.Name = "tblExpenses";
expensesTable.TableStyle = workbook.TableStyles["TableStyleMedium4"];
```

## Set Formulas

We create two custom names to represent "total month income" and "total month expenses", then set formulas to range to calculate the result of "Percentage of income spent", "Total monthly income", "Total monthly expenses", "Balance".


```csharp
worksheet.Names.Add("TotalMonthlyIncome", "=SUM(tblIncome[AMOUNT])");
worksheet.Names.Add("TotalMonthlyExpenses", "=SUM(tblExpenses[AMOUNT])");

worksheet.Range["E3"].Formula = "=TotalMonthlyExpenses";
worksheet.Range["G3"].Formula = "=TotalMonthlyExpenses/TotalMonthlyIncome";
worksheet.Range["G6"].Formula = "=TotalMonthlyIncome";
worksheet.Range["G7"].Formula = "=TotalMonthlyExpenses";
worksheet.Range["G9"].Formula = "=TotalMonthlyIncome-TotalMonthlyExpenses";
```


## Set Styles

There are two ways to change range styles. 
- Apply a built-in or custom name style to range
- Set range's styles directly

In this step, we modify "Currency", "Heading 1", "Percent" built-in name style, and then apply them to some ranges, For the other ranges, we modify their styles directly. 


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
heading1Style.Interior.Color = Color.FromRGB(32, 61, 64);

IStyle percentStyle = workbook.Styles["Percent"];
percentStyle.IncludeAlignment = true;
percentStyle.HorizontalAlignment = HorizontalAlignment.Center;
percentStyle.IncludeFont = true;
percentStyle.Font.Color = Color.FromRGB(32, 61, 64);
percentStyle.Font.Name = "Century Gothic";
percentStyle.Font.Bold = true;
percentStyle.Font.Size = 14;

worksheet.SheetView.DisplayGridlines = false;
worksheet.Range["C4:C7, C11:C23, G6:G7, G9"].Style = currencyStyle;
worksheet.Range["B2, B9, E2, E5"].Style = heading1Style;
worksheet.Range["G3"].Style = percentStyle;

worksheet.Range["E6:G6"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Medium;
worksheet.Range["E6:G6"].Borders[BordersIndex.EdgeBottom].Color = Color.FromRGB(32, 61, 64);
worksheet.Range["E7:G7"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Medium;
worksheet.Range["E7:G7"].Borders[BordersIndex.EdgeBottom].Color = Color.FromRGB(32, 61, 64);

worksheet.Range["E9:G9"].Interior.Color = Color.FromRGB(32, 61, 64);
worksheet.Range["E9:G9"].HorizontalAlignment = HorizontalAlignment.Left;
worksheet.Range["E9:G9"].VerticalAlignment = VerticalAlignment.Center;
worksheet.Range["E9:G9"].Font.Name = "Century Gothic";
worksheet.Range["E9:G9"].Font.Bold = true;
worksheet.Range["E9:G9"].Font.Size = 11;
worksheet.Range["E9:G9"].Font.Color = Color.White;
worksheet.Range["E3:F3"].Borders.Color = Color.FromRGB(32, 61, 64);
```


## Add Conditional Format

Spread.Services supports all types of conditional format rule, In this step, we create a gradient data bar rule to show the percentage of income spent, and just shows data bar, not show value.

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

## Add Chart 
To show the gap between income and expenses visually, we can create a column chart, in order to beautify its layout, we change series overlap and gap width, then customize formatting for some chart elements, such as chart area, axis line, tick labels font, data points fill and data labels' font.


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
chartSeries.Points[0].Format.Fill.Color.RGB = Color.FromRGB(176, 21, 19);
chartSeries.Points[1].Format.Fill.Color.RGB = Color.FromRGB(234, 99, 18);
chartSeries.DataLabels.Font.Size = 11;
chartSeries.DataLabels.Font.Color.RGB = Color.Black;
chartSeries.DataLabels.ShowValue = true;
chartSeries.DataLabels.Position = DataLabelPosition.OutsideEnd;

```

## Save to Excel

Finnaly, save it to an excel file named 'SimpleBudget.xlsx'

```csharp
workbook.Save(@"SimpleBudget.xlsx");
```

Then you can view the saved [SimpleBudget.xlsx](api/examples/xlsx/GrapeCity.Documents.Spread.Examples.Tutorial?fileName=SimpleBudget).
