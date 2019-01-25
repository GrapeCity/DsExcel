# DioDocs for Excel スプレッドシート API の開始

このチュートリアルでは、実例を作成しながら、DioDocs for Excel の機能について基本的な知識を習得します。このチュートリアルを完了すると、簡単な予算 Excel ファイルが完成します。

## 準備

1. [.NET Core](https://www.microsoft.com/net/core) をインストールします。このチュートリアルでは .NET Core を使用しますが、.NET Framework や Mono のプロジェクトでも同様の方法を使用できます。

2. **Visual Studio** で .NET Core コンソールアプリケーションを作成します。または、**dotnet CLI** を使用します。
> ```csharp
> dotnet new console
> ```

3. Visual Studio または dotnet CLI を使用して、**DioDocs for Excel** NuGet パッケージをインストールします。
> **Visual Studio**
> - プロジェクトファイルを右クリックし、［NuGet パッケージの管理］をクリックします。
> - パッケージソースとして **nuget.org** を選択し、「GrapeCity.Documents.Excel」を検索し、［インストール］をクリックします。
>
> **dotnet CLI** 
> - プロジェクトフォルダでコマンドウィンドウを開きます。
> - 次のコマンドを実行します。
> ```csharp
> dotnet add package GrapeCity.Documents.Excel
> ```

## 名前空間の追加

Program.cs を開き、次の 2 つの名前空間を追加します。

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

## ワークブックの作成

DioDocs for Excel API を使用して Excel ファイルを作成する最初の手順として、新しいワークブックを作成します。

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

## データの初期化

**DioDocs for Excel** でデータを初期化するには、2 次元配列を用意し、それをワークシート内の Range の Value に割り当てます。

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

## 行の高さと列の幅の設定

行の高さと列の幅をカスタマイズして、レイアウトやデータ表示を見栄えよくします。ワークシートのデフォルトの行の高さと列の幅を設定するには、"StandardHeight" と "StandardWidth" を使用します。

```csharp
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


## テーブルの作成

「Income」と「Expenses」という 2 つのテーブルを追加し、それぞれに組み込みテーブルスタイルを適用します。

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

## 式の設定

当月の収入と支出を集計する 2 つのカスタム名を作成し、月間収入合計、月間支出合計、収入に占める支出の割合、差額を計算する式を追加します。

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


## スタイルの設定

範囲スタイルを変更する方法には、次の 2 つがあります。
- 組み込みスタイルまたはカスタムスタイルを名前で適用する
- 要素ごとに個別スタイルを設定する

ここでは、"Currency"、"Heading 1"、および "Percent" 組み込みスタイルを変更し、これらをいくつかのセル範囲に適用します。さらに、他の範囲のスタイル要素を個別に変更します。

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


## 条件付き書式設定の追加

DioDocs for Excel では、さまざまな条件付き書式設定ルールがサポートされています。ここでは、収入に占める支出の割合を表示するグラデーション付きデータバールールを作成します。このルールは、値を表示しないデータバーを表示します。

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

## チャートの追加 

収入と支出の差を示す縦棒グラフを作成します。レイアウトを見栄えよくするために、系列の重なりとギャップ幅を変更し、さらに一部のチャート要素（チャート領域、軸線、目盛りのラベル、データポイント）の書式設定をカスタマイズします。

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

## Excel への保存

「SimpleBudget.xlsx」という名前で Excel ファイルに保存します。

- C#
```csharp
workbook.Save("SimpleBudget.xlsx");
```
- VB
```vbnet
workbook.Save("SimpleBudget.xlsx")
```

保存された [SimpleBudget.xlsx](api/examples/xlsx/GrapeCity.Documents.Excel.Examples.Tutorial?fileName=SimpleBudget) をダウンロードして表示できます。[チュートリアルのソースプロジェクト](GrapeCity.Documents.Excel.Tutorial.zip) をダウンロードしてご自身でコードを実行する場合は、事前に [.NET Core](https://www.microsoft.com/net/core) をお使いのマシンにインストールしてください。
