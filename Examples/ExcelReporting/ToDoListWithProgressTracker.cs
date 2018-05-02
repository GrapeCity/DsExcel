using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelReporting
{
    public class ToDoListWithProgressTracker : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
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
            worksheet.Range["3:7"].RowHeight = 30.25;
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
            style.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            style.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].ThemeColor = ThemeColor.Light1;
            style.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].TintAndShade = -0.14993743705557422;
            style.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;
            style.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].ThemeColor = ThemeColor.Light1;
            style.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].TintAndShade = -0.14993743705557422;

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
            iconSet.IconCriteria[2].Operator = FormatConditionOperator.GreaterEqual;
            iconSet.IconCriteria[2].Value = 1;
            iconSet.IconCriteria[2].Type = ConditionValueTypes.Number;
            iconSet.IconCriteria[1].Operator = FormatConditionOperator.GreaterEqual;
            iconSet.IconCriteria[1].Value = 0;
            iconSet.IconCriteria[1].Type = ConditionValueTypes.Number;

            //Add a cell value rule.
            IFormatCondition cellValue = worksheet.Range["H3:H7"].FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.NotEqual, 1) as IFormatCondition;
            cellValue.StopIfTrue = true;

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
            theme.ThemeFontScheme.Major[FontLanguageIndex.Latin].Name = "Helvetica Neue";
            theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Bookman Old Style";

            //Apply the above custom theme.
            workbook.Theme = theme;

            //Set active cell.
            worksheet.Range["G4"].Activate();
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
