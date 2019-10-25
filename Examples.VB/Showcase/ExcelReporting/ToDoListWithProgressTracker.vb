Imports System.Drawing

Namespace Showcase
    Public Class ToDoListWithProgressTracker
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            Dim data(,) As Object = {
                {"TASK", "PRIORITY", "STATUS", "START DATE", "DUE DATE", "% COMPLETE", "DONE?", "NOTES"},
                {"First Thing I Need To Do", "Normal", "Not Started", Nothing, Nothing, 0, Nothing, Nothing},
                {"Other Thing I Need To Finish", "High", "In Progress", Nothing, Nothing, 0.5, Nothing, Nothing},
                {"Something Else To Get Done", "Low", "Complete", Nothing, Nothing, 1, Nothing, Nothing},
                {"More Errands And Things", "Normal", "In Progress", Nothing, Nothing, 0.75, Nothing, Nothing},
                {"So Much To Get Done This Week", "High", "In Progress", Nothing, Nothing, 0.25, Nothing, Nothing}
            }

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Name = "To-Do List"
            worksheet.TabColor = Color.FromArgb(148, 112, 135)
            worksheet.SheetView.DisplayGridlines = False

            'Set Value.
            worksheet.Range!B1.Value = "To-Do List"
            worksheet.Range("B2:I7").Value = data

            'Set formula.
            worksheet.Range!E3.Formula = "=TODAY()"
            worksheet.Range!E4.Formula = "=TODAY()-30"
            worksheet.Range!E5.Formula = "=TODAY()-23"
            worksheet.Range!E6.Formula = "=TODAY()-15"
            worksheet.Range!E7.Formula = "=TODAY()-5"

            'Change the range's RowHeight and ColumnWidth.
            worksheet.StandardHeight = 30
            worksheet.StandardWidth = 8.88671875
            worksheet.Range("1:1").RowHeight = 72.75
            worksheet.Range("2:2").RowHeight = 33
            worksheet.Range("3:7").RowHeight = 30.25
            worksheet.Range("A:A").ColumnWidth = 2.77734375
            worksheet.Range("B:B").ColumnWidth = 29.109375
            worksheet.Range("C:G").ColumnWidth = 16.77734375
            worksheet.Range("H:H").ColumnWidth = 10.77734375
            worksheet.Range("I:I").ColumnWidth = 29.6640625
            worksheet.Range("J:J").ColumnWidth = 2.77734375

            'Modify the build in name styles.
            Dim nameStyle_Normal = workbook.Styles("Normal")
            nameStyle_Normal.VerticalAlignment = VerticalAlignment.Center
            nameStyle_Normal.WrapText = True
            nameStyle_Normal.Font.ThemeFont = ThemeFont.Minor
            nameStyle_Normal.Font.ThemeColor = ThemeColor.Dark1
            nameStyle_Normal.Font.TintAndShade = 0.25

            Dim nameStyle_Title = workbook.Styles("Title")
            nameStyle_Title.HorizontalAlignment = HorizontalAlignment.General
            nameStyle_Title.VerticalAlignment = VerticalAlignment.Bottom
            nameStyle_Title.Font.ThemeFont = ThemeFont.Minor
            nameStyle_Title.Font.Bold = True
            nameStyle_Title.Font.Size = 38
            nameStyle_Title.Font.ThemeColor = ThemeColor.Dark1
            nameStyle_Title.Font.TintAndShade = 0.249946592608417
            nameStyle_Title.Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Thick
            nameStyle_Title.Borders(BordersIndex.EdgeBottom).ThemeColor = ThemeColor.Dark1
            nameStyle_Title.IncludeAlignment = True
            nameStyle_Title.IncludeBorder = True

            Dim nameStyle_Percent = workbook.Styles("Percent")
            nameStyle_Percent.HorizontalAlignment = HorizontalAlignment.Right
            nameStyle_Percent.IndentLevel = 1
            nameStyle_Percent.VerticalAlignment = VerticalAlignment.Center
            nameStyle_Percent.IncludeAlignment = True

            Dim nameStyle_Heading_1 = workbook.Styles("Heading 1")
            nameStyle_Heading_1.HorizontalAlignment = HorizontalAlignment.Left
            nameStyle_Heading_1.VerticalAlignment = VerticalAlignment.Bottom
            nameStyle_Heading_1.Font.ThemeFont = ThemeFont.Major
            nameStyle_Heading_1.Font.Bold = False
            nameStyle_Heading_1.Font.Size = 11
            nameStyle_Heading_1.Font.ThemeColor = ThemeColor.Dark1
            nameStyle_Heading_1.Font.TintAndShade = 0.249946592608417
            nameStyle_Heading_1.Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.None
            nameStyle_Heading_1.IncludeNumber = True
            nameStyle_Heading_1.IncludeAlignment = True
            nameStyle_Heading_1.IncludeBorder = False

            Dim nameStyle_Heading_2 = workbook.Styles("Heading 2")
            nameStyle_Heading_2.HorizontalAlignment = HorizontalAlignment.Right
            nameStyle_Heading_2.IndentLevel = 2
            nameStyle_Heading_2.VerticalAlignment = VerticalAlignment.Bottom
            nameStyle_Heading_2.Font.ThemeFont = ThemeFont.Major
            nameStyle_Heading_2.Font.Bold = False
            nameStyle_Heading_2.Font.Size = 11
            nameStyle_Heading_2.Font.ThemeColor = ThemeColor.Dark2
            nameStyle_Heading_2.Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.None
            nameStyle_Heading_2.IncludeNumber = True
            nameStyle_Heading_2.IncludeAlignment = True

            'Create custom name styes.
            Dim nameStyle_Done As IStyle = workbook.Styles.Add("Done")
            nameStyle_Done.NumberFormat = """Done"";"""";"""""
            nameStyle_Done.HorizontalAlignment = HorizontalAlignment.Center
            nameStyle_Done.VerticalAlignment = VerticalAlignment.Center
            nameStyle_Done.Font.ThemeFont = ThemeFont.Minor
            nameStyle_Done.Font.ThemeColor = ThemeColor.Light1

            Dim nameStyle_Date As IStyle = workbook.Styles.Add("Date")
            nameStyle_Date.NumberFormat = "yyyy/m/d"
            nameStyle_Date.HorizontalAlignment = HorizontalAlignment.Right
            nameStyle_Date.VerticalAlignment = VerticalAlignment.Center
            nameStyle_Date.Font.ThemeFont = ThemeFont.Minor
            nameStyle_Date.Font.ThemeColor = ThemeColor.Dark1
            nameStyle_Date.Font.TintAndShade = 0.249946592608417
            nameStyle_Date.IncludeBorder = False
            nameStyle_Date.IncludePatterns = False

            'Apply the above name styles on ranges.
            worksheet.Range("B1:I1").Style = workbook.Styles("Title")
            worksheet.Range("B2:D2").Style = workbook.Styles("Heading 1")
            worksheet.Range("E2:F2").Style = workbook.Styles("Heading 2")
            worksheet.Range!G2.Style = workbook.Styles("Heading 1")
            worksheet.Range("H2:H7").Style = workbook.Styles("Done")
            worksheet.Range!I2.Style = workbook.Styles("Heading 1")
            worksheet.Range("E3:F7").Style = workbook.Styles("Date")
            worksheet.Range("G3:G7").Style = workbook.Styles("Percent")

            'Add one custom table style.
            Dim style As ITableStyle = workbook.TableStyles.Add("To-do List")
            style.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Thin
            style.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeBottom).ThemeColor = ThemeColor.Light1
            style.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeBottom).TintAndShade = -0.14993743705557422
            style.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.InsideHorizontal).LineStyle = BorderLineStyle.Thin
            style.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.InsideHorizontal).ThemeColor = ThemeColor.Light1
            style.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.InsideHorizontal).TintAndShade = -0.14993743705557422

            'Create a table and apply the above table style.
            Dim table As ITable = worksheet.Tables.Add(worksheet.Range("B2:I7"), True)
            table.Name = "ToDoList"
            table.TableStyle = style

            'Use table formula in table range.
            worksheet.Range!F3.Formula = "=[@[START DATE]]+7"
            worksheet.Range!F4.Formula = "=[@[START DATE]]+35"
            worksheet.Range!F5.Formula = "=[@[START DATE]]+10"
            worksheet.Range!F6.Formula = "=[@[START DATE]]+36"
            worksheet.Range!F7.Formula = "=[@[START DATE]]+14"
            worksheet.Range("H3:H7").Formula = "=--([@[% COMPLETE]]>=1)"

            'Add a expression rule.
            Dim expression As IFormatCondition = TryCast(worksheet.Range("B3:I7").FormatConditions.Add(FormatConditionType.Expression, FormatConditionOperator.Between, "=AND($G3=0,$G3<>"""")", Nothing), IFormatCondition)
            expression.Interior.ThemeColor = ThemeColor.Light1
            expression.Interior.TintAndShade = -0.0499893185216834

            'Add a data bar rule.
            Dim dataBar As IDataBar = worksheet.Range("G3:G7").FormatConditions.AddDatabar()
            dataBar.BarFillType = DataBarFillType.Solid
            dataBar.BarColor.ThemeColor = ThemeColor.Accent1
            dataBar.BarColor.TintAndShade = 0.39997558519241921

            'Add an icon set rule.
            Dim iconSet As IIconSetCondition = worksheet.Range("H3:H7").FormatConditions.AddIconSetCondition()
            iconSet.IconSet = workbook.IconSets(IconSetType.Icon3Symbols)
            iconSet.IconCriteria(2).Operator = FormatConditionOperator.GreaterEqual
            iconSet.IconCriteria(2).Value = 1
            iconSet.IconCriteria(2).Type = ConditionValueTypes.Number
            iconSet.IconCriteria(1).Operator = FormatConditionOperator.GreaterEqual
            iconSet.IconCriteria(1).Value = 0
            iconSet.IconCriteria(1).Type = ConditionValueTypes.Number

            'Add a cell value rule.
            Dim cellValue As IFormatCondition = TryCast(worksheet.Range("H3:H7").FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.NotEqual, 1), IFormatCondition)
            cellValue.StopIfTrue = True

            'Create list validations.
            worksheet.Range("C3:C7").Validation.Add(ValidationType.List, ValidationAlertStyle.Warning, ValidationOperator.Between, "Low, Normal, High", Nothing)
            Dim validation As IValidation = worksheet.Range("C3:C7").Validation
            validation.ErrorMessage = "Select entry from the list. Select CANCEL, then press ALT+DOWN ARROW to navigate the list. Select ENTER to make selection"

            worksheet.Range("D3:D7").Validation.Add(ValidationType.List, ValidationAlertStyle.Warning, ValidationOperator.Between, "Not Started,In Progress, Deferred, Complete", Nothing)
            validation = worksheet.Range("D3:D7").Validation
            validation.ErrorMessage = "Select entry from the list. Select CANCEL, then press ALT+DOWN ARROW to navigate the list. Select ENTER to make selection"

            worksheet.Range("G3:G7").Validation.Add(ValidationType.List, ValidationAlertStyle.Warning, ValidationOperator.Between, "0%,25%,50%,75%,100%", Nothing)
            validation = worksheet.Range("G3:G7").Validation
            validation.ErrorMessage = "Select entry from the list. Select CANCEL, then press ALT+DOWN ARROW to navigate the list. Select ENTER to make selection"

            'Create custom validation.
            worksheet.Range("F3:F7").Validation.Add(ValidationType.Custom, ValidationAlertStyle.Warning, ValidationOperator.Between, "=F3>=E3", Nothing)
            validation = worksheet.Range("F3:F7").Validation
            validation.ErrorMessage = "The Due Date must be greater than or equal to the Start Date. Select YES to keep the value, NO to retry or CANCEL to clear the entry"

            'Create none validations, set inputmessage.
            worksheet.Range!B2.Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, Nothing, Nothing)
            validation = worksheet.Range!B2.Validation
            validation.InputMessage = "Enter Task in this column under this heading. Use heading filters to find specific entries"

            worksheet.Range!C2.Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, Nothing, Nothing)
            validation = worksheet.Range!C2.Validation
            validation.InputMessage = "Select Priority in this column under this heading. Press ALT+DOWN ARROW to open the drop-down list, then ENTER to make selection"

            worksheet.Range!D2.Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, Nothing, Nothing)
            validation = worksheet.Range!D2.Validation
            validation.InputMessage = "Select Status in this column under this heading.  Press ALT+DOWN ARROW to open the drop-down list, then ENTER to make selection"

            worksheet.Range!E2.Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, Nothing, Nothing)
            validation = worksheet.Range!E2.Validation
            validation.InputMessage = "Enter Start Date in this column under this heading"

            worksheet.Range!F2.Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, Nothing, Nothing)
            validation = worksheet.Range!F2.Validation
            validation.InputMessage = "Enter Due Date in this column under this heading"

            worksheet.Range!G2.Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, Nothing, Nothing)
            validation = worksheet.Range!G2.Validation
            validation.InputMessage = "Select % Complete in this column. Press ALT+DOWN ARROW to open the drop-down list, then ENTER to make selection. A status bar indicates progress toward completion"

            worksheet.Range!H2.Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, Nothing, Nothing)
            validation = worksheet.Range!H2.Validation
            validation.InputMessage = "Icon indicator for task completion in this column under this heading is automatically updated as tasks complete"

            worksheet.Range!I2.Validation.Add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, Nothing, Nothing)
            validation = worksheet.Range!I2.Validation
            validation.InputMessage = "Enter Notes in this column under this heading"

            'Create customize theme.
            Dim themes As New Themes
            Dim theme As ITheme = themes.Add("test")
            theme.ThemeColorScheme(ThemeColor.Dark1).RGB = Color.FromArgb(0, 0, 0)
            theme.ThemeColorScheme(ThemeColor.Light1).RGB = Color.FromArgb(255, 255, 255)
            theme.ThemeColorScheme(ThemeColor.Dark2).RGB = Color.FromArgb(37, 28, 34)
            theme.ThemeColorScheme(ThemeColor.Light2).RGB = Color.FromArgb(240, 248, 246)
            theme.ThemeColorScheme(ThemeColor.Accent1).RGB = Color.FromArgb(148, 112, 135)
            theme.ThemeColorScheme(ThemeColor.Accent2).RGB = Color.FromArgb(71, 166, 181)
            theme.ThemeColorScheme(ThemeColor.Accent3).RGB = Color.FromArgb(234, 194, 53)
            theme.ThemeColorScheme(ThemeColor.Accent4).RGB = Color.FromArgb(107, 192, 129)
            theme.ThemeColorScheme(ThemeColor.Accent5).RGB = Color.FromArgb(233, 115, 61)
            theme.ThemeColorScheme(ThemeColor.Accent6).RGB = Color.FromArgb(251, 147, 59)
            theme.ThemeColorScheme(ThemeColor.Hyperlink).RGB = Color.FromArgb(71, 166, 181)
            theme.ThemeColorScheme(ThemeColor.FollowedHyperlink).RGB = Color.FromArgb(148, 112, 135)
            theme.ThemeFontScheme.Major(FontLanguageIndex.Latin).Name = "Helvetica Neue"
            theme.ThemeFontScheme.Minor(FontLanguageIndex.Latin).Name = "Bookman Old Style"

            'Apply the above custom theme.
            workbook.Theme = theme

            'Set active cell.
            worksheet.Range!G4.Activate()
        End Sub

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
