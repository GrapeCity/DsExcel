Imports System.Drawing

Namespace Showcase
    Public Class PersonalAddressBook
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



            '***************************Set RowHeight & Width****************************
            worksheet.StandardHeight = 30
            worksheet.Range("3:4").RowHeight = 30.25
            worksheet.Range("1:1").RowHeight = 103.5
            worksheet.Range("2:2").RowHeight = 38.25
            worksheet.Range("A:A").ColumnWidth = 2.625
            worksheet.Range("B:B").ColumnWidth = 22.25
            worksheet.Range("C:E").ColumnWidth = 17.25
            worksheet.Range("F:F").ColumnWidth = 31.875
            worksheet.Range("G:G").ColumnWidth = 22.625
            worksheet.Range("H:H").ColumnWidth = 30
            worksheet.Range("I:I").ColumnWidth = 20.25
            worksheet.Range("J:J").ColumnWidth = 17.625
            worksheet.Range("K:K").ColumnWidth = 12.625
            worksheet.Range("L:L").ColumnWidth = 37.25
            worksheet.Range("M:M").ColumnWidth = 2.625



            '*******************************Set Table Value & Formulas*************************************
            Dim table As ITable = worksheet.Tables.Add(worksheet.Range("B2:L4"), True)
            worksheet.Range("B2:L4").Value = New Object(,) {
                {"NAME", "WORK", "CELL", "HOME", "EMAIL", "BIRTHDAY", "ADDRESS", "CITY", "STATE", "ZIP", "NOTE"},
                {"Kim Abercrombie", 1235550123, 1235550123, 1235550123, "someone@example.com", Nothing, "123 N. Maple", "Cherryville", "WA", 98031, ""},
                {"John Smith", 3215550123, "", "", "someone@example.com", Nothing, "456 E. Aspen", "", "", "", ""}
            }
            worksheet.Range!B1.Value = "ADDRESS BOOK"
            worksheet.Range!G3.Formula = "=TODAY()"
            worksheet.Range!G4.Formula = "=TODAY()+5"



            '****************************Set Table Style********************************
            Dim tableStyle As ITableStyle = workbook.TableStyles.Add("Personal Address Book")
            workbook.DefaultTableStyle = "Personal Address Book"

            'Set WholeTable element style.
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders.Color = Color.FromArgb(179, 35, 23)
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeLeft).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeRight).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.InsideVertical).LineStyle = BorderLineStyle.Thin
            tableStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders(BordersIndex.InsideHorizontal).LineStyle = BorderLineStyle.Thin

            'Set FirstColumn element style.
            tableStyle.TableStyleElements(TableStyleElementType.FirstColumn).Font.Bold = True

            'Set SecondColumns element style.
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders.Color = Color.FromArgb(179, 35, 23)
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders(BordersIndex.EdgeTop).LineStyle = BorderLineStyle.Thick
            tableStyle.TableStyleElements(TableStyleElementType.HeaderRow).Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Thick



            '***********************************Set Named Styles*****************************
            Dim normalStyle As IStyle = workbook.Styles("Normal")
            normalStyle.Font.Name = "Arial"
            normalStyle.Font.Color = Color.FromArgb(179, 35, 23)
            normalStyle.HorizontalAlignment = HorizontalAlignment.Left
            normalStyle.IndentLevel = 1
            normalStyle.VerticalAlignment = VerticalAlignment.Center
            normalStyle.WrapText = True

            Dim titleStyle As IStyle = workbook.Styles("Title")
            titleStyle.IncludeAlignment = True
            titleStyle.HorizontalAlignment = HorizontalAlignment.Left
            titleStyle.VerticalAlignment = VerticalAlignment.Center
            titleStyle.Font.Name = "Arial"
            titleStyle.Font.Bold = True
            titleStyle.Font.Size = 72
            titleStyle.Font.Color = Color.FromArgb(179, 35, 23)

            Dim heading1Style As IStyle = workbook.Styles("Heading 1")
            heading1Style.IncludeBorder = False
            heading1Style.Font.Name = "Arial"
            heading1Style.Font.Size = 18
            heading1Style.Font.Color = Color.FromArgb(179, 35, 23)

            Dim dataStyle As IStyle = workbook.Styles.Add("Data")
            dataStyle.IncludeNumber = True
            dataStyle.NumberFormat = "m/d/yyyy"

            Dim phoneStyle As IStyle = workbook.Styles.Add("Phone")
            phoneStyle.IncludeNumber = True
            phoneStyle.NumberFormat = "[<=9999999]###-####;(###) ###-####"



            '****************************************Use NamedStyle**************************
            worksheet.SheetView.DisplayGridlines = False
            worksheet.Range("B2:L2").Interior.Color = Color.FromArgb(217, 217, 217)
            worksheet.Range("B3:B4").Font.Bold = True
            worksheet.Range("2:2").HorizontalAlignment = HorizontalAlignment.Left

            table.TableStyle = tableStyle
            worksheet.Range!B1.Style = titleStyle
            worksheet.Range("B2:L2").Style = heading1Style
            worksheet.Range("C3:E4").Style = phoneStyle
            worksheet.Range("G3:G4").Style = dataStyle

        End Sub

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
