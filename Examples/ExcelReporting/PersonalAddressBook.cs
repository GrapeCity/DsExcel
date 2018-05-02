using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.ExcelReporting
{
    public class PersonalAddressBook : ExampleBase
    {
        protected override void BeforeExecute(Workbook workbook, string[] userAgents)
        {
            if (AgentIsMac(userAgents))
            {
                Themes themes = new Themes();
                ITheme theme = themes.Add("testTheme", Themes.OfficeTheme);
                theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Trebuchet MS";
                workbook.Theme = theme;
                var style_Normal = workbook.Styles["Normal"];
                style_Normal.Font.ThemeFont = ThemeFont.Minor;
            }
        }
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {

            IWorksheet worksheet = workbook.Worksheets[0];


            
            //***************************Set RowHeight & Width****************************
                worksheet.StandardHeight = 30;
                worksheet.Range["3:4"].RowHeight = 30.25;
                worksheet.Range["1:1"].RowHeight = 103.50;
                worksheet.Range["2:2"].RowHeight = 38.25;
                worksheet.Range["A:A"].ColumnWidth = 2.625;
                worksheet.Range["B:B"].ColumnWidth = 22.25;
                worksheet.Range["C:E"].ColumnWidth = 17.25;
                worksheet.Range["F:F"].ColumnWidth = 31.875;
                worksheet.Range["G:G"].ColumnWidth = 22.625;
                worksheet.Range["H:H"].ColumnWidth = 30;
                worksheet.Range["I:I"].ColumnWidth = 20.25;
                worksheet.Range["J:J"].ColumnWidth = 17.625;
                worksheet.Range["K:K"].ColumnWidth = 12.625;
                worksheet.Range["L:L"].ColumnWidth = 37.25;
                worksheet.Range["M:M"].ColumnWidth = 2.625;

            

            //*******************************Set Table Value & Formulas*************************************
            ITable table = worksheet.Tables.Add(worksheet.Range["B2:L4"], true);
            worksheet.Range["B2:L4"].Value = new object[,]
            {
                { "NAME", "WORK", "CELL", "HOME", "EMAIL", "BIRTHDAY", "ADDRESS", "CITY", "STATE", "ZIP", "NOTE" },
                { "Kim Abercrombie", 1235550123, 1235550123, 1235550123, "someone@example.com", null, "123 N. Maple", "Cherryville", "WA", 98031, "" },
                { "John Smith", 3215550123, "", "", "someone@example.com", null, "456 E. Aspen", "", "", "", "" },
            };
            worksheet.Range["B1"].Value = "ADDRESS BOOK";
            worksheet.Range["G3"].Formula = "=TODAY()";
            worksheet.Range["G4"].Formula = "=TODAY()+5";

            

            //****************************Set Table Style********************************
            ITableStyle tableStyle = workbook.TableStyles.Add("Personal Address Book");
            workbook.DefaultTableStyle = "Personal Address Book";

            //Set WholeTable element style.
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders.Color = Color.FromRGB(179, 35, 23);
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.Thin;
            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.Thin;

            //Set FirstColumn element style.
            tableStyle.TableStyleElements[TableStyleElementType.FirstColumn].Font.Bold = true;

            //Set SecondColumns element style.
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders.Color = Color.FromRGB(179, 35, 23);
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thick;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;
            


            //***********************************Set Named Styles*****************************
            IStyle normalStyle = workbook.Styles["Normal"];
            normalStyle.Font.Name = "Arial";
            normalStyle.Font.Color = Color.FromRGB(179, 35, 23);
            normalStyle.HorizontalAlignment = HorizontalAlignment.Left;
            normalStyle.IndentLevel = 1;
            normalStyle.VerticalAlignment = VerticalAlignment.Center;
            normalStyle.WrapText = true;

            IStyle titleStyle = workbook.Styles["Title"];
            titleStyle.IncludeAlignment = true;
            titleStyle.HorizontalAlignment = HorizontalAlignment.Left;
            titleStyle.VerticalAlignment = VerticalAlignment.Center;
            titleStyle.Font.Name = "Arial";
            titleStyle.Font.Bold = true;
            titleStyle.Font.Size = 72;
            titleStyle.Font.Color = Color.FromRGB(179, 35, 23);

            IStyle heading1Style = workbook.Styles["Heading 1"];
            heading1Style.IncludeBorder = false;
            heading1Style.Font.Name = "Arial";
            heading1Style.Font.Size = 18;
            heading1Style.Font.Color = Color.FromRGB(179, 35, 23);

            IStyle dataStyle = workbook.Styles.Add("Data");
            dataStyle.IncludeNumber = true;
            dataStyle.NumberFormat = "m/d/yyyy";

            IStyle phoneStyle = workbook.Styles.Add("Phone");
            phoneStyle.IncludeNumber = true;
            phoneStyle.NumberFormat = "[<=9999999]###-####;(###) ###-####";



            //****************************************Use NamedStyle**************************
            worksheet.SheetView.DisplayGridlines = false;
            worksheet.Range["B2:L2"].Interior.Color = Color.FromRGB(217, 217, 217);
            worksheet.Range["B3:B4"].Font.Bold = true;
            worksheet.Range["2:2"].HorizontalAlignment = HorizontalAlignment.Left;

            table.TableStyle = tableStyle;
            worksheet.Range["B1"].Style = titleStyle;
            worksheet.Range["B2:L2"].Style = heading1Style;
            worksheet.Range["C3:E4"].Style = phoneStyle;
            worksheet.Range["G3:G4"].Style = dataStyle;

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
