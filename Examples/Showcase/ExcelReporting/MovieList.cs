using GrapeCity.Documents.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Showcase
{
    public class MovieList : ExampleBase
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

            //------------------Set RowHeight & ColumnWidth----------------
            worksheet.StandardHeight = 43.5;
            worksheet.StandardWidth = 8.43;

            worksheet.Range["1:1"].RowHeight = 171;
            worksheet.Range["2:2"].RowHeight = 12.75;
            worksheet.Range["3:3"].RowHeight = 22.5;
            worksheet.Range["4:7"].RowHeight = 43.75;
            worksheet.Range["A:A"].ColumnWidth = 2.887;
            worksheet.Range["B:B"].ColumnWidth = 8.441;
            worksheet.Range["C:C"].ColumnWidth = 12.777;
            worksheet.Range["D:D"].ColumnWidth = 25.109;
            worksheet.Range["E:E"].ColumnWidth = 12.109;
            worksheet.Range["F:F"].ColumnWidth = 41.664;
            worksheet.Range["G:G"].ColumnWidth = 18.555;
            worksheet.Range["H:H"].ColumnWidth = 11;
            worksheet.Range["I:I"].ColumnWidth = 13.664;
            worksheet.Range["J:J"].ColumnWidth = 15.109;
            worksheet.Range["K:K"].ColumnWidth = 38.887;
            worksheet.Range["L:L"].ColumnWidth = 2.887;


            //------------------------Set Table Values-------------------
            ITable table = worksheet.Tables.Add(worksheet.Range["B3:K7"], true);
            worksheet.Range["B3:K7"].Value = new object[,]
            {
                { "NO.", "YEAR", "TITLE", "REVIEW", "STARRING ACTORS", "DIRECTOR", "GENRE", "RATING", "FORMAT", "COMMENTS" },
                { 1, 1994, "Forrest Gump", "5 Stars", "Tom Hanks, Robin Wright, Gary Sinise", "Robert Zemeckis", "Drama", "PG-13", "DVD", "Based on the 1986 novel of the same name by Winston Groom" },
                { 2, 1946, "It’s a Wonderful Life", "2 Stars", "James Stewart, Donna Reed, Lionel Barrymore ", "Frank Capra", "Drama", "G", "VHS", "Colorized version" },
                { 3, 1988, "Big", "4 Stars", "Tom Hanks, Elizabeth Perkins, Robert Loggia ", "Penny Marshall", "Comedy", "PG", "DVD", "" },
                { 4, 1954, "Rear Window", "3 Stars", "James Stewart, Grace Kelly, Wendell Corey ", "Alfred Hitchcock", "Suspense", "PG", "Blu-ray", "" },
            };


            //-----------------------Set Table style--------------------
            ITableStyle tableStyle = workbook.TableStyles.Add("Movie List");
            workbook.DefaultTableStyle = "Movie List";

            tableStyle.TableStyleElements[TableStyleElementType.WholeTable].Interior.Color = Color.White;

            tableStyle.TableStyleElements[TableStyleElementType.FirstRowStripe].Interior.Color = Color.FromArgb(38, 38, 38);

            tableStyle.TableStyleElements[TableStyleElementType.SecondRowStripe].Interior.Color = Color.Black;

            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Font.Color = Color.Black;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders.Color = Color.FromArgb(38, 38, 38);
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Interior.Color = Color.FromArgb(68, 217, 255);
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeTop].LineStyle = BorderLineStyle.Thick;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeLeft].LineStyle = BorderLineStyle.None;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeRight].LineStyle = BorderLineStyle.None;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.None;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.InsideHorizontal].LineStyle = BorderLineStyle.None;
            tableStyle.TableStyleElements[TableStyleElementType.HeaderRow].Borders[BordersIndex.InsideVertical].LineStyle = BorderLineStyle.None;


            //--------------------------------Set Named Styles---------------------
            IStyle movieListBorderStyle = workbook.Styles.Add("Movie list border");
            movieListBorderStyle.IncludeNumber = true;
            movieListBorderStyle.IncludeAlignment = true;
            movieListBorderStyle.VerticalAlignment = VerticalAlignment.Center;
            movieListBorderStyle.WrapText = true;
            movieListBorderStyle.IncludeFont = true;
            movieListBorderStyle.Font.Name = "Helvetica";
            movieListBorderStyle.Font.Size = 11;
            movieListBorderStyle.Font.Color = Color.White;
            movieListBorderStyle.IncludeBorder = true;
            movieListBorderStyle.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;
            movieListBorderStyle.Borders[BordersIndex.EdgeBottom].Color = Color.FromArgb(38, 38, 38);
            movieListBorderStyle.IncludePatterns = true;
            movieListBorderStyle.Interior.Color = Color.FromArgb(238, 219, 78);

            IStyle nOStyle = workbook.Styles.Add("NO.");
            nOStyle.IncludeNumber = true;
            nOStyle.IncludeAlignment = true;
            nOStyle.HorizontalAlignment = HorizontalAlignment.Left;
            nOStyle.VerticalAlignment = VerticalAlignment.Center;
            nOStyle.IncludeFont = true;
            nOStyle.Font.Name = "Helvetica";
            nOStyle.Font.Size = 11;
            nOStyle.Font.Color = Color.White;
            nOStyle.IncludeBorder = true;
            nOStyle.IncludePatterns = true;
            nOStyle.Interior.Color = Color.FromArgb(38, 38, 38);

            IStyle reviewStyle = workbook.Styles.Add("Review");
            reviewStyle.IncludeNumber = true;
            reviewStyle.IncludeAlignment = true;
            reviewStyle.VerticalAlignment = VerticalAlignment.Center;
            reviewStyle.IncludeFont = true;
            reviewStyle.Font.Name = "Helvetica";
            reviewStyle.Font.Size = 11;
            reviewStyle.Font.Color = Color.White;
            reviewStyle.IncludeBorder = true;
            reviewStyle.IncludePatterns = true;
            reviewStyle.Interior.Color = Color.FromArgb(38, 38, 38);

            IStyle yearStyle = workbook.Styles.Add("Year");
            yearStyle.IncludeNumber = true;
            yearStyle.IncludeAlignment = true;
            yearStyle.HorizontalAlignment = HorizontalAlignment.Left;
            yearStyle.VerticalAlignment = VerticalAlignment.Center;
            yearStyle.IncludeFont = true;
            yearStyle.Font.Name = "Helvetica";
            yearStyle.Font.Size = 11;
            yearStyle.Font.Color = Color.White;
            yearStyle.IncludeBorder = true;
            yearStyle.IncludePatterns = true;
            yearStyle.Interior.Color = Color.FromArgb(38, 38, 38);

            IStyle heading1Style = workbook.Styles["Heading 1"];
            heading1Style.IncludeAlignment = true;
            heading1Style.VerticalAlignment = VerticalAlignment.Bottom;
            heading1Style.IncludeBorder = true;
            heading1Style.Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Thick;
            heading1Style.Borders[BordersIndex.EdgeBottom].Color = Color.FromArgb(68, 217, 255);
            heading1Style.IncludeFont = true;
            heading1Style.Font.Name = "Helvetica";
            heading1Style.Font.Bold = false;
            heading1Style.Font.Size = 12;
            heading1Style.Font.Color = Color.Black;

            IStyle normalStyle = workbook.Styles["Normal"];
            normalStyle.IncludeNumber = true;
            normalStyle.IncludeAlignment = true;
            normalStyle.VerticalAlignment = VerticalAlignment.Center;
            normalStyle.WrapText = true;
            normalStyle.IncludeFont = true;
            normalStyle.Font.Name = "Helvetica";
            normalStyle.Font.Size = 11;
            normalStyle.Font.Color = Color.White;
            normalStyle.IncludePatterns = true;
            normalStyle.Interior.Color = Color.FromArgb(38, 38, 38);


            //-----------------------------Use NamedStyle--------------------------
            worksheet.SheetView.DisplayGridlines = false;
            worksheet.TabColor = Color.FromArgb(38, 38, 38);
            table.TableStyle = tableStyle;

            worksheet.Range["A2:L2"].Style = movieListBorderStyle;
            worksheet.Range["B3:K3"].Style = heading1Style;
            worksheet.Range["B4:B7"].Style = nOStyle;
            worksheet.Range["C4:C7"].Style = yearStyle;
            worksheet.Range["E4:E7"].Style = reviewStyle;
            worksheet.Range["F4:F7"].IndentLevel = 1;
            worksheet.Range["F4:F7"].HorizontalAlignment = HorizontalAlignment.Left;


            //-----------------------------Add Shapes------------------------------
            //Movie picture
            System.IO.Stream stream = this.GetResourceStream("movie.png");
            IShape pictureShape = worksheet.Shapes.AddPicture(stream, ImageType.PNG, 0, 1, worksheet.Range["A:L"].Width, worksheet.Range["1:1"].Height - 1.5);
            pictureShape.Placement = Placement.Move;

            //Movie list picture
            System.IO.Stream stream2 = this.GetResourceStream("list.png");
            IShape pictureShape2 = worksheet.Shapes.AddPicture(stream2, ImageType.PNG, 1, 0.8, 325.572, 85.51);
            pictureShape2.Placement = Placement.Move;

            //Rounded Rectangular Callout 7
            IShape roundedRectangular = worksheet.Shapes.AddShape(AutoShapeType.RoundedRectangularCallout, 437.5, 22.75, 342, 143);
            roundedRectangular.Name = "Rounded Rectangular Callout 7";
            roundedRectangular.Placement = Placement.Move;
            roundedRectangular.TextFrame.TextRange.Font.Name = "Helvetica";
            roundedRectangular.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(38, 38, 38);

            roundedRectangular.Fill.Solid();
            roundedRectangular.Fill.Color.RGB = Color.FromArgb(68, 217, 255);
            roundedRectangular.Fill.Transparency = 0;
            roundedRectangular.Line.Solid();
            roundedRectangular.Line.Color.RGB = Color.FromArgb(0, 129, 162);
            roundedRectangular.Line.Weight = 2;
            roundedRectangular.Line.Transparency = 0;

            ITextRange roundedRectangular_p0 = roundedRectangular.TextFrame.TextRange.Paragraphs[0];
            roundedRectangular_p0.Runs.Font.Bold = true;
            roundedRectangular_p0.Runs.Add("TABLE");
            roundedRectangular_p0.Runs.Add(" TIP");

            roundedRectangular.TextFrame.TextRange.Paragraphs.Add("");

            ITextRange roundedRectangular_p2 = roundedRectangular.TextFrame.TextRange.Paragraphs.Add();
            roundedRectangular_p2.Runs.Add("Use the drop down arrows in the table headings to quickly filter your movie list. " +
                "For multiple entry fields, such as Starring Actors,  select the drop down arrow next to the field and enter text in the Search box. " +
                "For example, type Tom Hanks or James Stewart, and then select OK.");

            roundedRectangular.TextFrame.TextRange.Paragraphs.Add("");

            ITextRange roundedRectangular_p4 = roundedRectangular.TextFrame.TextRange.Paragraphs.Add();
            roundedRectangular_p4.Runs.Add("To delete this note, click the edge to select it and then press ");
            roundedRectangular_p4.Runs.Add("Delete");
            roundedRectangular_p4.Runs.Add(".");
            roundedRectangular_p4.Runs[2].Font.Bold = true;

            roundedRectangular.TextFrame.TextRange.Paragraphs.Add("");

            //Add Stright Line Shape
            IShape lineShape = worksheet.Shapes.AddConnector(ConnectorType.Straight, 455.228f, 57.35f, 756.228f, 57.35f);
            lineShape.Line.Solid();
            lineShape.Line.Weight = 3;
            lineShape.Line.Color.RGB = Color.FromArgb(38, 38, 38);
            lineShape.Line.DashStyle = LineDashStyle.SysDot;
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override string[] UsedResources
        {
            get
            {
                return new string[] { "movie.png", "list.png" };
            }
        }
    }
}
