using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Theme
{
    public class AddCustomThemes : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            //Base theme is office theme when not give parameter.
            ITheme theme = new GrapeCity.Documents.Excel.Theme("testTheme");
            //ITheme theme = new GrapeCity.Documents.Excel.Theme("testTheme", Themes.Badge);

            //Customize theme's color.
            theme.ThemeColorScheme[ThemeColor.Light1].RGB = Color.AntiqueWhite;
            theme.ThemeColorScheme[ThemeColor.Dark1].RGB = Color.Aqua;
            theme.ThemeColorScheme[ThemeColor.Light2].RGB = Color.Beige;
            theme.ThemeColorScheme[ThemeColor.Dark1].RGB = Color.Black;
            theme.ThemeColorScheme[ThemeColor.Accent1].RGB = Color.CadetBlue;
            theme.ThemeColorScheme[ThemeColor.Accent2].RGB = Color.Chartreuse;
            theme.ThemeColorScheme[ThemeColor.Accent3].RGB = Color.Chocolate;
            theme.ThemeColorScheme[ThemeColor.Accent4].RGB = Color.Coral;
            theme.ThemeColorScheme[ThemeColor.Accent5].RGB = Color.CornflowerBlue;
            theme.ThemeColorScheme[ThemeColor.Accent6].RGB = Color.Cornsilk;
            theme.ThemeColorScheme[ThemeColor.Hyperlink].RGB = Color.Honeydew;
            theme.ThemeColorScheme[ThemeColor.FollowedHyperlink].RGB = Color.HotPink;

            //Customize theme's font.
            theme.ThemeFontScheme.Major[FontLanguageIndex.Latin].Name = "Kristen ITC";
            theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Segoe Script";

            //Change workbook's theme to custom theme.
            workbook.Theme = theme;

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["B2"].Value = "Major Font:";
            worksheet.Range["B3"].Value = "Minor Font:";
            worksheet.Range["C2"].Value = "Kristen ITC";
            worksheet.Range["C3"].Value = "Segoe Script";
            worksheet.Range["C2"].Font.ThemeFont = ThemeFont.Major;
            worksheet.Range["C3"].Font.ThemeFont = ThemeFont.Minor;

            worksheet.Range["E2:E13"].Value = new object[]
            {
                "Light1",
                "Dark1",
                "Light2",
                "Dark2",
                "Accent1",
                "Accent2",
                "Accent3",
                "Accent4",
                "Accent5",
                "Accent6",
                "Hyperlink",
                "FollowedHyperlink"
            };

            worksheet.Range["F2"].Interior.ThemeColor = ThemeColor.Light1;
            worksheet.Range["F3"].Interior.ThemeColor = ThemeColor.Dark1;
            worksheet.Range["F4"].Interior.ThemeColor = ThemeColor.Light2;
            worksheet.Range["F5"].Interior.ThemeColor = ThemeColor.Dark2;
            worksheet.Range["F6"].Interior.ThemeColor = ThemeColor.Accent1;
            worksheet.Range["F7"].Interior.ThemeColor = ThemeColor.Accent2;
            worksheet.Range["F8"].Interior.ThemeColor = ThemeColor.Accent3;
            worksheet.Range["F9"].Interior.ThemeColor = ThemeColor.Accent4;
            worksheet.Range["F10"].Interior.ThemeColor = ThemeColor.Accent5;
            worksheet.Range["F11"].Interior.ThemeColor = ThemeColor.Accent6;
            worksheet.Range["F12"].Interior.ThemeColor = ThemeColor.Hyperlink;
            worksheet.Range["F13"].Interior.ThemeColor = ThemeColor.FollowedHyperlink;

        }

        public override bool IsUpdate
        {
            get
            {
                return true;
            }
        }
    }
}
