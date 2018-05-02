using System;
using System.Collections.Generic;
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
            theme.ThemeColorScheme[ThemeColor.Accent1].RGB = Color.AliceBlue;

            //Customize theme's font.
            theme.ThemeFontScheme.Major[FontLanguageIndex.Latin].Name = "Buxton Sketch";
            theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Segoe UI";

            //Change workbook's theme to custom theme.
            workbook.Theme = theme;

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["B2"].Value = "Text";

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
