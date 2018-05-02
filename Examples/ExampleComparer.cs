using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples
{
    internal class ExampleComparer : IComparer<ExampleBase>
    {
        private Dictionary<string, string> _sortOrders = new Dictionary<string, string>();
        public ExampleComparer()
        {
            // root children orders
            _sortOrders.Add("Tutorial", "a");
            _sortOrders.Add("Features", "b");
            _sortOrders.Add("SpreadSheetsViewer", "c");
            _sortOrders.Add("ExcelReporting", "d");
            _sortOrders.Add("ExcelTemplates", "e");

            // Features children orders
            _sortOrders.Add("RangeOperations", "a");
            _sortOrders.Add("Formatting", "b");
            _sortOrders.Add("Tables", "c");
            _sortOrders.Add("ConditionalFormatting", "d");
            _sortOrders.Add("DataValidation", "e");
            _sortOrders.Add("Formulas", "f");
            _sortOrders.Add("Grouping", "g");
            _sortOrders.Add("Filtering", "h");
            _sortOrders.Add("Sorting", "i");
            _sortOrders.Add("Sparklines", "j");
            _sortOrders.Add("Charts", "k");
            _sortOrders.Add("Shape", "l");
            _sortOrders.Add("Picture", "m");
            _sortOrders.Add("Slicer", "n");
            _sortOrders.Add("Comments", "o");
            _sortOrders.Add("PivotTable", "p");
            _sortOrders.Add("Hyperlinks", "q");
            _sortOrders.Add("Theme", "r");
            _sortOrders.Add("Workbook", "s");
            _sortOrders.Add("Worksheets", "t");
        }

        public int Compare(ExampleBase x, ExampleBase y)
        {
            if(x is Tutorial)
            {
                return -1;
            }
            else if(y is Tutorial)
            {
                return 1;
            }

            string xSortKey;
            if (!_sortOrders.TryGetValue(x.GetShortID(), out xSortKey))
            {
                xSortKey = x.ID;
            }

            string ySortKey;
            if (!_sortOrders.TryGetValue(y.GetShortID(), out ySortKey))
            {
                ySortKey = y.ID;
            }

            if (x is FolderExample)
            {
                if (y is FolderExample)
                {
                    return xSortKey.CompareTo(ySortKey);
                }
                else
                {
                    return -1;
                }
            }
            else
            {
                if (y is FolderExample)
                {
                    return 1;
                }
                else
                {
                    return xSortKey.CompareTo(ySortKey);
                }
            }
        }
    }
}
