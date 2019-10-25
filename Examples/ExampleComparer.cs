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
            _sortOrders.Add("Tutorial".ToLower(), "a");
            _sortOrders.Add("Features".ToLower(), "b");
            _sortOrders.Add("Showcase".ToLower(), "c");
            _sortOrders.Add("Templates".ToLower(), "d");
            _sortOrders.Add("SpreadSheetsViewer".ToLower(), "e");

            // Features children orders
            _sortOrders.Add("RangeOperations".ToLower(), "a");
            _sortOrders.Add("Formatting".ToLower(), "b");
            _sortOrders.Add("Tables".ToLower(), "c");
            _sortOrders.Add("ConditionalFormatting".ToLower(), "d");
            _sortOrders.Add("DataValidation".ToLower(), "e");
            _sortOrders.Add("Formulas".ToLower(), "f");
            _sortOrders.Add("Grouping".ToLower(), "g");
            _sortOrders.Add("Filtering".ToLower(), "h");
            _sortOrders.Add("Sorting".ToLower(), "i");
            _sortOrders.Add("Sparklines".ToLower(), "j");
            _sortOrders.Add("Charts".ToLower(), "k");
            _sortOrders.Add("Shape".ToLower(), "l");
            _sortOrders.Add("Picture".ToLower(), "m");
            _sortOrders.Add("Slicer".ToLower(), "n");
            _sortOrders.Add("Comments".ToLower(), "o");
            _sortOrders.Add("PivotTable".ToLower(), "p");
            _sortOrders.Add("Hyperlinks".ToLower(), "q");
            _sortOrders.Add("Theme".ToLower(), "r");
            _sortOrders.Add("Workbook".ToLower(), "s");
            _sortOrders.Add("Worksheets".ToLower(), "t");
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
